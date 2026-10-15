import os
import sys
import time
import asyncio
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import httpx
from selectolax.parser import HTMLParser
import gc
import logging
from logging.handlers import RotatingFileHandler

# Setup Logger with Rotation
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_file = 'crawler_errors.log'
my_handler = RotatingFileHandler(log_file, mode='a', maxBytes=5*1024*1024, backupCount=2, encoding='utf-8', delay=0)
my_handler.setFormatter(log_formatter)
logger = logging.getLogger('root')
logger.setLevel(logging.INFO)
logger.addHandler(my_handler)

# Suppress httpx spam
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)

class AsyncCrawler:
    def __init__(self, mode, excel_file, target_col, keyword, get_html, gui_update_callback):
        self.mode = mode
        self.excel_file = excel_file
        self.target_col = target_col
        self.keyword = keyword
        self.get_html = get_html
        self.gui_update = gui_update_callback
        self.cancel_requested = False
        
        self.results = []
        self.aliases = set()
        
        # E2E Features Tracking
        self.processed_count = 0
        self.consecutive_429_errors = 0
        self.is_hibernating = False
        self.failed_aliases = []
        
        self.temp_csv = 'temp_data.csv'
        self.failed_txt = 'failed_sku.txt'
        
        # Clean temp on start
        if os.path.exists(self.temp_csv):
            try: os.remove(self.temp_csv)
            except: pass
        if os.path.exists(self.failed_txt):
            try: os.remove(self.failed_txt)
            except: pass
        
    def log(self, msg):
        self.gui_update("log", msg)
        logger.info(msg)
        
    def progress(self, current, total, errors):
        self.gui_update("progress", (current, total, errors))
        
    def auto_save_and_gc(self):
        # Append to temp CSV and CLEAR RAM
        try:
            if self.results:
                df = pd.DataFrame(self.results)
                # Append to CSV
                header = not os.path.exists(self.temp_csv)
                df.to_csv(self.temp_csv, mode='a', header=header, index=False, encoding='utf-8')
                # CLEAR RAM completely
                self.results.clear()
            gc.collect()
        except Exception as e:
            logger.error(f"Auto-save failed: {e}")

    async def fetch_page_aliases(self, client, page):
        if self.cancel_requested: return []
        url = f"https://nguyenkhue.com/collections/all?page={page}"
        for attempt in range(4):
            try:
                r = await client.get(url, follow_redirects=False) # Prevent redirect loops
                if r.status_code == 200:
                    self.consecutive_429_errors = 0
                    tree = HTMLParser(r.text)
                    links = tree.css(".product-inner a, .product-loop a")
                    found = []
                    for link in links:
                        href = link.attributes.get('href')
                        if href and '/products/' in href:
                            alias = href.split('/products/')[-1].split('?')[0].split('#')[0]
                            found.append(alias)
                    return found
                elif r.status_code in [301, 302]: # Redirect usually means out of bounds
                    return []
                elif r.status_code == 429:
                    self.consecutive_429_errors += 1
                    logger.warning(f"[Page {page}] 429 Rate Limit. Attempt {attempt}")
                    await self.handle_429_backoff(attempt)
            except Exception as e:
                logger.error(f"[Page {page}] Fetch Error: {e}")
                await asyncio.sleep(2)
        return []

    async def handle_429_backoff(self, attempt):
        if self.consecutive_429_errors >= 10:
            if not self.is_hibernating:
                self.is_hibernating = True
                self.log("Bị chặn Rate Limit nặng. Hibernate 120s...")
                await asyncio.sleep(120)
                self.consecutive_429_errors = 0
                self.is_hibernating = False
            else:
                while self.is_hibernating:
                    await asyncio.sleep(5)
        else:
            await asyncio.sleep(5 * (attempt + 1))

    async def fetch_search_aliases(self, client, query):
        if self.cancel_requested: return []
        url = f"https://nguyenkhue.com/search?type=product&q={query}"
        for attempt in range(4):
            try:
                r = await client.get(url, follow_redirects=True)
                if r.status_code == 200:
                    self.consecutive_429_errors = 0
                    tree = HTMLParser(r.text)
                    links = tree.css(".product-inner a, .product-loop a")
                    found = []
                    for link in links:
                        href = link.attributes.get('href')
                        if href and '/products/' in href:
                            alias = href.split('/products/')[-1].split('?')[0].split('#')[0]
                            found.append(alias)
                    return found
                elif r.status_code == 429:
                    self.consecutive_429_errors += 1
                    await self.handle_429_backoff(attempt)
            except Exception as e:
                logger.error(f"[Search {query}] Fetch Error: {e}")
                await asyncio.sleep(2)
        return []

    async def fetch_product_details(self, client, alias, semaphore):
        if self.cancel_requested: return False
        async with semaphore:
            await asyncio.sleep(0.5) # Pace out requests
            json_url = f"https://nguyenkhue.com/products/{alias}.js"
            for attempt in range(8): # Increase retries to 8
                try:
                    r = await client.get(json_url, follow_redirects=True)
                    if r.status_code == 200:
                        self.consecutive_429_errors = 0
                        data = r.json()
                        name = data.get('title') or data.get('name', '')
                        
                        variants = data.get('variants', [{}])
                        img_url = data.get('featured_image', '')
                        if img_url and img_url.startswith('//'):
                            img_url = 'https:' + img_url
                            
                        raw_desc = data.get('content', '')
                        
                        # Parse Specs Table
                        specs = []
                        desc = raw_desc
                        if raw_desc:
                            tree = HTMLParser(raw_desc)
                            tables = tree.css("table")
                            if tables:
                                for tr in tables[0].css("tr"):
                                    cols = tr.css("td, th")
                                    if len(cols) >= 2:
                                        k = cols[0].text(strip=True)
                                        v = cols[1].text(strip=True)
                                        if k and v:
                                            specs.append(f"{k}: {v}")
                            if not self.get_html:
                                desc = tree.text(separator=' ', strip=True)
                                
                        specs_str = "\n".join(specs)
                        opt_keys = [opt.get('name') for opt in data.get('options', [])]
                        
                        for var in variants:
                            sku = var.get('sku', '')
                            
                            # Fallback nếu Nguyễn Khuê không nhập SKU vào ô cấu hình
                            if not sku:
                                # Lấy từ cuối cùng của tên sản phẩm (thường là mã)
                                last_word = name.strip().split()[-1] if name else ''
                                # Nếu từ cuối cùng có chứa số và chữ, khả năng cao là mã SKU
                                if any(c.isdigit() for c in last_word) and len(last_word) >= 3:
                                    sku = last_word
                                else:
                                    # Fallback 2: Lấy từ URL alias
                                    alias_parts = alias.split('-')
                                    if alias_parts:
                                        sku = alias_parts[-1].upper()
                                        
                            price = var.get('price', 0)
                            compare_price = var.get('compare_at_price', price)
                            
                            try:
                                price_str = f"{int(price):,}".replace(',', '.') if price else "0"
                            except: price_str = str(price)
                            
                            try:
                                comp_price_str = f"{int(compare_price):,}".replace(',', '.') if compare_price else "0"
                            except: comp_price_str = str(compare_price)
                            
                            var_opts = var.get('options', [])
                            var_details = {}
                            for idx, ok in enumerate(opt_keys):
                                if idx < len(var_opts) and ok != "Tiêu đề":
                                    var_details[ok] = var_opts[idx]
                                    
                            row = {
                                'SKU': sku,
                                'Name': name,
                                'Variant_Name': var.get('title', ''),
                                'Price': price_str,
                                'Compare_Price': comp_price_str,
                                'Image': img_url,
                                'URL': f"https://nguyenkhue.com/products/{alias}",
                                'Spec': specs_str,
                                'Description': desc
                            }
                            row.update(var_details)
                            self.results.append(row)
                        return True
                    elif r.status_code == 429:
                        self.consecutive_429_errors += 1
                        await self.handle_429_backoff(attempt)
                    else:
                        await asyncio.sleep(3)
                except Exception as e:
                    logger.error(f"[Product {alias}] Exception: {e}")
                    await asyncio.sleep(3)
            
            # If all attempts fail, log it
            self.failed_aliases.append(alias)
            with open(self.failed_txt, 'a', encoding='utf-8') as f:
                f.write(f"{alias}\n")
            return False

    async def run_crawl_all(self, client):
        self.log("Phase 1: Quét danh mục (Fast Async)...")
        page = 1
        last_aliases_count = 0
        
        while not self.cancel_requested:
            self.log(f"Đang quét trang {page} - {page+9}...")
            # Fetch 10 pages concurrently
            tasks = [self.fetch_page_aliases(client, p) for p in range(page, page+10)]
            results = await asyncio.gather(*tasks)
            
            found_any = False
            for page_aliases in results:
                if page_aliases:
                    found_any = True
                    for a in page_aliases:
                        self.aliases.add(a)
                        
            # Infinite Loop Prevention: If count stops growing, break
            if len(self.aliases) == last_aliases_count:
                break
            last_aliases_count = len(self.aliases)
            
            if not found_any:
                break
                
            page += 10
            self.progress(0, len(self.aliases), 0)

    async def run_crawl_excel(self, client):
        self.log("Phase 1: Đọc file Excel và tìm kiếm SKU...")
        try:
            df = pd.read_excel(self.excel_file, dtype=str)
            if self.target_col not in df.columns:
                self.gui_update("error", f"Không tìm thấy cột '{self.target_col}' trong Excel!")
                return
                
            skus = df[self.target_col].dropna().unique().tolist()
            total_skus = len(skus)
            self.log(f"Tìm thấy {total_skus} SKU cần tra cứu. Đang search...")
            
            # Concurrently search max 3 at a time
            sem = asyncio.Semaphore(3)
            completed = 0
            
            async def safe_search(sku):
                nonlocal completed
                async with sem:
                    await asyncio.sleep(0.5)
                    res = await self.fetch_search_aliases(client, sku)
                    completed += 1
                    # Update progress every 10 SKUs to avoid GUI spam
                    if completed % 10 == 0 or completed == total_skus:
                        self.gui_update("search_progress", (completed, total_skus, len(self.aliases)))
                    return res
                    
            tasks = [safe_search(sku) for sku in skus]
            results = await asyncio.gather(*tasks)
            
            for res in results:
                for a in res:
                    self.aliases.add(a)
        except Exception as e:
            logger.error(f"Excel read error: {e}")
            self.gui_update("error", f"Lỗi đọc Excel: {e}")

    async def run_crawl_keyword(self, client):
        self.log(f"Phase 1: Tìm kiếm từ khóa '{self.keyword}'...")
        aliases = await self.fetch_search_aliases(client, self.keyword)
        for a in aliases:
            self.aliases.add(a)

    async def execute(self):
        # Strict timeout for Deadlock prevention
        timeout = httpx.Timeout(30.0, connect=10.0)
        limits = httpx.Limits(max_connections=5, max_keepalive_connections=3)
        
        async with httpx.AsyncClient(timeout=timeout, limits=limits, headers={'User-Agent': 'Mozilla/5.0'}) as client:
            if self.mode == "ALL":
                await self.run_crawl_all(client)
            elif self.mode == "EXCEL":
                await self.run_crawl_excel(client)
            elif self.mode == "KEYWORD":
                await self.run_crawl_keyword(client)
                
            if self.cancel_requested:
                self.gui_update("stopped", None)
                return
                
            aliases_list = list(self.aliases)
            total = len(aliases_list)
            self.log(f"Phase 2: Tìm thấy {total} sản phẩm. Bắt đầu lấy dữ liệu chi tiết...")
            
            if total == 0:
                self.gui_update("error", "Không tìm thấy sản phẩm nào để lấy dữ liệu (hoặc bị block 429 Rate Limit)!")
                return
                
            sem = asyncio.Semaphore(3) # Reduced to 3 to avoid Haravan 429 block
            tasks = []
            
            # Progress tracking variables
            completed = 0
            errors = 0
            
            async def wrapped_fetch(alias):
                nonlocal completed, errors
                success = await self.fetch_product_details(client, alias, sem)
                completed += 1
                self.processed_count += 1
                if not success: errors += 1
                
                # Update progress every 5 items
                if completed % 5 == 0 or completed == total:
                    self.progress(completed, total, errors)
                
                # Auto save and GC every 50 items
                if self.processed_count >= 50:
                    self.processed_count = 0
                    self.auto_save_and_gc()
            
            for alias in aliases_list:
                tasks.append(asyncio.create_task(wrapped_fetch(alias)))
                
            await asyncio.gather(*tasks)
            
            # Lấy data còn sót lại trên RAM
            final_df = pd.DataFrame(self.results) if self.results else pd.DataFrame()
            
            # Gộp với data đã dump ra CSV tạm
            if os.path.exists(self.temp_csv):
                try:
                    temp_df = pd.read_csv(self.temp_csv, dtype=str)
                    final_df = pd.concat([temp_df, final_df], ignore_index=True)
                except Exception as e:
                    logger.error(f"Error reading temp_csv: {e}")
            
            if not final_df.empty:
                self.log("Đang lưu file Excel chính thức...")
                output_file = "nguyenkhue_data_optimized.xlsx"
                try:
                    final_df.to_excel(output_file, index=False)
                    # Clean up temp file on success
                    if os.path.exists(self.temp_csv):
                        os.remove(self.temp_csv)
                        
                    if self.cancel_requested:
                        self.gui_update("stopped", output_file)
                    else:
                        self.gui_update("done", output_file)
                except Exception as e:
                    logger.error(f"Excel save error: {e}")
                    self.gui_update("error", f"Lỗi khi lưu file: {e}")
            else:
                self.gui_update("error", "Không có dữ liệu hợp lệ để lưu!")


class NguyenKhueGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Nguyen Khue Crawler V2.0 - Doanh Nghiệp (E2E)")
        self.root.geometry("650x650")
        
        self.crawler = None
        self.is_running = False
        self.excel_path = ""
        self.columns = []
        
        self.setup_ui()
        
    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="⚡ CÔNG CỤ CRAWL DATA NGUYENKHUE.COM (E2E)", font=('Arial', 12, 'bold')).pack(pady=(0, 15))
        
        # --- Mode Selection ---
        mode_frame = ttk.LabelFrame(main_frame, text="1. Chế độ hoạt động", padding="10")
        mode_frame.pack(fill=tk.X, pady=5)
        
        self.mode_var = tk.StringVar(value="ALL")
        
        r1 = ttk.Radiobutton(mode_frame, text="Crawl toàn bộ Web", variable=self.mode_var, value="ALL", command=self.update_ui_state)
        r1.pack(anchor=tk.W, pady=2)
        
        r2 = ttk.Radiobutton(mode_frame, text="Tìm theo từ khóa", variable=self.mode_var, value="KEYWORD", command=self.update_ui_state)
        r2.pack(anchor=tk.W, pady=2)
        
        self.entry_keyword = ttk.Entry(mode_frame, width=40, state=tk.DISABLED)
        self.entry_keyword.pack(anchor=tk.W, padx=20, pady=2)
        
        r3 = ttk.Radiobutton(mode_frame, text="Tra cứu từ File Excel", variable=self.mode_var, value="EXCEL", command=self.update_ui_state)
        r3.pack(anchor=tk.W, pady=2)
        
        excel_frame = ttk.Frame(mode_frame)
        excel_frame.pack(fill=tk.X, padx=20, pady=2)
        
        self.btn_browse = ttk.Button(excel_frame, text="Chọn File", command=self.browse_file, state=tk.DISABLED)
        self.btn_browse.pack(side=tk.LEFT)
        
        self.lbl_file = ttk.Label(excel_frame, text="Chưa chọn file")
        self.lbl_file.pack(side=tk.LEFT, padx=10)
        
        col_frame = ttk.Frame(mode_frame)
        col_frame.pack(fill=tk.X, padx=20, pady=2)
        
        ttk.Label(col_frame, text="Chọn cột chứa Mã (SKU):").pack(side=tk.LEFT)
        self.cb_columns = ttk.Combobox(col_frame, state=tk.DISABLED, width=20)
        self.cb_columns.pack(side=tk.LEFT, padx=5)
        
        # --- Options ---
        opt_frame = ttk.LabelFrame(main_frame, text="2. Tùy chọn", padding="10")
        opt_frame.pack(fill=tk.X, pady=10)
        
        self.var_get_html = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_frame, text="Giữ nguyên mã HTML ở cột mô tả (Làm file Excel nặng hơn)", variable=self.var_get_html).pack(anchor=tk.W)
        
        # --- Progress ---
        prog_frame = ttk.LabelFrame(main_frame, text="3. Tiến độ (Auto-Save 50 items/lần)", padding="10")
        prog_frame.pack(fill=tk.X, pady=10)
        
        self.lbl_status = ttk.Label(prog_frame, text="Trạng thái: Sẵn sàng")
        self.lbl_status.pack(anchor=tk.W)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(prog_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        self.lbl_stats = ttk.Label(prog_frame, text="Sản phẩm: 0/0 | Lỗi: 0")
        self.lbl_stats.pack(anchor=tk.W)
        
        # --- Buttons ---
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        self.btn_start = ttk.Button(btn_frame, text="🚀 BẮT ĐẦU CRAWL", command=self.start_crawl)
        self.btn_start.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.btn_stop = ttk.Button(btn_frame, text="⛔ DỪNG LẠI", command=self.stop_crawl, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
    def update_ui_state(self):
        mode = self.mode_var.get()
        if mode == "ALL":
            self.entry_keyword.config(state=tk.DISABLED)
            self.btn_browse.config(state=tk.DISABLED)
            self.cb_columns.config(state=tk.DISABLED)
        elif mode == "KEYWORD":
            self.entry_keyword.config(state=tk.NORMAL)
            self.btn_browse.config(state=tk.DISABLED)
            self.cb_columns.config(state=tk.DISABLED)
        elif mode == "EXCEL":
            self.entry_keyword.config(state=tk.DISABLED)
            self.btn_browse.config(state=tk.NORMAL)
            if self.columns:
                self.cb_columns.config(state="readonly")
                
    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.excel_path = file_path
            self.lbl_file.config(text=os.path.basename(file_path))
            try:
                df = pd.read_excel(file_path, nrows=0)
                self.columns = list(df.columns)
                self.cb_columns.config(values=self.columns)
                self.cb_columns.config(state="readonly")
                
                # Auto select column if it has SKU or Mã
                for i, col in enumerate(self.columns):
                    if "sku" in col.lower() or "mã" in col.lower() or "code" in col.lower():
                        self.cb_columns.current(i)
                        break
                else:
                    if self.columns: self.cb_columns.current(0)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file Excel: {e}")

    def gui_update_callback(self, msg_type, data):
        # Must be safe to call from another thread
        self.root.after(0, self._process_update, msg_type, data)
        
    def _process_update(self, msg_type, data):
        if msg_type == "log":
            self.lbl_status.config(text=f"Trạng thái: {data}")
        elif msg_type == "search_progress":
            completed, total, found = data
            if total > 0:
                self.progress_var.set((completed / total) * 100)
            self.lbl_status.config(text=f"Trạng thái: Đang tìm kiếm sản phẩm trên web...")
            self.lbl_stats.config(text=f"Đang tìm SKU: {completed}/{total} | Đã tìm thấy: {found} sản phẩm")
        elif msg_type == "progress":
            current, total, errors = data
            if total > 0:
                self.progress_var.set((current / total) * 100)
            self.lbl_status.config(text=f"Trạng thái: Đang lấy chi tiết sản phẩm...")
            self.lbl_stats.config(text=f"Sản phẩm: {current}/{total} | Lỗi: {errors}")
        elif msg_type == "done":
            self.lbl_status.config(text="Trạng thái: HOÀN THÀNH!")
            self.reset_buttons()
            msg = f"Đã lưu kết quả tại:\n{data}"
            if os.path.exists('failed_sku.txt'):
                msg += "\n\nCảnh báo: Có một số sản phẩm bị lỗi, đã lưu danh sách vào failed_sku.txt"
            messagebox.showinfo("Thành công", msg)
        elif msg_type == "error":
            self.lbl_status.config(text="Trạng thái: CÓ LỖI!")
            self.reset_buttons()
            messagebox.showerror("Lỗi", data)
        elif msg_type == "stopped":
            self.lbl_status.config(text="Trạng thái: ĐĐA DỪNG!")
            self.reset_buttons()
            if data:
                messagebox.showinfo("Đã dừng", f"Đã lưu kết quả bị gián đoạn tại:\n{data}")

    def reset_buttons(self):
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self.is_running = False

    def start_crawl(self):
        if self.is_running: return
        mode = self.mode_var.get()
        keyword = self.entry_keyword.get().strip()
        target_col = self.cb_columns.get()
        
        if mode == "KEYWORD" and not keyword:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập từ khóa!")
            return
        if mode == "EXCEL" and not self.excel_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Excel!")
            return
        if mode == "EXCEL" and not target_col:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cột chứa mã sản phẩm!")
            return
            
        self.is_running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.progress_var.set(0)
        self.lbl_stats.config(text="Sản phẩm: 0/0 | Lỗi: 0")
        
        self.crawler = AsyncCrawler(
            mode=mode,
            excel_file=self.excel_path,
            target_col=target_col,
            keyword=keyword,
            get_html=self.var_get_html.get(),
            gui_update_callback=self.gui_update_callback
        )
        
        # Start async loop in a separate thread
        threading.Thread(target=self.run_async_loop, daemon=True).start()
        
    def run_async_loop(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(self.crawler.execute())
        except Exception as e:
            logger.error(f"System crash: {e}")
            self.gui_update_callback("error", f"Lỗi hệ thống:\n{str(e)}")
        finally:
            loop.close()

    def stop_crawl(self):
        if self.crawler:
            self.crawler.cancel_requested = True
            self.gui_update_callback("log", "Đang tiến hành dừng và lưu dữ liệu...")
            self.btn_stop.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = NguyenKhueGUI(root)
    root.mainloop()
