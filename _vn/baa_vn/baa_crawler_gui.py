import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import aiohttp
import asyncio
import pandas as pd
import re
import os
import random
import sys
from selectolax.parser import HTMLParser

try:
    import uvloop
    asyncio.set_event_loop_policy(uvloop.EventLoopPolicy())
except ImportError:
    pass

class BaaCrawlerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("BAA.VN Max Speed Crawler")
        self.root.geometry("600x420")

        # Mode Selection
        mode_frame = ttk.LabelFrame(root, text="Chế độ Crawl", padding=10)
        mode_frame.pack(fill="x", padx=10, pady=5)
        
        self.crawl_mode = tk.StringVar(value="excel")
        ttk.Radiobutton(mode_frame, text="Cào từ File Excel (Danh sách Model/Mã SP)", variable=self.crawl_mode, value="excel", command=self.toggle_mode).pack(anchor="w")
        ttk.Radiobutton(mode_frame, text="Cào từ Link Chuyên mục/Hãng BAA.VN", variable=self.crawl_mode, value="url", command=self.toggle_mode).pack(anchor="w")

        # Input Frame
        self.input_frame = ttk.LabelFrame(root, text="Đầu vào", padding=10)
        self.input_frame.pack(fill="x", padx=10, pady=5)
        
        # Excel Input
        self.excel_frame = ttk.Frame(self.input_frame)
        self.sku_file = tk.StringVar()
        ttk.Entry(self.excel_frame, textvariable=self.sku_file, width=50).pack(side="left", padx=5)
        ttk.Button(self.excel_frame, text="Chọn File", command=self.browse_file).pack(side="left")
        
        # URL Input
        self.url_frame = ttk.Frame(self.input_frame)
        self.target_url = tk.StringVar()
        ttk.Entry(self.url_frame, textvariable=self.target_url, width=50).pack(side="left", padx=5)
        ttk.Label(self.url_frame, text="(Vd: https://baa.vn/vn/Category/...)").pack(side="left")

        # Progress Frame
        progress_frame = ttk.LabelFrame(root, text="Tiến trình", padding=10)
        progress_frame.pack(fill="x", padx=10, pady=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=5, pady=5)
        
        self.status_label = ttk.Label(progress_frame, text="Sẵn sàng")
        self.status_label.pack()

        # Controls
        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10)
        self.start_btn = ttk.Button(btn_frame, text="Bắt đầu Crawl Siêu Tốc", command=self.start_crawl)
        self.start_btn.pack(side="left", padx=10)
        
        self.toggle_mode()

    def toggle_mode(self):
        if self.crawl_mode.get() == "excel":
            self.url_frame.pack_forget()
            self.excel_frame.pack(fill="x")
        else:
            self.excel_frame.pack_forget()
            self.url_frame.pack(fill="x")

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.sku_file.set(filename)

    def start_crawl(self):
        if self.crawl_mode.get() == "excel" and not self.sku_file.get():
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Excel!")
            return
        if self.crawl_mode.get() == "url" and not self.target_url.get().startswith("http"):
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập Link hợp lệ (bắt đầu bằng http)!")
            return
            
        self.start_btn.config(state="disabled")
        self.progress_var.set(0)
        self.status_label.config(text="Khởi động engine cào Bảo An Automation...")
        
        threading.Thread(target=self.run_async_crawl, daemon=True).start()

    def run_async_crawl(self):
        asyncio.run(self.crawl_process())

    async def fetch_url_with_retry(self, session, url, semaphore, max_retries=5):
        # Full browser-like headers
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'vi,en-US;q=0.7,en;q=0.3',
            'Connection': 'keep-alive'
        }
        async with semaphore:
            for attempt in range(max_retries):
                try:
                    # Random delay between 100ms and 300ms to avoid WAF block
                    await asyncio.sleep(random.uniform(0.1, 0.3))
                    
                    async with session.get(url, headers=headers, timeout=20) as r:
                        if r.status == 200:
                            return await r.text()
                        elif r.status in [429, 503]:
                            backoff = (2 ** attempt) + random.uniform(0.5, 1.5)
                            await asyncio.sleep(backoff)
                        else:
                            # Other status codes: wait slightly and retry
                            await asyncio.sleep(1)
                except Exception:
                    await asyncio.sleep(2 ** attempt)
            return None

    async def crawl_process(self):
        # Concurrency limit 8 to avoid triggering WAF blocking
        semaphore = asyncio.Semaphore(8)
        
        async with aiohttp.ClientSession() as session:
            try:
                if self.crawl_mode.get() == "excel":
                    await self.crawl_excel_mode(session, semaphore)
                else:
                    await self.crawl_url_mode(session, semaphore)
            except Exception as e:
                self.update_status(f"Lỗi: {str(e)}")
            finally:
                self.root.after(0, lambda: self.start_btn.config(state="normal"))

    def update_status(self, text, progress_val=None):
        self.root.after(0, lambda: self.status_label.config(text=text))
        if progress_val is not None:
            self.root.after(0, lambda: self.progress_var.set(progress_val))

    async def crawl_excel_mode(self, session, semaphore):
        file_path = self.sku_file.get()
        self.update_status("Đang đọc file Excel...")
        
        try:
            df_in = pd.read_excel(file_path)
            # Find the column containing model name
            col_name = None
            for col in df_in.columns:
                if any(kw in str(col).lower() for kw in ['model', 'mã', 'sku', 'product', 'sản phẩm']):
                    col_name = col
                    break
            if col_name is None:
                col_name = df_in.columns[0]
                
            skus = df_in[col_name].dropna().astype(str).str.strip().tolist()
            skus = list(set(skus)) # Deduplicate
        except Exception as e:
            self.update_status(f"Lỗi đọc file: {str(e)}")
            return

        total_skus = len(skus)
        self.update_status(f"Tìm thấy {total_skus} model cần tìm kiếm...")
        
        # Phase 1: Search and resolve product URLs
        resolved_urls = []
        for idx, sku in enumerate(skus):
            self.update_status(f"Đang tìm kiếm model ({idx+1}/{total_skus}): {sku}...", (idx / total_skus) * 30)
            search_url = f"https://baa.vn/vn/tim-kiem/?key={urllib.parse.quote(sku)}"
            
            html = await self.fetch_url_with_retry(session, search_url, semaphore)
            if html:
                parser = HTMLParser(html)
                # Check if it directly loaded a product page (e.g. H1 tag exists and contains the model)
                h1_el = parser.css_first('h1')
                if h1_el and sku.lower() in h1_el.text().lower():
                    # Direct match
                    resolved_urls.append(search_url)
                else:
                    # Search listing page, extract first matching product URL
                    prod_link = None
                    for a in parser.css('a'):
                        href = a.attributes.get('href', '')
                        if '/vn/san-pham/' in href:
                            if href.startswith('/'):
                                prod_link = 'https://baa.vn' + href
                            else:
                                prod_link = href
                            break
                    if prod_link:
                        resolved_urls.append(prod_link)
        
        resolved_urls = list(set(resolved_urls))
        total_urls = len(resolved_urls)
        self.update_status(f"Tìm thấy {total_urls} link sản phẩm tương thích. Bắt đầu cào chi tiết...")
        
        # Phase 2: Crawl detailed product data
        results = []
        completed = 0
        
        tasks = [self.parse_product_page(session, url, semaphore) for url in resolved_urls]
        for future in asyncio.as_completed(tasks):
            data = await future
            completed += 1
            if data:
                results.append(data)
            self.update_status(f"Đang cào: {completed}/{total_urls} sản phẩm...", 30 + (completed / total_urls) * 70)
            
        if results:
            self.save_to_excel(results, "Cào từ Excel BAA")
        else:
            self.update_status("Không lấy được dữ liệu chi tiết nào!")

    async def crawl_url_mode(self, session, semaphore):
        target = self.target_url.get()
        self.update_status("Giai đoạn 1: Đang quét danh sách trang để thu thập link...")
        
        # Fetch page 1
        html = await self.fetch_url_with_retry(session, target, semaphore)
        if not html:
            self.update_status("Không thể tải trang mục tiêu!")
            return
            
        parser = HTMLParser(html)
        
        # Extract title of category
        title_el = parser.css_first('title')
        cat_name = title_el.text().strip() if title_el else "Bao An"
        cat_name = re.sub(r'[\/\\:\*\?"<>\|]', '_', cat_name)[:50]
        
        # Detect pagination links dynamically
        page_urls = [target]
        max_page = 1
        for a in parser.css('a'):
            href = a.attributes.get('href', '')
            text = a.text().strip()
            # If href has page query
            m = re.search(r'[?&]page=(\d+)', href)
            if m:
                p_num = int(m.group(1))
                if p_num > max_page:
                    max_page = p_num
                    
        # Generate page URLs
        # Format: base_url?page=X (preserve existing params if any)
        base_url = target.split('?')[0]
        existing_params = target.split('?')[1] if '?' in target else ""
        
        for p in range(2, max_page + 1):
            if existing_params:
                p_url = f"{base_url}?{existing_params}&page={p}"
            else:
                p_url = f"{base_url}?page={p}"
            page_urls.append(p_url)
            
        self.update_status(f"Tìm thấy {max_page} trang. Đang thu thập link sản phẩm từ tất cả các trang...")
        
        # Crawl listing pages in parallel to gather product links
        product_urls = set()
        completed_pages = 0
        page_tasks = [self.fetch_url_with_retry(session, u, semaphore) for u in page_urls]
        for future in asyncio.as_completed(page_tasks):
            page_html = await future
            completed_pages += 1
            self.update_status(f"Quét trang: {completed_pages}/{max_page}...", (completed_pages / max_page) * 30)
            if page_html:
                p_parser = HTMLParser(page_html)
                for a in p_parser.css('a'):
                    href = a.attributes.get('href', '')
                    if '/vn/san-pham/' in href:
                        if href.startswith('/'):
                            href = 'https://baa.vn' + href
                        product_urls.add(href)
                        
        product_urls = list(product_urls)
        total_products = len(product_urls)
        self.update_status(f"Thu thập được {total_products} sản phẩm. Bắt đầu cào chi tiết...")
        
        # Crawl product details
        results = []
        completed = 0
        tasks = [self.parse_product_page(session, url, semaphore) for url in product_urls]
        for future in asyncio.as_completed(tasks):
            data = await future
            completed += 1
            if data:
                results.append(data)
            self.update_status(f"Đang cào: {completed}/{total_products} sản phẩm...", 30 + (completed / total_products) * 70)
            
        if results:
            self.save_to_excel(results, cat_name)
        else:
            self.update_status("Không lấy được dữ liệu chi tiết nào!")

    async def parse_product_page(self, session, url, semaphore):
        html = await self.fetch_url_with_retry(session, url, semaphore)
        if not html:
            return None
            
        parser = HTMLParser(html)
        
        # 1. Product Name (H1)
        h1_el = parser.css_first('h1')
        name = h1_el.text().strip() if h1_el else ""
        if not name:
            title_el = parser.css_first('title')
            name = title_el.text().strip() if title_el else ""
            
        # 2. Extract right-column key-values
        # Mã hàng, Thương hiệu, Series, Danh mục
        code = ""
        brand = ""
        series = ""
        category_name = ""
        
        # Loop through header attributes blocks
        for header_div in parser.css('.product__symbol-header'):
            text_content = header_div.text(separator=' ', strip=True)
            if 'Mã hàng:' in text_content:
                val_el = header_div.css_first('.product__symbol__value')
                if val_el: code = val_el.text(strip=True)
            elif 'Thương hiệu:' in text_content:
                val_el = header_div.css_first('.product__symbol__value')
                if val_el: brand = val_el.text(strip=True)
            elif 'Series:' in text_content:
                val_el = header_div.css_first('.product__symbol__value')
                if val_el: series = val_el.text(strip=True)
            elif 'Danh Mục:' in text_content:
                val_el = header_div.css_first('.product__symbol__value')
                if val_el: category_name = val_el.text(strip=True)

        # 3. Extract Breadcrumbs for category levels
        breadcrumbs = []
        for a in parser.css('.breadcrumb a, ol.breadcrumb li a, ul.breadcrumb li a'):
            txt = a.text().strip()
            if txt and txt not in breadcrumbs:
                breadcrumbs.append(txt)
                
        # 4. Extract Price (Giá chưa VAT, có VAT, VAT percent)
        # Search HTML for specific text
        # If there's an active numeric price, it will be in the format e.g. "1.200.000 VNĐ"
        # Otherwise it says "Liên hệ"
        price_no_vat = ""
        price_vat = ""
        currency = ""
        vat_percent = ""
        
        # Default price is Contact
        price_text = "Liên hệ"
        
        # Look for any price spans/divs or text labels
        raw_text = parser.body.text() if parser.body else html
        price_matches = re.findall(r'(\b\d{1,3}(?:\.\d{3})+\b)\s*(?:đ|VND|VNĐ)', raw_text, re.IGNORECASE)
        if price_matches:
            price_vat = price_matches[0].replace('.', '')
            currency = "VND"
            price_text = price_matches[0]
            
        # 5. Extract stock, warranty, origin
        # BAA.VN spec table or details block usually contains these fields
        warranty = ""
        origin = ""
        stock = ""
        
        # Search specifications table
        spec_rows = []
        table = parser.css_first('table')
        if table:
            for tr in table.css('tr'):
                cells = [c.text().strip() for c in tr.css('td, th')]
                if len(cells) >= 2:
                    lbl = cells[0].strip()
                    val = cells[1].strip()
                    spec_rows.append(f"{lbl}: {val}")
                    
                    lbl_lower = lbl.lower()
                    if 'bảo hành' in lbl_lower:
                        warranty = val
                    elif 'xuất xứ' in lbl_lower:
                        origin = val
                    elif 'tình trạng' in lbl_lower or 'tồn kho' in lbl_lower:
                        stock = val
                        
        specs_str = "\n".join(spec_rows)
        
        # If stock not found in table, look for badge or status text
        if not stock:
            # Common statuses: "Còn hàng", "Hết hàng", "Liên hệ"
            for status_lbl in ["Tình trạng:", "Tình trạng hàng:"]:
                idx = raw_text.find(status_lbl)
                if idx != -1:
                    snippet = raw_text[idx:idx+50]
                    # Extract status text after label
                    m = re.search(r'(?:tình trạng|tình trạng hàng):\s*([^\n\r]+)', snippet, re.IGNORECASE)
                    if m:
                        stock = m.group(1).strip()
                        break
        # Ensure stock contains a number if user wanted "bao nhiêu cái", else fallback to text
        if not stock:
            stock = "Liên hệ"

        # 6. Product image URL
        img_url = ""
        img_meta = parser.css_first('meta[property="og:image"]')
        if img_meta:
            img_url = img_meta.attributes.get('content', '').strip()
        else:
            img_el = parser.css_first('#imgProduct, .product-image img, img.main-image')
            if img_el:
                img_url = img_el.attributes.get('src', '').strip()
                if img_url.startswith('/'):
                    img_url = 'https://baa.vn' + img_url
                    
        # Structure the data
        data = {
            "Tên sản phẩm": name,
            "Mã hàng": code,
            "Thương hiệu": brand,
            "Series": series,
            "Danh mục": category_name or (breadcrumbs[-1] if breadcrumbs else ""),
            "Đơn giá": price_text,
            "Đơn vị": currency or "VND",
            "Tình trạng": stock,
            "Bảo hành": warranty or "Theo tiêu chuẩn hãng",
            "Xuất xứ": origin or "Không xác định",
            "Ảnh sản phẩm": img_url,
            "Thông số kỹ thuật": specs_str,
            "Link": url
        }
        
        # Split categories dynamically (Phân loại 1 -> 5)
        for i in range(5):
            val = breadcrumbs[i] if i < len(breadcrumbs) else ""
            data[f"Phân loại {i+1}"] = val
            
        return data

    def save_to_excel(self, results, name_prefix):
        df = pd.DataFrame(results)
        
        # Rearrange columns so Phân loại 1 -> 5 are grouped
        main_cols = [
            "Tên sản phẩm", "Mã hàng", "Thương hiệu", "Series", "Danh mục",
            "Phân loại 1", "Phân loại 2", "Phân loại 3", "Phân loại 4", "Phân loại 5",
            "Đơn giá", "Đơn vị", "Tình trạng", "Bảo hành", "Xuất xứ", "Ảnh sản phẩm", "Thông số kỹ thuật", "Link"
        ]
        
        # Keep only existing columns
        cols = [c for c in main_cols if c in df.columns]
        df = df[cols]
        
        # File Name
        file_name = f"{name_prefix}_result.xlsx"
        self.update_status(f"Đang ghi dữ liệu vào file {file_name}...")
        try:
            df.to_excel(file_name, index=False)
            self.update_status(f"Hoàn thành! Kết quả lưu tại: {os.path.abspath(file_name)}", 100)
            messagebox.showinfo("Thành công", f"Đã cào xong dữ liệu!\nLưu tại: {file_name}")
        except Exception as e:
            self.update_status(f"Lỗi ghi file: {str(e)}")
            messagebox.showerror("Lỗi", f"Không thể lưu file Excel: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = BaaCrawlerGUI(root)
    root.mainloop()
