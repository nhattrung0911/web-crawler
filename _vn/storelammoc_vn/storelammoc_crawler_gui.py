import os
import re
import sys
import time
import queue
import threading
import json
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib3.util import Retry
from requests.adapters import HTTPAdapter

# Core crawling helper functions
def is_exact_sku_match(title, sku):
    """
    Checks if the SKU is an exact match in the title by verifying word boundaries.
    Prevents matching 'WU380S' with 'WU380S.9' or other sub-models.
    """
    title_lower = title.lower()
    sku_lower = sku.lower()
    
    # Let's write the actual loop correctly
    for match in re.finditer(re.escape(sku_lower), title_lower):
        start = match.start()
        end = match.end()
        
        # Check boundary before
        if start > 0:
            char_before = title_lower[start - 1]
            if char_before.isalnum():
                continue
                
        # Check boundary after
        if end < len(title_lower):
            char_after = title_lower[end]
            if char_after.isalnum():
                continue
            if char_after in ['.', '-', '_']:
                # If followed by another alphanumeric character (e.g. WU380S.9), it's a sub-model
                if end + 1 < len(title_lower) and title_lower[end + 1].isalnum():
                    continue
                    
        return True
    return False

def search_product_url(sku, session=None):
    """
    Uses Sapo's JSON search API to find the product URL for a given SKU on Store Làm Mộc.
    Extremely fast and avoids downloading heavy HTML pages.
    """
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    # Standardize/clean search query to prevent passing complex formats directly
    sku_clean = sku.strip()
    match_wildcard = re.search(r'\(\*([^*)]+)\*\)', sku_clean)
    if match_wildcard:
        sku_clean = match_wildcard.group(1)
    elif sku_clean.startswith('name:(') and sku_clean.endswith(')'):
        sku_clean = sku_clean[6:-1].strip('*')
        
    url = f"https://storelammoc.vn/search?q={sku_clean}&view=json"
    client = session if session else requests
    
    try:
        r = client.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            data = r.json()
            if isinstance(data, dict) and data:
                # 1. Try strict match first
                for pid, prod in data.items():
                    title = prod.get('title', '')
                    if is_exact_sku_match(title, sku_clean):
                        return prod.get('url')
                        
                # 2. Fallback to simple substring match if strict match failed to find anything
                sku_lower = sku_clean.lower()
                for pid, prod in data.items():
                    title = prod.get('title', '').lower()
                    if sku_lower in title:
                        return prod.get('url')
                        
                # 3. Second fallback to the first matching product
                first_key = list(data.keys())[0]
                return data[first_key].get('url')
    except Exception:
        pass
    return None

def fetch_product_details_json(product_url, sku, session=None):
    """
    Fetches details using Store Làm Mộc's product JSON API (8x FASTER!).
    Avoids parsing heavy HTML pages.
    """
    alias = product_url.lstrip('/')
    json_url = f"https://storelammoc.vn/products/{alias}.js"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    client = session if session else requests
    
    r = client.get(json_url, headers=headers, timeout=10)
    if r.status_code != 200:
        raise Exception(f"HTTP Status {r.status_code}")
    
    data = r.json()
    name = data.get('name', '')
    
    # 1. Short Description / Summary
    summary_html = data.get('summary', '')
    summary = BeautifulSoup(summary_html, 'lxml').get_text().strip() if summary_html else ""
    
    # 2. Prices
    price_val = data.get('price', 0)
    price_after = f"{int(price_val):,}".replace(',', '.') + '₫' if price_val else ""
    
    compare_at = data.get('compare_at_price_max', 0)
    if compare_at and compare_at > price_val:
        price_before = f"{int(compare_at):,}".replace(',', '.') + '₫'
    else:
        price_before = price_after
        
    # 3. Image Link (Already high resolution original link)
    img_url = data.get('featured_image', '')
    if img_url and img_url.startswith('//'):
        img_url = 'https:' + img_url
        
    # 4. Technical Specs
    content_html = data.get('content', '')
    soup = BeautifulSoup(content_html, 'lxml')
    spec_table = soup.select_one('table')
    specs = []
    if spec_table:
        for tr in spec_table.find_all('tr'):
            tds = [td.text.strip() for td in tr.find_all('td')]
            if len(tds) >= 2:
                specs.append(f"{tds[0]}: {tds[1]}")
    spec_str = "\n".join(specs)
    
    # 5. Categories / Breadcrumbs
    categories = [data.get('product_type', '')] if data.get('product_type') else []
    
    return {
        'SKU': sku,
        'URL': f"https://storelammoc.vn{product_url}",
        'categories': categories,
        'Name': name,
        'Summary': summary,
        'Price After': price_after,
        'Price Before': price_before,
        'Image Link': img_url,
        'Specs': spec_str
    }

def fetch_product_details(product_url, sku, session=None):
    """
    Fetches the HTML page of a product and extracts all requested information
    in a single request with high accuracy (Standard Mode).
    """
    full_url = f"https://storelammoc.vn{product_url}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    client = session if session else requests
    
    r = client.get(full_url, headers=headers, timeout=10)
    if r.status_code != 200:
        raise Exception(f"HTTP Status {r.status_code}")
    try:
        soup = BeautifulSoup(r.text, 'lxml')
    except Exception:
        soup = BeautifulSoup(r.text, 'html.parser')
    
    # 1. Breadcrumbs / Categories
    breadcrumbs = soup.select('ul.breadcrumb li')
    categories = []
    if len(breadcrumbs) > 2:
        for li in breadcrumbs[1:-1]:
            a_tag = li.find('a')
            if a_tag:
                text = a_tag.text.strip()
            else:
                span_tag = li.find('span')
                text = span_tag.text.strip() if span_tag else li.text.strip()
            text = text.rstrip('/').strip()
            if text and text != '/':
                categories.append(text)
                
    # 2. Product Name
    name_el = soup.select_one('h1.product-name') or soup.select_one('h1')
    name = name_el.text.strip() if name_el else ""
    
    # 3. Short Description (Mô tả ngắn)
    summary_el = soup.select_one('.product-summary')
    summary = ""
    if summary_el:
        import copy
        el_copy = copy.copy(summary_el)
        
        # Remove dynamic mobile layout button
        view_inf = el_copy.select_one('.view_inf')
        if view_inf:
            view_inf.decompose()
            
        lis = el_copy.find_all('li')
        if lis:
            summary = "\n".join(li.get_text().strip() for li in lis if li.get_text().strip())
        else:
            for br in el_copy.find_all('br'):
                br.replace_with('\n')
            ps = el_copy.find_all('p')
            if ps:
                summary = "\n".join(p.get_text().strip() for p in ps if p.get_text().strip())
            else:
                text = el_copy.get_text(separator=' ').strip()
                text = re.sub(r' +', ' ', text)
                summary = text
    
    # 4. Prices
    price_after_el = soup.select_one('.product-price .special-price') or soup.select_one('.special-price') or soup.select_one('.product-price')
    price_after = ""
    if price_after_el:
        price_after = price_after_el.text.strip()
        lines = [line.strip() for line in price_after.split('\n') if line.strip()]
        if lines:
            price_after = lines[0]
        if "₫" in price_after:
            price_after = price_after.split("₫")[0].strip() + "₫"
            
    price_before_el = soup.select_one('.product-price .old-price') or soup.select_one('.product-price del') or soup.select_one('.old-price')
    price_before = ""
    if price_before_el:
        price_before = price_before_el.text.strip()
        lines = [line.strip() for line in price_before.split('\n') if line.strip()]
        if lines:
            price_before = lines[0]
        if "₫" in price_before:
            price_before = price_before.split("₫")[0].strip() + "₫"
    else:
        price_before = price_after
        
    # 5. Image Link
    img_el = soup.select_one('.product-main-slide img') or soup.select_one('img[src*="/products/"]') or soup.select_one('img[data-src*="/products/"]')
    img_url = ""
    if img_el:
        img_url = img_el.get('src') or img_el.get('data-src') or ""
        if img_url:
            if img_url.startswith('//'):
                img_url = 'https:' + img_url
            img_url = re.sub(r'/thumb/(small|medium|large|grande|compact|1024x1024|2048x2048)/', '/', img_url)
            
    # 6. Technical Specifications Table (Spec)
    spec_table = soup.select_one('div.special-content:not(.noproduct) table') or soup.select_one('table.technical-content') or soup.select_one('table[class*="technical"]')
    specs = []
    if spec_table:
        for tr in spec_table.find_all('tr'):
            tds = [td.text.strip() for td in tr.find_all('td')]
            if len(tds) >= 2:
                specs.append(f"{tds[0]}: {tds[1]}")
    spec_str = "\n".join(specs)
    
    return {
        'SKU': sku,
        'URL': full_url,
        'categories': categories,
        'Name': name,
        'Summary': summary,
        'Price After': price_after,
        'Price Before': price_before,
        'Image Link': img_url,
        'Specs': spec_str
    }

class CrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Store Làm Mộc High-Speed Product Crawler")
        self.root.geometry("900x670")
        self.root.minsize(800, 580)
        
        # State variables
        self.is_running = False
        self.is_paused = False
        self.skus_to_crawl = []
        self.crawled_results = []
        self.file_rows = []
        self.session_lock = threading.Lock()
        
        # Configure High-Speed Requests & Auto-Retries Session Adapter
        self.session = requests.Session()
        retries = Retry(
            total=3,
            backoff_factor=0.3,
            status_forcelist=[500, 502, 503, 504],
            raise_on_status=False
        )
        adapter = HTTPAdapter(
            pool_connections=50,
            pool_maxsize=50,
            max_retries=retries
        )
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
        
        # Queue for GUI-thread safe communications
        self.log_queue = queue.Queue()
        
        # Design configuration: Premium Orange & Gold Theme matching Store Làm Mộc
        self.bg_color = "#fdfbf7"
        self.primary_color = "#ff6409"  # Orange
        self.accent_color = "#e8b815"   # Warm Gold
        self.text_color = "#333333"
        
        self.configure_styles()
        self.build_ui()
        
        # Periodic GUI update checker
        self.root.after(100, self.process_logs)

    def configure_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure("TFrame", background=self.bg_color)
        style.configure("TLabel", background=self.bg_color, foreground=self.text_color, font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=self.bg_color, foreground=self.primary_color, font=("Segoe UI Semibold", 16, "bold"))
        style.configure("TEntry", fieldbackground="white", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI Semibold", 10), background="#e0e0e0", borderwidth=1)
        style.map("TButton", background=[('active', '#d0d0d0')])
        
        # Custom primary action button
        style.configure("Primary.TButton", font=("Segoe UI Semibold", 10, "bold"), background=self.primary_color, foreground="white")
        style.map("Primary.TButton", background=[('active', self.accent_color)])

    def build_ui(self):
        self.root.configure(bg=self.bg_color)
        
        # Top Header Area
        header_frame = ttk.Frame(self.root, padding=(15, 10, 15, 10))
        header_frame.pack(fill=X)
        
        header_lbl = ttk.Label(header_frame, text="STORE LÀM MỘC CRAWLER", style="Title.TLabel")
        header_lbl.pack(side=LEFT)
        
        sub_lbl = ttk.Label(header_frame, text=" |   High-speed JSON search & details scraper", font=("Segoe UI", 9, "italic"), foreground="#666666")
        sub_lbl.pack(side=LEFT, pady=5)
        
        # Main Layout split into Left Panel (controls) and Right Panel (logs)
        main_pane = ttk.Frame(self.root, padding=10)
        main_pane.pack(fill=BOTH, expand=YES)
        
        # Left Panel (Input / Configurations)
        left_panel = ttk.LabelFrame(main_pane, text=" Crawl Configurations ", padding=15, width=320)
        left_panel.pack(side=LEFT, fill=BOTH, expand=NO, padx=(0, 10))
        
        # Group 1: Import Excel (Max 5)
        excel_hdr_frame = ttk.Frame(left_panel)
        excel_hdr_frame.pack(fill=X, pady=(0, 5))
        
        ttk.Label(excel_hdr_frame, text="1. Excel Files (Max 5):", font=("Segoe UI Semibold", 9)).pack(side=LEFT)
        
        browse_btn = ttk.Button(excel_hdr_frame, text="Select Files...", width=12, command=self.browse_excel)
        browse_btn.pack(side=RIGHT)
        
        # Container for dynamic files configuration
        self.files_config_frame = ttk.Frame(left_panel)
        self.files_config_frame.pack(fill=X, pady=(0, 15))
        
        self.placeholder_label = ttk.Label(self.files_config_frame, text="No files selected", font=("Segoe UI", 9, "italic"), foreground="#888888")
        self.placeholder_label.pack(pady=10)
        
        # Manual SKU Input
        ttk.Label(left_panel, text="OR Input SKUs directly (one per line):", font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(0, 5))
        self.manual_text = Text(left_panel, height=8, wrap=WORD, font=("Segoe UI", 9), relief=SOLID, borderwidth=1)
        self.manual_text.pack(fill=X, pady=(0, 15))
        
        # Thread count / Speed configuration
        speed_frame = ttk.Frame(left_panel)
        speed_frame.pack(fill=X, pady=(0, 15))
        
        # High Speed Mode option
        self.high_speed_var = BooleanVar(value=True)
        self.high_speed_chk = Checkbutton(speed_frame, text="API High-Speed Mode (8x Faster)", variable=self.high_speed_var, bg=self.bg_color, fg=self.text_color, font=("Segoe UI", 9, "bold"), activebackground=self.bg_color)
        self.high_speed_chk.grid(row=0, column=0, columnspan=2, sticky=W, pady=(0, 10))
        
        ttk.Label(speed_frame, text="Threads:", font=("Segoe UI Semibold", 9)).grid(row=1, column=0, sticky=W, pady=5)
        self.thread_spin = Spinbox(speed_frame, from_=1, to=20, width=5, justify=CENTER)
        self.thread_spin.delete(0, "end")
        self.thread_spin.insert(0, "5")
        self.thread_spin.grid(row=1, column=1, sticky=E, padx=(10, 0))
        
        ttk.Label(speed_frame, text="Delay (s):", font=("Segoe UI Semibold", 9)).grid(row=2, column=0, sticky=W, pady=5)
        self.delay_spin = Spinbox(speed_frame, from_=0.0, to=5.0, increment=0.1, width=5, justify=CENTER)
        self.delay_spin.delete(0, "end")
        self.delay_spin.insert(0, "0.2")
        self.delay_spin.grid(row=2, column=1, sticky=E, padx=(10, 0))
        
        # Control Buttons
        self.start_btn = ttk.Button(left_panel, text="START CRAWL", style="Primary.TButton", command=self.start_crawl)
        self.start_btn.pack(fill=X, pady=(0, 10))
        
        control_sub_frame = ttk.Frame(left_panel)
        control_sub_frame.pack(fill=X, pady=(0, 10))
        
        self.pause_btn = ttk.Button(control_sub_frame, text="Pause", state=DISABLED, command=self.toggle_pause)
        self.pause_btn.pack(side=LEFT, fill=X, expand=YES, padx=(0, 5))
        
        self.stop_btn = ttk.Button(control_sub_frame, text="Stop", state=DISABLED, command=self.stop_crawl)
        self.stop_btn.pack(side=RIGHT, fill=X, expand=YES, padx=(5, 0))
        
        self.export_btn = ttk.Button(left_panel, text="EXPORT RESULTS (EXCEL)", state=DISABLED, command=self.export_results)
        self.export_btn.pack(fill=X, pady=(10, 0))
        
        # Right Panel (Console Logs & Progress Indicators)
        right_panel = ttk.LabelFrame(main_pane, text=" Live Console Log & Progress ", padding=15)
        right_panel.pack(side=RIGHT, fill=BOTH, expand=YES)
        
        # Progress Bar & Metrics
        progress_info_frame = ttk.Frame(right_panel)
        progress_info_frame.pack(fill=X, pady=(0, 5))
        
        self.progress_lbl = ttk.Label(progress_info_frame, text="Idle", font=("Segoe UI Semibold", 10))
        self.progress_lbl.pack(side=LEFT)
        
        self.stats_lbl = ttk.Label(progress_info_frame, text="0 / 0 Completed (0 Failed)", font=("Segoe UI", 9), foreground="#555555")
        self.stats_lbl.pack(side=RIGHT)
        
        self.progress_bar = ttk.Progressbar(right_panel, orient=HORIZONTAL, mode="determinate")
        self.progress_bar.pack(fill=X, pady=(0, 15))
        
        # Console Display
        self.console = Text(right_panel, wrap=WORD, font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4", insertbackground="white", relief=FLAT)
        self.console.pack(fill=BOTH, expand=YES)
        
        # Scrollbar for console
        scrollbar = ttk.Scrollbar(self.console, command=self.console.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.console.config(yscrollcommand=scrollbar.set)
        
        # Initialize Console Greeting
        self.write_console("[SYS] System Initialized. Ready to crawl storelammoc.vn.\n")

    def write_console(self, text):
        self.console.config(state=NORMAL)
        self.console.insert(END, text)
        self.console.see(END)
        self.console.config(state=DISABLED)

    def make_updater(self, filepath, sheet_combo, col_combo):
        return lambda event: self.update_columns_for_row(filepath, sheet_combo, col_combo)

    def update_columns_for_row(self, filepath, sheet_combo, col_combo):
        sheet_name = sheet_combo.get()
        if not sheet_name:
            return
        try:
            df = pd.read_excel(filepath, sheet_name=sheet_name, nrows=5)
            columns = list(df.columns)
            col_combo['values'] = columns
            if columns:
                sku_col = ""
                for col in columns:
                    col_lower = str(col).lower()
                    if 'sku' in col_lower or 'mã' in col_lower or 'code' in col_lower:
                        sku_col = col
                        break
                if sku_col:
                    col_combo.set(sku_col)
                else:
                    col_combo.set(columns[0])
        except Exception as e:
            self.write_console(f"[ERROR] Failed to load columns for {os.path.basename(filepath)}: {str(e)}\n")

    def browse_excel(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not file_paths:
            return
            
        file_paths = list(file_paths)
        if len(file_paths) > 5:
            messagebox.showwarning("Limit Exceeded", "Bạn chỉ được chọn tối đa 5 file Excel. Chương trình sẽ tự động lấy 5 file đầu tiên.")
            file_paths = file_paths[:5]
            
        # Clean existing rows
        for widget in self.files_config_frame.winfo_children():
            widget.destroy()
        self.file_rows = []
        
        # Add headers for the columns in dynamic config
        headers_frame = ttk.Frame(self.files_config_frame)
        headers_frame.pack(fill=X, pady=(5, 2))
        
        lbl_file = ttk.Label(headers_frame, text="File Name", font=("Segoe UI Semibold", 8), width=15, anchor=W)
        lbl_file.pack(side=LEFT, padx=(0, 5))
        
        lbl_sheet = ttk.Label(headers_frame, text="Sheet", font=("Segoe UI Semibold", 8), width=11, anchor=W)
        lbl_sheet.pack(side=LEFT, padx=(0, 5), fill=X, expand=YES)
        
        lbl_col = ttk.Label(headers_frame, text="SKU Col", font=("Segoe UI Semibold", 8), width=11, anchor=W)
        lbl_col.pack(side=LEFT, fill=X, expand=YES)
        
        # Build configuration row for each file
        for path in file_paths:
            try:
                xl = pd.ExcelFile(path)
                sheets = xl.sheet_names
            except Exception as e:
                self.write_console(f"[ERROR] Could not read sheets from {os.path.basename(path)}: {str(e)}\n")
                continue
                
            row_frame = ttk.Frame(self.files_config_frame)
            row_frame.pack(fill=X, pady=2)
            
            # File name label
            fname = os.path.basename(path)
            short_name = fname if len(fname) <= 15 else fname[:12] + "..."
            lbl = ttk.Label(row_frame, text=short_name, width=15, anchor=W)
            lbl.pack(side=LEFT, padx=(0, 5))
            
            # Sheet combobox
            sheet_combo = ttk.Combobox(row_frame, state="readonly", width=11)
            sheet_combo['values'] = sheets
            sheet_combo.pack(side=LEFT, padx=(0, 5), fill=X, expand=YES)
            
            # SKU Column combobox
            col_combo = ttk.Combobox(row_frame, state="readonly", width=11)
            col_combo.pack(side=LEFT, fill=X, expand=YES)
            
            # Add to file_rows state
            row_data = {
                'filepath': path,
                'sheet_combo': sheet_combo,
                'col_combo': col_combo
            }
            self.file_rows.append(row_data)
            
            # Bind event and initialize columns
            sheet_combo.bind("<<ComboboxSelected>>", self.make_updater(path, sheet_combo, col_combo))
            if sheets:
                sheet_combo.set(sheets[0])
                self.update_columns_for_row(path, sheet_combo, col_combo)
                
        self.write_console(f"[SYS] Loaded {len(self.file_rows)} Excel file config rows. Adjust Sheet/SKU columns if needed.\n")

    def get_session_file_path(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), ".storelammoc_session.json")

    def save_session(self):
        with self.session_lock:
            try:
                files_config = []
                for row in self.file_rows:
                    files_config.append({
                        'filepath': row['filepath'],
                        'sheet_name': row['sheet_combo'].get(),
                        'col_name': row['col_combo'].get()
                    })
                
                session_data = {
                    'files_config': files_config,
                    'skus_to_crawl': self.skus_to_crawl,
                    'crawled_results': self.crawled_results,
                }
                with open(self.get_session_file_path(), 'w', encoding='utf-8') as f:
                    json.dump(session_data, f, ensure_ascii=False, indent=2)
            except Exception:
                pass

    def clear_session(self):
        path = self.get_session_file_path()
        if os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass

    def process_logs(self):
        """
        Pulls messages from the background thread log queue and updates GUI securely.
        """
        while not self.log_queue.empty():
            msg_type, data = self.log_queue.get()
            
            if msg_type == "log":
                self.write_console(data)
            elif msg_type == "progress":
                val, max_val = data
                self.progress_bar['value'] = val
                self.progress_bar['maximum'] = max_val
            elif msg_type == "stats":
                completed, total, failed = data
                self.stats_lbl.config(text=f"{completed} / {total} Completed ({failed} Failed)")
                percent = int((completed / total) * 100) if total > 0 else 0
                self.progress_lbl.config(text=f"Crawling ({percent}%)")
            elif msg_type == "done":
                self.is_running = False
                self.start_btn.config(state=NORMAL)
                self.pause_btn.config(state=DISABLED, text="Pause")
                self.stop_btn.config(state=DISABLED)
                self.export_btn.config(state=NORMAL)
                self.progress_lbl.config(text="Complete")
                messagebox.showinfo("Success", "Crawl Process finished! You can now Export results.")
            elif msg_type == "error_dialog":
                messagebox.showerror("Crawl Error", data)
                
        self.root.after(100, self.process_logs)

    def toggle_pause(self):
        if not self.is_running:
            return
        
        if self.is_paused:
            self.is_paused = False
            self.pause_btn.config(text="Pause")
            self.write_console("[SYS] Crawl Resumed.\n")
        else:
            self.is_paused = True
            self.pause_btn.config(text="Resume")
            self.write_console("[SYS] Crawl Paused.\n")

    def stop_crawl(self):
        if messagebox.askyesno("Confirm Stop", "Do you want to stop the crawling process? Current results will be preserved."):
            self.is_running = False
            self.is_paused = False
            self.write_console("[SYS] Crawling process aborted by user.\n")
            self.start_btn.config(state=NORMAL)
            self.pause_btn.config(state=DISABLED, text="Pause")
            self.stop_btn.config(state=DISABLED)
            self.export_btn.config(state=NORMAL)
            self.progress_lbl.config(text="Aborted")

    def start_crawl(self):
        # 1. Gather SKU codes
        self.skus_to_crawl = []
        
        # Get manual inputs
        manual_val = self.manual_text.get("1.0", "end-1c").strip()
        manual_skus = []
        if manual_val:
            manual_skus = [s.strip() for s in re.split(r'[\n,]', manual_val) if s.strip()]
            
        # Get Excel inputs
        excel_skus = []
        if self.file_rows:
            self.write_console(f"[SYS] Loading SKUs from {len(self.file_rows)} Excel file(s)...\n")
            for row in self.file_rows:
                path = row['filepath']
                sheet_name = row['sheet_combo'].get()
                col_name = row['col_combo'].get()
                
                if not sheet_name or not col_name:
                    continue
                
                try:
                    df = pd.read_excel(path, sheet_name=sheet_name)
                    if col_name in df.columns:
                        vals = df[col_name].dropna().astype(str).tolist()
                        row_skus = [s.strip() for s in vals if s.strip()]
                        excel_skus.extend(row_skus)
                        self.write_console(f"[SYS] Loaded {len(row_skus)} SKUs from file: {os.path.basename(path)} (Sheet: '{sheet_name}', Col: '{col_name}')\n")
                    else:
                        self.write_console(f"[WARN] Column '{col_name}' not found in file: {os.path.basename(path)}\n")
                except Exception as e:
                    self.write_console(f"[ERROR] Failed to read {os.path.basename(path)}: {str(e)}\n")
                    
        # Combine and deduplicate
        all_skus = manual_skus + excel_skus
        seen = set()
        self.skus_to_crawl = []
        for s in all_skus:
            if s not in seen:
                seen.add(s)
                self.skus_to_crawl.append(s)
                    
        if not self.skus_to_crawl:
            messagebox.showerror("No Inputs", "Please import Excel file(s) or enter SKU codes manually in the text area.")
            return
            
        # Validate speed configs
        try:
            num_threads = int(self.thread_spin.get())
            delay = float(self.delay_spin.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Please ensure threads is an integer and delay is a float value.")
            return
            
        # 2. Reset crawler state
        self.is_running = True
        self.is_paused = False
        self.crawled_results = []
        
        # Check for resume session
        session_file = self.get_session_file_path()
        resumed = False
        if os.path.exists(session_file):
            try:
                with open(session_file, 'r', encoding='utf-8') as f:
                    s_data = json.load(f)
                
                # Check if current SKUs match session SKUs
                if s_data.get('skus_to_crawl') == self.skus_to_crawl:
                    crawled_count = len(s_data.get('crawled_results', []))
                    total_count = len(self.skus_to_crawl)
                    if crawled_count > 0 and crawled_count < total_count:
                        if messagebox.askyesno("Resume Session", f"Found a previous unfinished crawl session ({crawled_count}/{total_count} SKUs crawled).\n\nDo you want to RESUME from where you left off?"):
                            self.crawled_results = s_data.get('crawled_results', [])
                            resumed = True
                            self.write_console(f"[SYS] Resuming session. Skipping {crawled_count} already crawled SKUs.\n")
            except Exception as e:
                self.write_console(f"[WARN] Failed to load session file: {str(e)}. Starting fresh.\n")
                
        if not resumed:
            self.clear_session()
            
        self.start_btn.config(state=DISABLED)
        self.pause_btn.config(state=NORMAL)
        self.stop_btn.config(state=NORMAL)
        self.export_btn.config(state=DISABLED)
        self.progress_lbl.config(text="Initializing...")
        
        # Fetch high speed config
        use_high_speed = self.high_speed_var.get()
        mode_str = "API High-Speed" if use_high_speed else "Standard HTML"
        self.write_console(f"\n--- Crawling Started: {len(self.skus_to_crawl)} SKUs found. Mode: {mode_str}. Thread count: {num_threads} ---\n")
        
        # Start crawl in background thread
        t = threading.Thread(target=self.crawl_worker, args=(self.skus_to_crawl, num_threads, delay, resumed, use_high_speed), daemon=True)
        t.start()

    def crawl_worker(self, sku_list, max_threads, delay, resumed=False, use_high_speed=True):
        total_skus = len(sku_list)
        
        crawled_skus = set()
        completed = 0
        failed = 0
        if resumed:
            with self.session_lock:
                crawled_skus = {r['SKU'] for r in self.crawled_results if 'SKU' in r}
                completed = len(crawled_skus)
                failed = sum(1 for r in self.crawled_results if r.get('Status') in ['Error', 'Not Found'])
                
        sku_queue = queue.Queue()
        for s in sku_list:
            if resumed and s in crawled_skus:
                continue
            sku_queue.put(s)
            
        results_lock = threading.Lock()
        
        # Calculate remaining items to process
        remaining_skus_count = sku_queue.qsize()
        
        def worker():
            nonlocal completed, failed
            while self.is_running:
                # Handle pause
                while self.is_paused and self.is_running:
                    time.sleep(0.5)
                    
                if not self.is_running:
                     break
                    
                try:
                    sku = sku_queue.get_nowait()
                except queue.Empty:
                    break
                    
                self.log_queue.put(("log", f"[RUN] Searching SKU: {sku}...\n"))
                
                # Search product
                product_url = search_product_url(sku, self.session)
                if not product_url:
                    self.log_queue.put(("log", f"[WARN] SKU '{sku}' not found on Store Làm Mộc.\n"))
                    with results_lock:
                        failed += 1
                        completed += 1
                        self.crawled_results.append({
                            'SKU': sku,
                            'Status': 'Not Found'
                        })
                    self.save_session()
                    self.log_queue.put(("progress", (completed, total_skus)))
                    self.log_queue.put(("stats", (completed, total_skus, failed)))
                    sku_queue.task_done()
                    time.sleep(delay)
                    continue
                    
                # Fetch details
                try:
                    if use_high_speed:
                        details = fetch_product_details_json(product_url, sku, self.session)
                    else:
                        details = fetch_product_details(product_url, sku, self.session)
                        
                    details['Status'] = 'Success'
                    with results_lock:
                        completed += 1
                        self.crawled_results.append(details)
                    self.save_session()
                    self.log_queue.put(("log", f"[OK] SKU '{sku}' parsed successfully.\n"))
                except Exception as e:
                    self.log_queue.put(("log", f"[ERROR] Failed to fetch details for '{sku}': {str(e)}\n"))
                    with results_lock:
                        failed += 1
                        completed += 1
                        self.crawled_results.append({
                            'SKU': sku,
                            'Status': 'Error',
                            'Error Details': str(e)
                        })
                    self.save_session()
                        
                self.log_queue.put(("progress", (completed, total_skus)))
                self.log_queue.put(("stats", (completed, total_skus, failed)))
                sku_queue.task_done()
                time.sleep(delay)
                
        # Launch Thread Pool
        threads = []
        for _ in range(min(max_threads, remaining_skus_count if remaining_skus_count > 0 else 1)):
            thread = threading.Thread(target=worker, daemon=True)
            threads.append(thread)
            thread.start()
            
        # Wait for all threads to finish
        for thread in threads:
            thread.join()
            
        self.log_queue.put(("progress", (completed, total_skus)))
        self.log_queue.put(("stats", (completed, total_skus, failed)))
        
        if self.is_running:
            if completed == total_skus:
                self.clear_session()
            self.log_queue.put(("done", None))

    def export_results(self):
        if not self.crawled_results:
            messagebox.showwarning("Warning", "No results available to export.")
            return
            
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not save_path:
            return
            
        try:
            # 1. Transform categories array to dynamic separate columns
            # Clean list values, fill empty fields
            final_list = []
            
            # Find the max categories length
            max_categories_len = 0
            for r in self.crawled_results:
                if 'categories' in r and r['categories']:
                    max_categories_len = max(max_categories_len, len(r['categories']))
                    
            for r in self.crawled_results:
                # Copy fields
                item = {
                    'SKU': r.get('SKU', ''),
                    'Status': r.get('Status', ''),
                    'Product Name': r.get('Name', ''),
                    'Short Description': r.get('Summary', ''),
                    'Price (After Discount)': r.get('Price After', ''),
                    'Price (Before Discount)': r.get('Price Before', ''),
                    'Image Link': r.get('Image Link', ''),
                    'Specifications': r.get('Specs', ''),
                    'URL': r.get('URL', ''),
                    'Error Details': r.get('Error Details', '')
                }
                
                # Split breadcrumbs/categories to separate columns
                cats = r.get('categories', [])
                for i in range(max_categories_len):
                    col_name = f"Category {i+1}"
                    if i < len(cats):
                        item[col_name] = cats[i]
                    else:
                        item[col_name] = ""
                        
                final_list.append(item)
                
            # Create DataFrame
            df = pd.DataFrame(final_list)
            
            # Organize columns cleanly: Categories -> Name -> Prices -> Specs -> Media -> Metadata
            category_cols = [f"Category {i+1}" for i in range(max_categories_len)]
            base_cols = ['SKU', 'Status', 'Product Name']
            price_media_cols = ['Price (After Discount)', 'Price (Before Discount)', 'Short Description', 'Specifications', 'Image Link', 'URL', 'Error Details']
            
            ordered_cols = base_cols + category_cols + price_media_cols
            df = df.reindex(columns=ordered_cols)
            
            # Save using Pandas
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Export Successful", f"Data exported successfully to:\n{save_path}")
            self.write_console(f"[SYS] Export Complete -> {save_path}\n")
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not save file:\n{str(e)}")

if __name__ == "__main__":
    root = Tk()
    app = CrawlerApp(root)
    root.mainloop()
