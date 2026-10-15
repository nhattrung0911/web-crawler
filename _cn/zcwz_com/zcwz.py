"""
==============================================================================
C.ZCWZ CRAWLER V1.0 - Production Ready
Based on Ketnoi V2.0 architecture (Pydantic + SQLAlchemy + Asyncio)
==============================================================================
"""

import asyncio
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import urljoin

import aiohttp
from pydantic import BaseModel, Field
from selectolax.parser import HTMLParser
from sqlalchemy import (
    Boolean, Column, DateTime, Float, Integer, String, Text, create_engine
)
from sqlalchemy.orm import declarative_base, Session, sessionmaker
from tenacity import (
    retry, stop_after_attempt, wait_exponential, retry_if_exception_type
)
import argparse

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
    handlers=[
        logging.FileHandler('c_zcwz_crawler.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# =====================================================
# CONFIGURATION
# =====================================================

class CrawlerConfig(BaseModel):
    concurrency: int = 15
    timeout: int = 10
    rate_limit_per_second: float = 10.0
    db_path: str = "data/zcwz.db"
    base_url: str = "https://c.zcwz.com"
    list_url_template: str = "https://c.zcwz.com/bearing/list_{page}.html"
    max_pages: int = 20116

# =====================================================
# DATABASE MODELS
# =====================================================

Base = declarative_base()

class ProductURL(Base):
    __tablename__ = 'product_urls'

    id = Column(Integer, primary_key=True)
    url = Column(String(500), unique=True, nullable=False, index=True)
    page_number = Column(Integer, nullable=False)
    discovered_at = Column(DateTime, default=datetime.utcnow)
    crawled = Column(Boolean, default=False, index=True)

class Product(Base):
    __tablename__ = 'products'

    id = Column(Integer, primary_key=True)
    url = Column(String(500), unique=True, nullable=False, index=True)

    name = Column(String(500))
    brand = Column(String(200), index=True)
    model = Column(String(200), index=True)
    series = Column(String(200))
    
    # Specs
    inner_d = Column(String(50))
    outer_D = Column(String(50))
    width_B = Column(String(50))
    weight_m = Column(String(50))
    
    # Extra Specs
    specifications = Column(Text)
    preview_info = Column(Text)
    images = Column(Text)

    crawled_at = Column(DateTime, default=datetime.utcnow)
    crawl_status = Column(String(50), default='success')
    error_message = Column(Text)

class DatabaseManager:
    def __init__(self, db_path: str):
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        self.engine = create_engine(f'sqlite:///{db_path}', connect_args={'check_same_thread': False})
        
        from sqlalchemy import event
        @event.listens_for(self.engine, "connect")
        def set_sqlite_pragma(dbapi_connection, connection_record):
            cursor = dbapi_connection.cursor()
            cursor.execute("PRAGMA journal_mode=WAL")
            cursor.execute("PRAGMA synchronous=OFF")
            cursor.close()
            
        Base.metadata.create_all(self.engine)
        self.SessionLocal = sessionmaker(bind=self.engine)
        logger.info(f"Database initialized at {db_path}")

    def get_session(self) -> Session:
        return self.SessionLocal()

    def add_url(self, url: str, page: int) -> bool:
        with self.get_session() as session:
            try:
                existing = session.query(ProductURL).filter_by(url=url).first()
                if existing: return False
                session.add(ProductURL(url=url, page_number=page))
                session.commit()
                return True
            except Exception as e:
                session.rollback()
                logger.error(f"Error adding {url}: {e}")
                return False

    def get_uncrawled(self, limit: Optional[int] = None) -> List[str]:
        with self.get_session() as session:
            query = session.query(ProductURL.url).filter_by(crawled=False)
            if limit: query = query.limit(limit)
            return [url for (url,) in query.all()]

    def mark_crawled(self, url: str):
        with self.get_session() as session:
            session.query(ProductURL).filter_by(url=url).update({'crawled': True})
            session.commit()

    def save_product(self, data: Dict[str, Any]):
        with self.get_session() as session:
            try:
                product = Product(**data)
                session.merge(product)
                session.commit()
            except Exception as e:
                session.rollback()
                logger.error(f"Error saving product {data.get('url')}: {e}")

# =====================================================
# HTTP CLIENT
# =====================================================

class RateLimiter:
    def __init__(self, rate: float):
        self.rate = rate
        self.tokens = rate
        self.last_update = time.monotonic()
        self.lock = asyncio.Lock()

    async def acquire(self):
        async with self.lock:
            now = time.monotonic()
            elapsed = now - self.last_update
            self.tokens = min(self.rate, self.tokens + elapsed * self.rate)
            self.last_update = now
            if self.tokens < 1:
                wait_time = (1 - self.tokens) / self.rate
                await asyncio.sleep(wait_time)
                self.tokens = 0
            else:
                self.tokens -= 1

class HTTPClient:
    def __init__(self, config: CrawlerConfig):
        self.config = config
        self.limiter = RateLimiter(config.rate_limit_per_second)
        self.session = None
        self.pause_event = asyncio.Event()
        self.pause_event.set()
        self._bg_tasks = set()

    async def _resume_after(self, delay: int):
        await asyncio.sleep(delay)
        logger.warning("WAF Pause ended. Resuming traffic.")
        self.pause_event.set()

    async def __aenter__(self):
        timeout = aiohttp.ClientTimeout(total=self.config.timeout)
        self.session = aiohttp.ClientSession(timeout=timeout)
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        if self.session: await self.session.close()

    @retry(stop=stop_after_attempt(5), wait=wait_exponential(multiplier=2, min=5, max=60), retry=retry_if_exception_type((aiohttp.ClientError, asyncio.TimeoutError)))
    async def fetch(self, url: str) -> str:
        await self.pause_event.wait()
        await self.limiter.acquire()
        async with self.session.get(url, headers={"User-Agent": "Mozilla/5.0"}) as resp:
            if resp.status == 404:
                return ""
            if resp.status in (403, 429, 503, 520, 521, 522):
                if self.pause_event.is_set():
                    logger.warning(f"WAF Blocked (Status {resp.status}) at {url}. Pausing all traffic for 60s...")
                    self.pause_event.clear()
                    task = asyncio.create_task(self._resume_after(60))
                    self._bg_tasks.add(task)
                    task.add_done_callback(self._bg_tasks.discard)
                resp.raise_for_status()
            resp.raise_for_status()
            return await resp.text()

# =====================================================
# PARSER & PIPELINES
# =====================================================

def parse_list(html: str, base_url: str) -> List[str]:
    tree = HTMLParser(html)
    urls = []
    for dt in tree.css("div.list_bottom dl dt a"):
        href = dt.attributes.get('href')
        if href:
            urls.append(urljoin(base_url, href))
    return urls

def parse_detail(html: str, url: str) -> Dict[str, Any]:
    tree = HTMLParser(html)
    data = {'url': url, 'crawl_status': 'success'}
    
    try:
        # Images
        images = []
        for img in tree.css("div#query_container01 div.swiper-slide img"):
            src = img.attributes.get('src')
            if src: images.append(src)
        data['images'] = " | ".join(images)

        # Title area (preview_info dl)
        dl = tree.css_first("div.preview_info dl")
        if dl:
            dt = dl.css_first("dt")
            if dt: data['name'] = dt.text(strip=True)
            
            preview_texts = []
            for dd in dl.css("dd"):
                text = dd.text(strip=True).replace("\xa0", " ")
                preview_texts.append(text)
                if text.startswith("品牌："):
                    data['brand'] = text.replace("品牌：", "").strip()
                elif text.startswith("类型系列："):
                    data['series'] = text.replace("类型系列：", "").strip()
                elif text.startswith("型号："):
                    data['model'] = text.replace("型号：", "").strip()
                else:
                    # e.g., "内径 : 26.988mm    外径 : 62mm    宽度 : 38.1mm    重量：0.32Kg"
                    parts = re.split(r'\s{2,}', text)
                    for part in parts:
                        if "d :" in part or "d:" in part or "内径" in part: data['inner_d'] = part.split(":")[-1].strip()
                        elif "D :" in part or "D:" in part or "外径" in part: data['outer_D'] = part.split(":")[-1].strip()
                        elif "B :" in part or "B:" in part or "宽度" in part: data['width_B'] = part.split(":")[-1].strip()
                        elif "重量" in part: data['weight_m'] = part.split("：")[-1].strip()
            data['preview_info'] = "\n".join(preview_texts)
        
        # Specs tables (#cpcs)
        specs = []
        for tr in tree.css("div#cpcs table tr"):
            th = tr.css_first("th")
            tds = tr.css("td")
            if th and len(tds) >= 1:
                key = th.text(strip=True)
                val = tds[0].text(strip=True)
                unit = tds[1].text(strip=True) if len(tds) > 1 and tds[1].text(strip=True) != '-' else ''
                val_str = f"{val}{unit}"
                specs.append(f"{key}: {val_str}")
                
        data['specifications'] = "\n".join(specs)
    except Exception as e:
        data['crawl_status'] = 'error'
        data['error_message'] = str(e)
        logger.error(f"Error parsing {url}: {e}")
        
    return data

async def run_pipe1(config: CrawlerConfig, db: DatabaseManager, start_page: int, end_page: int, mode: str = 'all'):
    logger.info(f"Starting Pipe 1 (mode: {mode}): pages {start_page} to {end_page}")
    sem = asyncio.Semaphore(config.concurrency)
    
    async def process_page(client: HTTPClient, page_url: str):
        async with sem:
            try:
                html = await client.fetch(page_url)
                if not html: return
                urls = parse_list(html, config.base_url)
                added = 0
                for u in urls:
                    if db.add_url(u, 1): added += 1
                if urls:
                    logger.info(f"Found {added}/{len(urls)} URLs at {page_url}")
            except Exception as e:
                logger.error(f"Error fetching {page_url}: {e}")

    async with HTTPClient(config) as client:
        if mode == 'all':
            tasks = [process_page(client, config.list_url_template.format(page=p)) for p in range(start_page, end_page + 1)]
            await asyncio.gather(*tasks)
        elif mode == 'brand-deep':
            import re
            
            async def discover_and_crawl_brand(client: HTTPClient, brand_id: int):
                cat_ids = []
                # Use semaphore for the initial discovery request
                async with sem:
                    root_url = f"{config.base_url}/bearing/{brand_id}_list_1.html"
                    try:
                        html = await client.fetch(root_url)
                        if not html: return
                        
                        # Add root items from page 1
                        urls = parse_list(html, config.base_url)
                        added = 0
                        for u in urls:
                            if db.add_url(u, 1): added += 1
                        if urls:
                            logger.info(f"Found {added}/{len(urls)} URLs at {root_url}")
                        
                        # Find sub-categories dynamically
                        cat_pattern = rf"href=[\'\"](?:/bearing/)?{brand_id}_(\d+)_list_1\.html[\'\"]"
                        cat_ids = list(set(re.findall(cat_pattern, html)))
                    except Exception as e:
                        logger.error(f"Error discovering brand {brand_id}: {e}")
                        return

                # Spawn tasks outside of this specific semaphore block (process_page uses its own)
                brand_tasks = []
                # Add root pages 2..10
                brand_tasks.extend([process_page(client, f"{config.base_url}/bearing/{brand_id}_list_{p}.html") for p in range(2, 11)])
                
                # Add category pages 1..10
                for cat_id in cat_ids:
                    brand_tasks.extend([process_page(client, f"{config.base_url}/bearing/{brand_id}_{cat_id}_list_{p}.html") for p in range(1, 11)])
                
                if brand_tasks:
                    # Execute in chunks to avoid slamming the event loop too hard
                    for i in range(0, len(brand_tasks), 50):
                        await asyncio.gather(*brand_tasks[i:i+50])

            # Iterate brands 1 to 150
            for brand_id in range(1, 151):
                logger.info(f"Crawling Brand ID {brand_id}...")
                await discover_and_crawl_brand(client, brand_id)

async def run_pipe_excel(config: CrawlerConfig, db: DatabaseManager, excel_path: str, col_name: str = 'Code'):
    import pandas as pd
    import json
    from urllib.parse import quote
    
    logger.info(f"Starting Pipe Excel Mode with file {excel_path}, column {col_name}")
    try:
        df = pd.read_excel(excel_path)
        if col_name not in df.columns:
            logger.warning(f"Column '{col_name}' not found in {excel_path}. Using first column instead.")
            codes = df.iloc[:, 0].dropna().astype(str).unique().tolist()
        else:
            codes = df[col_name].dropna().astype(str).unique().tolist()
        logger.info(f"Loaded {len(codes)} unique codes from Excel.")
    except Exception as e:
        logger.error(f"Failed to read Excel: {e}")
        return

    sem = asyncio.Semaphore(config.concurrency)
    
    async def search_and_add(client: HTTPClient, code: str):
        async with sem:
            try:
                url = f"{config.base_url}/bearing/search/getmodel.html?keyword={quote(code)}"
                resp_text = await client.fetch(url)
                if not resp_text: return
                data = json.loads(resp_text)
                if data.get('code') == 0 and data.get('data'):
                    added = 0
                    for item in data['data']:
                        enid = item.get('enid')
                        if enid:
                            detail_url = f"{config.base_url}/detail/{enid}.html"
                            if db.add_url(detail_url, 1): added += 1
                    logger.info(f"Code {code}: found {len(data['data'])}, added {added} URLs")
                else:
                    logger.info(f"Code {code}: no results found.")
            except Exception as e:
                logger.error(f"Error searching code {code}: {e}")

    async with HTTPClient(config) as client:
        tasks = [search_and_add(client, code) for code in codes]
        for i in range(0, len(tasks), 50):
            await asyncio.gather(*tasks[i:i+50])

async def run_pipe2(config: CrawlerConfig, db: DatabaseManager, limit: int = 0):
    logger.info("Starting Pipe 2 (Detail scraping)")
    urls = db.get_uncrawled(limit if limit > 0 else None)
    logger.info(f"Found {len(urls)} URLs to crawl")
    sem = asyncio.Semaphore(config.concurrency)
    
    async def process_url(client: HTTPClient, url: str):
        async with sem:
            try:
                html = await client.fetch(url)
                if not html:
                    db.mark_crawled(url)
                    return
                data = parse_detail(html, url)
                db.save_product(data)
                db.mark_crawled(url)
                logger.info(f"Saved: {data.get('name', url)}")
            except Exception as e:
                logger.error(f"Error detail {url}: {e}")

    async with HTTPClient(config) as client:
        tasks = [process_url(client, u) for u in urls]
        for i in range(0, len(tasks), 100):
            await asyncio.gather(*tasks[i:i+100])

def main():
    parser = argparse.ArgumentParser(description="zcwz.com Crawler")
    parser.add_argument('--pipe1', action='store_true', help='Run Pipe 1 (Extract URLs)')
    parser.add_argument('--pipe2', action='store_true', help='Run Pipe 2 (Extract Details)')
    parser.add_argument('--mode', type=str, choices=['all', 'brand-deep', 'excel'], default='all', help='Crawl mode')
    parser.add_argument('--excel_file', type=str, help='Path to Excel file for excel mode')
    parser.add_argument('--start', type=int, default=1, help='Start page')
    parser.add_argument('--end', type=int, default=20116, help='End page')
    parser.add_argument('--limit', type=int, default=10, help='Limit detail crawl')
    parser.add_argument('--export', choices=['csv', 'xlsx', 'json'], help='Export data to format')
    parser.add_argument('--gen-master', action='store_true', help='Generate Master_Bearing_Codes.xlsx')
    args = parser.parse_args()

    if args.gen_master:
        _gen_master()
        return

    config = CrawlerConfig()
    db = DatabaseManager(config.db_path)

    if args.export:
        import pandas as pd
        import sqlite3
        conn = sqlite3.connect(config.db_path)
        df = pd.read_sql("SELECT * FROM products", conn)
        conn.close()
        out_file = f"zcwz_export.{args.export}"
        if args.export == 'csv':
            df.to_csv(out_file, index=False, encoding='utf-8-sig')
        elif args.export == 'xlsx':
            df.to_excel(out_file, index=False)
        elif args.export == 'json':
            df.to_json(out_file, orient='records', force_ascii=False, indent=4)
        print(f"Exported {len(df)} records to {out_file}")
        return

    if args.pipe1:
        if args.mode == 'excel':
            if not args.excel_file:
                logger.error("Excel file required for excel mode")
                return
            asyncio.run(run_pipe_excel(config, db, args.excel_file))
        else:
            asyncio.run(run_pipe1(config, db, args.start, args.end, mode=args.mode))
    if args.pipe2:
        asyncio.run(run_pipe2(config, db, args.limit))


# =====================================================
# MASTER CODE GENERATOR
# =====================================================

def generate_bearing_codes():
    codes = []
    standard_prefixes = [
        "60", "62", "63", "64", "68", "69", "160", "618", "619",
        "70", "72", "73", "74", "32", "33", "52", "53",
        "302", "303", "313", "320", "322", "323", "329", "330", "331", "332",
        "NU2", "NU3", "NU4", "NU10", "NU22", "NU23",
        "NJ2", "NJ3", "NJ4", "NJ22", "NJ23",
        "NUP2", "NUP3", "NUP22", "NUP23",
        "N2", "N3", "N4", "NN30", "NNU49",
        "213", "222", "223", "230", "231", "232", "238", "239", "240", "241",
        "511", "512", "513", "514", "522", "523", "532", "533", "292", "293", "294",
    ]
    micro_suffixes = [str(i) for i in range(1, 10)]
    basic_suffixes = ["00", "01", "02", "03"]
    large_suffixes = [f"{i:02d}" for i in range(4, 100)]
    for p in standard_prefixes:
        if p in ["60", "62", "63", "68", "69"]:
            for s in micro_suffixes:
                codes.append(f"{p}{s}")
        for s in basic_suffixes + large_suffixes:
            codes.append(f"{p}{s}")
    pillow_prefixes = ["UC", "UK", "UCP", "UCF", "UCFL", "UCT", "UCPA", "UCFC"]
    pillow_suffixes = [str(i) for i in range(201, 221)] + [str(i) for i in range(305, 329)]
    for p in pillow_prefixes:
        for s in pillow_suffixes:
            codes.append(f"{p}{s}")
    return sorted(set(codes))


def _gen_master():
    import pandas as pd
    codes = generate_bearing_codes()
    pd.DataFrame({"Code": codes}).to_excel("Master_Bearing_Codes.xlsx", index=False)
    print(f"Generated {len(codes)} codes -> Master_Bearing_Codes.xlsx")


# =====================================================
# GUI
# =====================================================

def launch_gui():
    import customtkinter as ctk
    from tkinter import filedialog, messagebox

    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

    class RedirectText:
        def __init__(self, text_widget):
            self.text_widget = text_widget

        def write(self, string):
            self.text_widget.configure(state="normal")
            self.text_widget.insert("end", string)
            self.text_widget.see("end")
            self.text_widget.configure(state="disabled")

        def flush(self):
            pass

    class CrawlerApp(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title("ZCWZ Crawler Toolkit")
            self.geometry("800x600")
            self.grid_columnconfigure(1, weight=1)
            self.grid_rowconfigure(0, weight=1)

            self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
            self.sidebar.grid(row=0, column=0, sticky="nsew")
            self.sidebar.grid_rowconfigure(4, weight=1)
            ctk.CTkLabel(self.sidebar, text="ZCWZ Crawler", font=ctk.CTkFont(size=20, weight="bold")).grid(
                row=0, column=0, padx=20, pady=(20, 10))

            self.mode_var = ctk.StringVar(value="all")
            ctk.CTkRadioButton(self.sidebar, text="Lay ngau nhien (Root)", variable=self.mode_var,
                               value="all", command=self.change_mode).grid(row=1, column=0, pady=10, padx=20, sticky="w")
            ctk.CTkRadioButton(self.sidebar, text="Cao sau ngach (Brand x Cat)", variable=self.mode_var,
                               value="brand-deep", command=self.change_mode).grid(row=2, column=0, pady=10, padx=20, sticky="w")
            ctk.CTkRadioButton(self.sidebar, text="Tim theo Ma (Excel)", variable=self.mode_var,
                               value="excel", command=self.change_mode).grid(row=3, column=0, pady=10, padx=20, sticky="w")

            ctk.CTkLabel(self.sidebar, text="Export Data", font=ctk.CTkFont(weight="bold")).grid(
                row=5, column=0, padx=20, pady=(10, 0), sticky="w")
            ctk.CTkButton(self.sidebar, text="Export CSV", command=lambda: self.export_data("csv")).grid(
                row=6, column=0, pady=5, padx=20, sticky="ew")
            ctk.CTkButton(self.sidebar, text="Export XLSX", command=lambda: self.export_data("xlsx")).grid(
                row=7, column=0, pady=5, padx=20, sticky="ew")
            ctk.CTkButton(self.sidebar, text="Export JSON", command=lambda: self.export_data("json")).grid(
                row=8, column=0, pady=5, padx=20, sticky="ew")

            self.main_frame = ctk.CTkFrame(self)
            self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
            self.main_frame.grid_columnconfigure(0, weight=1)
            self.main_frame.grid_rowconfigure(3, weight=1)

            self.input_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
            self.input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
            self.input_frame.grid_columnconfigure(1, weight=1)
            self.input_label = ctk.CTkLabel(self.input_frame, text="Pages to crawl (e.g. 1-20116):")
            self.input_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
            self.input_entry = ctk.CTkEntry(self.input_frame, placeholder_text="1-20116")
            self.input_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
            self.btn_browse = ctk.CTkButton(self.input_frame, text="Browse", command=self.browse_file)
            self.btn_browse.grid(row=0, column=2, padx=5, pady=5)
            self.btn_browse.grid_remove()

            self.btn_start = ctk.CTkButton(self.main_frame, text="START CRAWLER", command=self.start_crawling, height=40)
            self.btn_start.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
            ctk.CTkLabel(self.main_frame, text="Console Output:").grid(row=2, column=0, padx=10, pady=(10, 0), sticky="w")
            self.console = ctk.CTkTextbox(self.main_frame, state="disabled", font=("Consolas", 12))
            self.console.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")

            sys.stdout = RedirectText(self.console)
            sys.stderr = RedirectText(self.console)
            self.file_path = ""
            self.change_mode()

        def change_mode(self):
            mode = self.mode_var.get()
            if mode == "all":
                self.input_label.configure(text="Pages (Start-End):")
                self.input_entry.configure(placeholder_text="1-10", state="normal")
                self.input_entry.delete(0, "end")
                self.input_entry.insert(0, "1-10")
                self.input_entry.grid()
                self.btn_browse.grid_remove()
            elif mode == "brand-deep":
                self.input_label.configure(text="No configuration needed")
                self.input_entry.delete(0, "end")
                self.input_entry.configure(state="disabled")
                self.input_entry.grid()
                self.btn_browse.grid_remove()
            elif mode == "excel":
                self.input_label.configure(text="Excel File:")
                self.input_entry.configure(placeholder_text="Select an excel file...", state="normal")
                self.input_entry.delete(0, "end")
                self.input_entry.grid()
                self.btn_browse.grid()

        def browse_file(self):
            self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
            if self.file_path:
                self.input_entry.delete(0, "end")
                self.input_entry.insert(0, self.file_path)

        def start_crawling(self):
            mode = self.mode_var.get()
            val = self.input_entry.get()
            self.btn_start.configure(state="disabled", text="RUNNING...")
            threading.Thread(target=self.run_crawler_task, args=(mode, val), daemon=True).start()

        def run_crawler_task(self, mode, val):
            print(f"--- STARTING CRAWLER MODE: {mode.upper()} ---")
            try:
                config = CrawlerConfig()
                db = DatabaseManager(config.db_path)
                if mode == "all":
                    try:
                        start_page, end_page = map(int, val.split("-"))
                    except Exception:
                        start_page, end_page = 1, 10
                    print(f"Running Pipe 1 for root pages {start_page} to {end_page}")
                    asyncio.run(run_pipe1(config, db, start_page, end_page, mode='all'))
                elif mode == "brand-deep":
                    print("Running Pipe 1 for all brands and sub-pages")
                    asyncio.run(run_pipe1(config, db, 1, 10, mode='brand-deep'))
                elif mode == "excel":
                    if not os.path.exists(val):
                        print(f"File not found: {val}")
                        return
                    print(f"Running Pipe Excel for {val}")
                    asyncio.run(run_pipe_excel(config, db, val))
                print("Running Pipe 2 to scrape all discovered URLs")
                asyncio.run(run_pipe2(config, db, 0))
                print("--- DONE ---")
            except Exception as e:
                print(f"ERROR: {e}")
            finally:
                self.btn_start.configure(state="normal", text="START CRAWLER")

        def export_data(self, fmt):
            try:
                import sqlite3
                import pandas as pd
                conn = sqlite3.connect("data/zcwz.db")
                df = pd.read_sql("SELECT * FROM products", conn)
                conn.close()
                if df.empty:
                    print("No data found in database to export.")
                    return
                out_file = f"zcwz_export.{fmt}"
                if fmt == "csv":
                    df.to_csv(out_file, index=False, encoding="utf-8-sig")
                elif fmt == "xlsx":
                    df.to_excel(out_file, index=False)
                elif fmt == "json":
                    df.to_json(out_file, orient="records", force_ascii=False, indent=4)
                print(f"Exported {len(df)} records to {out_file} successfully.")
            except Exception as e:
                print(f"Export error: {e}")

    CrawlerApp().mainloop()


if __name__ == '__main__':
    if len(sys.argv) > 1:
        main()
    else:
        launch_gui()
