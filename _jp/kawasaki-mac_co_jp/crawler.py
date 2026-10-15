# backend.py
# ==========================================================
# Kawasaki Tools Crawler Backend - ENTERPRISE V2
# - aiohttp + asyncio (high-performance)
# - autosave CSV (append + flush)
# - resume (state.json + existing CSV)
# - Pipe 1: crawl categories + product list
# - Pipe 2: crawl details (ALL / by parts), optimized by grouping category pages
# - Pipe 3: download images + pdf/print, export ZIP
# ==========================================================

import asyncio
import aiohttp
import csv
import json
import os
import re
import time
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Set, Tuple
from bs4 import BeautifulSoup
from yarl import URL
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

# ---------------- CONFIG ----------------
@dataclass(frozen=True)
class Config:
    base_url: str = "https://tools.kawasaki-mac.co.jp"
    entry_category: str = "https://tools.kawasaki-mac.co.jp/products/category/"
    out_dir: str = "out"
    pipe1_csv: str = "out/pipe1_products.csv"
    pipe2_csv: str = "out/pipe2_details.csv"
    state1: str = "out/state_pipe1.json"
    state2: str = "out/state_pipe2.json"
    downloads_dir: str = "out/downloads"

    timeout_s: int = 30
    max_conn: int = 30
    concurrency: int = 20
    chunk_size: int = 60
    retry_attempts: int = 4

CFG = Config()

PIPE1_FIELDS = [
    "category_id", "category_name", "category_url", "category_image",
    "product_id", "part_number", "product_url", "print_url"
]
PIPE2_FIELDS = [
    "category_name", "category_id",
    "product_id",
    "part_number",       # <h1>
    "product_name",      # <h4> type
    "spec",              # single cell: "Key: Value\nKey: Value..."
    "image_urls",        # ";" separated
    "pdf_urls",          # ";" separated (if any)
    "product_url",
    "print_url",
]

# ---------------- UTILS ----------------
def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)

def normalize_part(s: str) -> str:
    if not s:
        return ""
    s = s.strip().upper()
    # remove spaces & punctuation, keep alnum
    s = re.sub(r"[^A-Z0-9]+", "", s)
    return s

def split_parts_input(text: str) -> List[str]:
    if not text:
        return []
    raw = re.split(r"[,\n;\t ]+", text.strip())
    return [x.strip() for x in raw if x.strip()]

def chunked(lst: List, n: int):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

def load_json(path: str) -> Dict:
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def save_json(path: str, obj: Dict) -> None:
    ensure_dir(os.path.dirname(path) or ".")
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def read_existing_keys(csv_path: str, key_col: str) -> Set[str]:
    if not os.path.exists(csv_path) or os.path.getsize(csv_path) == 0:
        return set()
    out = set()
    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames or key_col not in reader.fieldnames:
            return set()
        for r in reader:
            v = (r.get(key_col) or "").strip()
            if v:
                out.add(v)
    return out

class CsvAppendWriter:
    def __init__(self, path: str, fieldnames: List[str]):
        ensure_dir(os.path.dirname(path) or ".")
        self.path = path
        self._f = open(path, "a", newline="", encoding="utf-8")
        self._w = csv.DictWriter(self._f, fieldnames=fieldnames)
        if os.path.getsize(path) == 0:
            self._w.writeheader()
            self._f.flush()

    def write(self, row: Dict):
        self._w.writerow(row)
        self._f.flush()

    def close(self):
        try:
            self._f.flush()
        finally:
            self._f.close()

# ---------------- HTTP ----------------
class HttpClient:
    def __init__(self, timeout_s: int, max_conn: int):
        timeout = aiohttp.ClientTimeout(total=timeout_s)
        connector = aiohttp.TCPConnector(limit=max_conn, ttl_dns_cache=300)
        self.session = aiohttp.ClientSession(
            timeout=timeout,
            connector=connector,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                              "AppleWebKit/537.36 (KHTML, like Gecko) "
                              "Chrome/120 Safari/537.36"
            }
        )

    async def close(self):
        if not self.session.closed:
            await self.session.close()

    @retry(
        reraise=True,
        stop=lambda retry_state: stop_after_attempt(CFG.retry_attempts)(retry_state),
        wait=wait_exponential(multiplier=0.7, min=0.7, max=8),
        retry=retry_if_exception_type((aiohttp.ClientError, asyncio.TimeoutError)),
    )
    async def get_text(self, url: str) -> str:
        async with self.session.get(url) as r:
            r.raise_for_status()
            return await r.text()

    @retry(
        reraise=True,
        stop=lambda retry_state: stop_after_attempt(CFG.retry_attempts)(retry_state),
        wait=wait_exponential(multiplier=0.7, min=0.7, max=8),
        retry=retry_if_exception_type((aiohttp.ClientError, asyncio.TimeoutError)),
    )
    async def get_bytes(self, url: str) -> bytes:
        async with self.session.get(url) as r:
            r.raise_for_status()
            return await r.read()

# ---------------- PARSERS ----------------
def parse_categories(html: str) -> List[Dict]:
    soup = BeautifulSoup(html, "lxml")
    cats = []
    seen = set()

    for a in soup.select('a[href*="/products/category/?id="]'):
        href = (a.get("href") or "").strip()
        if not href:
            continue
        try:
            u = str(URL(href))
        except Exception:
            continue

        m = re.search(r"[?&]id=(\d+)", u)
        if not m:
            continue
        cid = m.group(1)
        if cid in seen:
            continue
        seen.add(cid)

        title = ""
        h3 = a.select_one(".category-product-slider__info__title")
        if h3:
            title = h3.get_text(strip=True)
        else:
            title = a.get_text(" ", strip=True)

        img = ""
        div_img = a.select_one(".category-product-slider__image")
        if div_img:
            style = div_img.get("style", "")
            mm = re.search(r"url\(([^)]+)\)", style)
            if mm:
                img = mm.group(1).strip().strip('"').strip("'")

        cats.append({
            "category_id": cid,
            "category_name": title,
            "category_url": u,
            "category_image": img,
        })

    return cats

def parse_products_from_category_page(category_url: str, html: str) -> List[Dict]:
    soup = BeautifulSoup(html, "lxml")
    out = []
    seen = set()

    for a in soup.select(".fixnav__body__inner a.fixnav__unit__title, a.fixnav__unit__title"):
        title = a.get_text(" ", strip=True)
        href = (a.get("href") or "").strip()
        if not href:
            continue
        abs_url = href if href.startswith("http") else str(URL(category_url).join(URL(href)))
        mm = re.search(r"#sec_(\d+)", abs_url)
        if not mm:
            continue
        pid = mm.group(1)
        if pid in seen:
            continue
        seen.add(pid)

        out.append({
            "product_id": pid,
            "part_number": title,
            "product_url": abs_url,
            "print_url": str(URL(category_url).with_path("/products/print/").with_query({"id": pid})),
        })
    return out

def _find_section_root(soup: BeautifulSoup, product_id: str):
    sec = soup.find(id=f"sec_{product_id}")
    if sec is not None:
        return sec
    # fallback: search any element containing anchor "sec_###"
    anchor = soup.find(attrs={"name": f"sec_{product_id}"})
    return anchor if anchor is not None else None

def parse_detail_from_category_html(product_id: str, html: str) -> Dict:
    soup = BeautifulSoup(html, "lxml")
    sec = _find_section_root(soup, product_id)
    scope = sec if sec else soup

    h4 = scope.select_one("h4.item__info__type")
    h1 = scope.select_one("h1.item__info__title")

    product_name = h4.get_text(" ", strip=True) if h4 else ""
    part_number = h1.get_text(" ", strip=True) if h1 else ""

    # specs table
    spec_lines = []
    for row in scope.select("div.item__info__table__row"):
        head = row.select_one("div.item__info__table__head")
        body = row.select_one("div.item__info__table__body")
        k = head.get_text(" ", strip=True) if head else ""
        v = body.get_text(" ", strip=True) if body else ""
        if k or v:
            spec_lines.append(f"{k}: {v}".strip(": ").strip())
    spec_text = "\n".join(spec_lines)

    # images
    imgs = []
    for img in scope.select("div.item__image__nav img[src], div.item__image__nav__unit img[src]"):
        src = (img.get("src") or "").strip()
        if src and src not in imgs:
            imgs.append(src)

    # pdf/doc links (if present)
    pdfs = []
    for a in scope.select("div.item__info__documents a[href], a[href$='.pdf'], a[href*='.pdf?']"):
        href = (a.get("href") or "").strip()
        if href and ".pdf" in href.lower() and href not in pdfs:
            pdfs.append(href)

    return {
        "part_number": part_number,
        "product_name": product_name,
        "spec": spec_text,
        "image_urls": imgs,
        "pdf_urls": pdfs,
        "has_section": bool(sec),
    }

# ---------------- PIPE 1 ----------------
async def pipe1_run(
    cfg: Config = CFG,
    log: Callable[[str], None] = lambda _: None,
    progress: Callable[[float], None] = lambda _: None,
) -> str:
    ensure_dir(cfg.out_dir)
    state = load_json(cfg.state1)
    done_ids = set(state.get("done_product_ids", [])) | read_existing_keys(cfg.pipe1_csv, "product_id")

    http = HttpClient(cfg.timeout_s, cfg.max_conn)
    writer = CsvAppendWriter(cfg.pipe1_csv, PIPE1_FIELDS)

    try:
        log(f"[Pipe1] GET {cfg.entry_category}")
        entry_html = await http.get_text(cfg.entry_category)
        cats = parse_categories(entry_html)
        log(f"[Pipe1] Categories: {len(cats)}")

        sem = asyncio.Semaphore(cfg.concurrency)

        async def fetch_cat(cat: Dict) -> Tuple[Dict, str]:
            async with sem:
                html = await http.get_text(cat["category_url"])
                return cat, html

        total = len(cats)
        done_cat = 0

        for batch in chunked(cats, max(1, cfg.chunk_size // 5)):
            results = await asyncio.gather(*[fetch_cat(c) for c in batch], return_exceptions=True)
            for res in results:
                if isinstance(res, Exception):
                    log(f"[Pipe1][WARN] Category fetch failed: {res}")
                    continue
                cat, cat_html = res
                products = parse_products_from_category_page(cat["category_url"], cat_html)

                new_count = 0
                for p in products:
                    pid = str(p["product_id"])
                    if pid in done_ids:
                        continue
                    writer.write({**cat, **p})
                    done_ids.add(pid)
                    new_count += 1

                log(f"[Pipe1] {cat['category_name']} id={cat['category_id']}: products={len(products)} new={new_count}")

                done_cat += 1
                progress(done_cat / max(1, total))

                save_json(cfg.state1, {"done_product_ids": sorted(done_ids)})

        log(f"[Pipe1] DONE. Total products in CSV: {len(done_ids)}")
        return cfg.pipe1_csv

    finally:
        writer.close()
        await http.close()

# ---------------- PIPE 2 ----------------
async def pipe2_run(
    mode: str = "all",                      # "all" | "parts"
    parts_input: Optional[List[str]] = None,
    cfg: Config = CFG,
    log: Callable[[str], None] = lambda _: None,
    progress: Callable[[float], None] = lambda _: None,
) -> str:
    ensure_dir(cfg.out_dir)

    parts_input = parts_input or []
    want = {normalize_part(x) for x in parts_input if x.strip()}
    want.discard("")

    state = load_json(cfg.state2)
    done_ids = set(state.get("done_product_ids", [])) | read_existing_keys(cfg.pipe2_csv, "product_id")

    # read pipe1
    if not os.path.exists(cfg.pipe1_csv) or os.path.getsize(cfg.pipe1_csv) == 0:
        raise RuntimeError("Pipe1 CSV missing. Run Pipe 1 first.")

    with open(cfg.pipe1_csv, "r", encoding="utf-8") as f:
        src = list(csv.DictReader(f))

    if mode == "parts":
        before = len(src)
        src = [r for r in src if normalize_part(r.get("part_number", "")) in want]
        log(f"[Pipe2] Filter by parts: {before} -> {len(src)}")
    else:
        log(f"[Pipe2] Mode ALL: {len(src)} products")

    # Group by category_url to avoid re-downloading same page many times (huge speedup)
    by_cat: Dict[str, List[Dict]] = {}
    for r in src:
        by_cat.setdefault(r["category_url"], []).append(r)

    cat_urls = list(by_cat.keys())
    total_products = sum(len(v) for v in by_cat.values())
    processed = 0

    http = HttpClient(cfg.timeout_s, cfg.max_conn)
    writer = CsvAppendWriter(cfg.pipe2_csv, PIPE2_FIELDS)

    sem = asyncio.Semaphore(cfg.concurrency)

    async def fetch_cat_page(url: str) -> Tuple[str, str]:
        async with sem:
            html = await http.get_text(url)
            return url, html

    try:
        t0 = time.time()
        for batch in chunked(cat_urls, max(1, cfg.chunk_size // 5)):
            results = await asyncio.gather(*[fetch_cat_page(u) for u in batch], return_exceptions=True)
            for res in results:
                if isinstance(res, Exception):
                    log(f"[Pipe2][WARN] Category page fetch failed: {res}")
                    continue

                cat_url, cat_html = res
                items = by_cat.get(cat_url, [])

                for item in items:
                    pid = (item.get("product_id") or "").strip()
                    if not pid:
                        processed += 1
                        progress(processed / max(1, total_products))
                        continue
                    if pid in done_ids:
                        processed += 1
                        progress(processed / max(1, total_products))
                        continue

                    detail = parse_detail_from_category_html(pid, cat_html)

                    # If section not found, we still write row with minimal data to avoid infinite retry loops
                    row = {
                        "category_name": item.get("category_name", ""),
                        "category_id": item.get("category_id", ""),
                        "product_id": pid,
                        "part_number": detail.get("part_number") or item.get("part_number", ""),
                        "product_name": detail.get("product_name", ""),
                        "spec": detail.get("spec", ""),
                        "image_urls": ";".join(detail.get("image_urls", [])),
                        "pdf_urls": ";".join(detail.get("pdf_urls", [])),
                        "product_url": item.get("product_url", ""),
                        "print_url": item.get("print_url", ""),
                    }
                    writer.write(row)

                    done_ids.add(pid)
                    processed += 1
                    progress(processed / max(1, total_products))

                    # checkpoint frequently (enterprise-safe)
                    if processed % 40 == 0:
                        save_json(cfg.state2, {
                            "done_product_ids": sorted(done_ids),
                            "mode": mode,
                            "parts_input": parts_input
                        })

        save_json(cfg.state2, {
            "done_product_ids": sorted(done_ids),
            "mode": mode,
            "parts_input": parts_input
        })
        dt = time.time() - t0
        log(f"[Pipe2] DONE. Details total={len(done_ids)} time={dt:.1f}s")
        return cfg.pipe2_csv

    finally:
        writer.close()
        await http.close()

# ---------------- PIPE 3 ----------------
def _safe_folder_name(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return "unknown"
    s2 = re.sub(r"[^A-Za-z0-9\-_]+", "_", s)
    s2 = s2.strip("_")
    return (s2[:120] or "unknown")

async def pipe3_download_and_zip(
    parts_filter: Optional[List[str]] = None,
    download_images: bool = True,
    download_pdfs: bool = True,
    include_print_html: bool = True,
    cfg: Config = CFG,
    log: Callable[[str], None] = lambda _: None,
    progress: Callable[[float], None] = lambda _: None,
) -> Tuple[str, str]:
    """
    Reads Pipe2 CSV, downloads assets into out/downloads/<part>/,
    then zips the downloads dir and returns (downloads_dir, zip_path).
    """
    ensure_dir(cfg.downloads_dir)

    if not os.path.exists(cfg.pipe2_csv) or os.path.getsize(cfg.pipe2_csv) == 0:
        raise RuntimeError("Pipe2 CSV missing. Run Pipe 2 first (or generate details for chosen parts).")

    want = set()
    if parts_filter:
        want = {normalize_part(x) for x in parts_filter if x.strip()}
        want.discard("")

    rows = []
    with open(cfg.pipe2_csv, "r", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            pn = r.get("part_number", "") or ""
            if want and normalize_part(pn) not in want:
                continue
            rows.append(r)

    # Build download list
    jobs: List[Tuple[str, str]] = []  # (url, filepath)
    for r in rows:
        part = _safe_folder_name(r.get("part_number", "unknown"))
        folder = os.path.join(cfg.downloads_dir, part)
        ensure_dir(folder)

        if download_images:
            imgs = [x.strip() for x in (r.get("image_urls") or "").split(";") if x.strip()]
            for i, u in enumerate(imgs, 1):
                ext = os.path.splitext(URL(u).path)[1] or ".png"
                jobs.append((u, os.path.join(folder, f"image_{i:02d}{ext}")))

        if download_pdfs:
            pdfs = [x.strip() for x in (r.get("pdf_urls") or "").split(";") if x.strip()]
            for i, u in enumerate(pdfs, 1):
                ext = os.path.splitext(URL(u).path)[1] or ".pdf"
                if ".pdf" not in ext.lower():
                    ext = ".pdf"
                jobs.append((u, os.path.join(folder, f"doc_{i:02d}{ext}")))

        if include_print_html:
            pr = (r.get("print_url") or "").strip()
            if pr:
                jobs.append((pr, os.path.join(folder, "print.html")))

    if not jobs:
        log("[Pipe3] No files to download (check filters).")
        # still create zip of current downloads_dir
        return cfg.downloads_dir, _zip_dir(cfg.downloads_dir)

    http = HttpClient(cfg.timeout_s, cfg.max_conn)
    sem = asyncio.Semaphore(cfg.concurrency)
    done = 0
    total = len(jobs)

    async def dl_one(url: str, path: str):
        nonlocal done
        async with sem:
            try:
                data = await http.get_bytes(url)
                ensure_dir(os.path.dirname(path))
                with open(path, "wb") as f:
                    f.write(data)
            except Exception as e:
                log(f"[Pipe3][WARN] Failed: {url} | {e}")
            finally:
                done += 1
                progress(done / max(1, total))

    try:
        # Execute in manageable chunks
        for batch in chunked(jobs, max(50, cfg.chunk_size)):
            await asyncio.gather(*[dl_one(u, p) for u, p in batch], return_exceptions=True)

        await http.close()

        zip_path = _zip_dir(cfg.downloads_dir)
        log(f"[Pipe3] DONE. downloads={cfg.downloads_dir} zip={zip_path}")
        return cfg.downloads_dir, zip_path

    finally:
        await http.close()

def _zip_dir(folder: str) -> str:
    import zipfile
    base = os.path.abspath(folder)
    out_zip = os.path.abspath(os.path.join(os.path.dirname(base), "downloads.zip"))
    # recreate zip
    if os.path.exists(out_zip):
        os.remove(out_zip)

    with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(base):
            for fn in files:
                fp = os.path.join(root, fn)
                rel = os.path.relpath(fp, base)
                z.write(fp, arcname=rel)
    return out_zip
