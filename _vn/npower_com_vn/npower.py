
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
npower_scraper.py — Crawler chuyên biệt cho https://npower.com.vn/
- Đọc Excel cột Part (SKU/mã) — có thể CHỌN FILE bằng hộp thoại nếu không truyền --excel
- Tìm kiếm trên site qua ô search -> chọn kết quả đúng mã
- Vào trang sản phẩm (PDP)
- Thu thập TẤT CẢ ảnh từ gallery (thumbnail + ảnh lớn)
- Tải về với kích thước lớn nhất có thể (ưu tiên 1200x1200 khi là CDN Shopify)
- Xuất báo cáo CSV/XLSX

Cách chạy:
  python npower_scraper.py --headless true  # sẽ bật hộp thoại chọn Excel
  hoặc:
  python npower_scraper.py --excel "list.xlsx" --headless true --outdir out_npower

Yêu cầu:
  pip install playwright pandas requests openpyxl
  playwright install
"""

from __future__ import annotations
import os, re, csv, time, json, argparse, sys, hashlib
from typing import List, Optional, Tuple
from pathlib import Path
from urllib.parse import urlparse, unquote, parse_qs

import pandas as pd
import requests
from playwright.sync_api import sync_playwright, Page, BrowserContext, TimeoutError as PlaywrightTimeoutError

# ---- GUI picker (tkinter) ----
def pick_excel_gui() -> Optional[str]:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.update()
        file_path = filedialog.askopenfilename(
            title="Chọn file Excel chứa cột Part/SKU",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        root.destroy()
        return file_path if file_path else None
    except Exception:
        return None

NPOWER_BASE = "https://npower.com.vn/"
DEFAULT_OUTDIR = "out_npower"
DEFAULT_HEADLESS = True
DEFAULT_RATEWAIT = 0.4
DEFAULT_RETRIES = 2
DEFAULT_TIMEOUT = 30

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119 Safari/537.36"

def now() -> str:
    return time.strftime("%H:%M:%S")

def log(msg: str) -> None:
    print(f"[{now()}] {msg}", flush=True)

def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)

def slugify(value: str, maxlen: int = 120) -> str:
    value = re.sub(r"[^\w\s-]+", "-", value, flags=re.UNICODE)
    value = re.sub(r"\s+", "_", value.strip())
    value = re.sub(r"-{2,}", "-", value)
    return value[:maxlen] or "part"

def normalize(text: str) -> str:
    s = (text or "").strip().upper()
    s = re.sub(r"[^\w]+", "", s)
    return s

def parse_bool(s: str) -> bool:
    return str(s).lower() in ("1","true","t","yes","y")

def read_parts_from_excel(xlsx_path: str, sheet: Optional[str]) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path, sheet_name=sheet) if sheet else pd.read_excel(xlsx_path)
    cols = {c.strip(): c for c in df.columns}
    lower_map = {k.lower(): v for k, v in cols.items()}
    candidates = ["part", "sku", "code", "mã", "ma", "product code", "part number", "part_number"]
    part_col = None
    for cand in candidates:
        if cand in lower_map:
            part_col = lower_map[cand]
            break
    if not part_col:
        raise RuntimeError("Không tìm thấy cột Part/SKU trong Excel (thử: Part, SKU, Code, Mã, ...)")
    df = df.rename(columns={part_col: "Part"})
    df["Part"] = df["Part"].astype(str).fillna("").str.strip()
    df = df[df["Part"]!=""]
    return df.reset_index(drop=True)

def open_home(page: Page) -> None:
    page.goto(NPOWER_BASE, timeout=45000, wait_until="domcontentloaded")
    try:
        page.wait_for_selector("input.form-field-input.live-search-form-field[name='q']", timeout=15000)
    except Exception:
        pass

def search_part(page: Page, part: str) -> None:
    try:
        inp = page.locator("input.form-field-input.live-search-form-field[name='q']").first
    except Exception:
        inp = page.locator("input[name='q']").first
    inp.fill("")
    inp.fill(part, timeout=15000)
    # Try click search button, else Enter
    try:
        btn = page.locator("button.live-search-button").first
        if btn and btn.count()>0:
            btn.click()
        else:
            inp.press("Enter")
    except Exception:
        inp.press("Enter")
    # Wait results grid
    try:
        page.wait_for_selector(".productgrid--items", timeout=15000)
    except Exception:
        pass

def pick_best_card(page: Page, part: str):
    target = normalize(part)
    cards = page.locator(".productgrid--items .productgrid--item")
    n = cards.count()
    best = None
    for i in range(n):
        c = cards.nth(i)
        title = ""
        try:
            title = c.locator("h2.productitem--title, h2 a").first.inner_text(timeout=800)
        except Exception:
            pass
        if target and normalize(title).find(target) != -1:
            return c
        # secondary: description contains code
        try:
            desc = c.locator(".productitem--description").first.inner_text(timeout=500)
            if target and normalize(desc).find(target) != -1:
                best = best or c
        except Exception:
            pass
    return best or (cards.first if n>0 else None)

def open_card(page: Page, card) -> bool:
    try:
        with page.expect_navigation(timeout=20000):
            card.locator("a[data-product-page-link], a.productitem--image-link, h2 a").first.click()
        return True
    except Exception:
        # fallback: focus + Enter
        try:
            link = card.locator("a").first
            link.focus()
            with page.expect_navigation(timeout=20000):
                page.keyboard.press("Enter")
            return True
        except Exception:
            return False

def confirm_pdp(page: Page, part: str) -> bool:
    target = normalize(part)
    if target in normalize(page.url):
        return True
    # Title, sku, body scan
    for sel in ["h1", "h2", ".productitem--title", ".product__title", ".product__sku"]:
        for i in range(min(page.locator(sel).count(), 5)):
            try:
                txt = page.locator(sel).nth(i).inner_text(timeout=800)
                if target in normalize(txt):
                    return True
            except Exception:
                pass
    try:
        body = page.locator("body").inner_text(timeout=1200)
        return target in normalize(body)
    except Exception:
        return False

def collect_gallery_urls(page: Page) -> list[str]:
    """
    Gom URL ảnh từ:
    - Thumbnail: .gallery-navigation--scroller [data-gallery-thumbnail] img (src, srcset, data-...)
    - Ảnh lớn sau khi click từng thumbnail: main image (src, srcset)
    Chuẩn hoá: thêm https:, nâng kích thước _75x75_crop_center -> _1200x1200 (Shopify)
    """
    urls: list[str] = []
    # 1) thumbnails
    thumbs = page.locator(".gallery-navigation--scroller [data-gallery-thumbnail] img")
    tcount = thumbs.count()
    for i in range(min(tcount, 50)):
        node = thumbs.nth(i)
        for attr in ("src","data-src","data-rimg-template"):
            v = node.get_attribute(attr)
            if v: urls.append(v)
        ss = node.get_attribute("srcset") or ""
        if ss:
            for part in ss.split(","):
                u = part.strip().split(" ")[0]
                if u: urls.append(u)

    # 2) click từng thumb => đọc ảnh lớn
    for i in range(min(tcount, 50)):
        try:
            btn = page.locator("[data-gallery-thumbnail]").nth(i)
            btn.click()
            page.wait_for_timeout(250)
            main = page.locator("img[data-rimg-template], .product-gallery--main img, .product__media img")
            if main and main.count()>0:
                node = main.first
                for attr in ("src","data-src","data-rimg-template"):
                    v = node.get_attribute(attr)
                    if v: urls.append(v)
                ss = node.get_attribute("srcset") or ""
                if ss:
                    for part in ss.split(","):
                        u = part.strip().split(" ")[0]
                        if u: urls.append(u)
        except Exception:
            pass

    # 3) normalize protocol + upgrade size
    normed: list[str] = []
    for u in urls:
        if not u: 
            continue
        if u.startswith("//"):
            u = "https:" + u
        # Shopify size upgrades (jpg & png & webp & jpeg)
        u = re.sub(r"_(\d+x\d+)(_crop_center)?\.(jpg|png|jpeg|webp)$", r"_1200x1200.\3", u)
        # Nếu là rimg-template có {size}
        u = u.replace("{size}", "1200x1200")
        normed.append(u)

    # unique giữ thứ tự
    seen=set(); out=[]
    for u in normed:
        if u not in seen:
            seen.add(u); out.append(u)
    return out

def filename_from_cd(cd: str | None) -> Optional[str]:
    if not cd:
        return None
    m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^\";]+)"?', cd, flags=re.IGNORECASE)
    if m:
        name = unquote(m.group(1))
        name = re.sub(r"[^\w\.\-]+", "_", name)
        return name
    return None

def safe_name_from_url(url: str) -> str:
    parsed = urlparse(url)
    fname = os.path.basename(parsed.path)
    if not fname:
        q = parse_qs(parsed.query)
        cand = (q.get("filename") or ["image"])[0]
        fname = unquote(cand)
    fname = re.sub(r"[^\w\.\-]+", "_", fname)
    if not fname or fname == "_":
        h = hashlib.md5(url.encode("utf-8")).hexdigest()[:8]
        fname = f"img_{h}.jpg"
    return fname

def session_from_context_cookies(context: BrowserContext) -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": UA, "Accept": "*/*", "Accept-Language": "vi,en;q=0.9"})
    try:
        cookies = context.cookies()
        jar = requests.cookies.RequestsCookieJar()
        for c in cookies:
            if "npower.com.vn" in c.get("domain",""):
                jar.set(c["name"], c["value"], domain=c.get("domain"), path=c.get("path","/"))
        s.cookies = jar
    except Exception:
        pass
    return s

def try_download(session: requests.Session, url: str, dest: Path, referer: Optional[str]=None, retries:int=2, timeout:int=30) -> bool:
    headers = {"User-Agent": UA, "Accept": "*/*"}
    if referer:
        headers["Referer"] = referer
    for attempt in range(1, retries+1):
        try:
            with session.get(url, stream=True, timeout=timeout, headers=headers, allow_redirects=True) as r:
                if r.status_code != 200:
                    continue
                tmp = dest.with_suffix(dest.suffix + ".part")
                with open(tmp, "wb") as f:
                    for chunk in r.iter_content(chunk_size=1<<14):
                        if chunk:
                            f.write(chunk)
                tmp.rename(dest)
                return True
        except Exception as e:
            if attempt >= retries:
                log(f"Tải lỗi ({e}): {url}")
    return False

def process_one(context: BrowserContext, part: str, outdir: Path, delay: float, retries:int, timeout:int) -> dict:
    page = context.new_page()
    page.set_default_timeout(15000)
    res = {
        "Part": part,
        "status": "",
        "product_url": "",
        "image_files": "",
        "errors": ""
    }
    errs: list[str] = []
    try:
        open_home(page)
        search_part(page, part)
        card = pick_best_card(page, part)
        if not card:
            res["status"] = "not_found_in_search"
            return res
        if not open_card(page, card):
            res["status"] = "open_card_failed"
            return res
        if not confirm_pdp(page, part):
            errs.append("PDP không xác nhận khớp mã")
        res["product_url"] = page.url

        # Thu URL ảnh
        urls = collect_gallery_urls(page)
        if not urls:
            errs.append("Không thu được URL ảnh từ gallery")
        # Tạo session từ cookie context (tránh hotlink protection)
        session = session_from_context_cookies(context)

        folder = outdir / slugify(part)
        ensure_dir(folder)
        saved = []
        for idx, u in enumerate(urls, start=1):
            name = safe_name_from_url(u)
            # Ép đuôi về .jpg nếu có dạng không có đuôi
            if not re.search(r"\.(jpg|jpeg|png|webp)$", name, re.IGNORECASE):
                name += ".jpg"
            dest = folder / f"{idx:02d}_{name}"
            ok = try_download(session, u, dest, referer=page.url, retries=retries, timeout=timeout)
            if not ok:
                # fallback: giảm kích thước nếu 1200 không có
                u2 = re.sub(r"_1200x1200\.", "_1024x1024.", u)
                if u2 != u:
                    ok = try_download(session, u2, dest, referer=page.url, retries=retries, timeout=timeout)
            if ok:
                saved.append(str(dest))
        res["image_files"] = "|".join(saved)
        res["status"] = "ok" if saved else "opened_no_image"
        res["errors"] = "; ".join(errs)
        time.sleep(max(0.0, delay))
        return res
    except Exception as e:
        res["status"] = "error"
        res["errors"] = "; ".join(errs + [str(e)])
        return res
    finally:
        try:
            page.close()
        except Exception:
            pass

def run(excel: Optional[str], sheet: Optional[str], outdir: str, headless: bool, rate_wait: float, start:int, limit:int,
        retries:int, timeout:int, proxy: Optional[str]) -> None:
    # Nếu không có --excel, mở hộp thoại chọn file
    if not excel:
        picked = pick_excel_gui()
        if not picked:
            print("Bạn chưa chọn file Excel. Thoát.")
            sys.exit(2)
        excel = picked

    df = read_parts_from_excel(excel, sheet)
    parts = df["Part"].tolist()
    if start or limit:
        end = start + limit if limit>0 else len(parts)
        parts = parts[start:end]

    out = Path(outdir or DEFAULT_OUTDIR)
    ensure_dir(out)

    rows: list[dict] = []
    with sync_playwright() as p:
        launch_args = [
            "--disable-features=IsolateOrigins,site-per-process,AudioServiceOutOfProcess",
            "--disable-renderer-backgrounding", "--disable-backgrounding-occluded-windows",
            "--disable-background-timer-throttling", "--disable-dev-shm-usage",
        ]
        browser = p.chromium.launch(headless=headless, args=launch_args)
        ctx_kwargs = dict(locale="vi-VN")
        if proxy:
            ctx_kwargs["proxy"] = {"server": proxy}
        context: BrowserContext = browser.new_context(**ctx_kwargs)

        # chặn analytics để nhanh hơn (không chặn CDN ảnh)
        from fnmatch import fnmatch
        BLOCK = ["*googletagmanager.com/*","*google-analytics.com/*","*doubleclick.net/*","*facebook.net/*","*hotjar.com/*","*optimizely.com/*","*segment.com/*","*fullstory.com/*"]
        def _should_block(u: str)->bool: return any(fnmatch(u, pat) for pat in BLOCK)
        try:
            context.route("**/*", lambda route: route.abort() if _should_block(route.request.url) else route.continue_())
        except Exception:
            pass

        try:
            for i, part in enumerate(parts):
                log(f"== [{i+1}/{len(parts)}] {part} ==")
                r = process_one(context, part, out, rate_wait, retries, timeout)
                rows.append(r)
        finally:
            try: context.close()
            except Exception: pass
            try: browser.close()
            except Exception: pass

    if rows:
        csv_path = out / "report.csv"
        xlsx_path = out / "report.xlsx"
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            writer.writeheader()
            for row in rows:
                writer.writerow(row)
        try:
            pd.DataFrame(rows).to_excel(xlsx_path, index=False)
        except Exception:
            pass
    log("DONE.")

def main():
    ap = argparse.ArgumentParser(description="Crawler ảnh cho npower.com.vn")
    ap.add_argument("--excel", required=False, help="Đường dẫn file Excel có cột Part/SKU. Bỏ trống để chọn bằng hộp thoại.")
    ap.add_argument("--sheet", default=None, help="Tên sheet (tuỳ chọn).")
    ap.add_argument("--outdir", default=DEFAULT_OUTDIR, help=f"Thư mục output (mặc định: {DEFAULT_OUTDIR}).")
    ap.add_argument("--rate-wait", type=float, default=DEFAULT_RATEWAIT, help="Nghỉ giữa các Part (giây).")
    ap.add_argument("--headless", default=str(DEFAULT_HEADLESS).lower(), help="true/false.")
    ap.add_argument("--start", type=int, default=0, help="Bắt đầu từ index (0-based).")
    ap.add_argument("--limit", type=int, default=0, help="Số lượng cần chạy (0 = tất cả).")
    ap.add_argument("--retries", type=int, default=DEFAULT_RETRIES, help="Số lần retry tải ảnh.")
    ap.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help="Timeout (giây) mỗi request ảnh.")
    ap.add_argument("--proxy", default=None, help="VD: http://host:port hoặc http://user:pass@host:port")
    args = ap.parse_args()

    run(
        excel=args.excel,
        sheet=args.sheet,
        outdir=args.outdir,
        headless=parse_bool(args.headless),
        rate_wait=args.rate_wait,
        start=args.start,
        limit=args.limit,
        retries=max(1, args.retries),
        timeout=max(10, args.timeout),
        proxy=args.proxy,
    )

if __name__ == "__main__":
    main()
