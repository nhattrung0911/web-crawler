#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
taianhnew_nameonly.py — Lấy TÊN SẢN PHẨM nhanh & chính xác

Chức năng:
- Đọc danh sách Part từ file Excel (tự bật hộp thoại nếu không truyền tham số)
- Tìm sản phẩm trên trang se.com/vn/vi
- Mở đúng trang PDP
- Lấy tên sản phẩm từ phần tử qds-title.main-product-info__description
- Xuất ra Excel kết quả (Part, product_name, product_url, status, note)

Yêu cầu:
  pip install playwright pandas openpyxl requests
  playwright install
"""

from __future__ import annotations
import os, re, sys, time, hashlib
from pathlib import Path
from dataclasses import dataclass, asdict
from urllib.parse import urljoin
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError, Page

# ================== CONFIG ==================
BASE = "https://www.se.com/vn/vi/"
OUTDIR = "out_name"
DEFAULT_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
}

# ================== UTILS ==================
def log(msg): print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)
def ensure_dir(p: Path): p.mkdir(parents=True, exist_ok=True)
def normalize_part(s: str) -> str:
    return re.sub(r"[\s\-\._]+", "", str(s or "").strip().upper())

def pick_excel_file():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw()
        path = filedialog.askopenfilename(
            title="Chọn file Excel danh sách Part",
            filetypes=[("Excel files","*.xlsx *.xls")]
        )
        root.update(); root.destroy()
        return path if path else None
    except Exception:
        return None

def read_parts_from_excel(xlsx_path: str):
    df = pd.read_excel(xlsx_path)
    lower = {c.lower(): c for c in df.columns}
    part_col = None
    for cand in ["part","part number","mã","sku","code"]:
        if cand in lower:
            part_col = lower[cand]; break
    if not part_col:
        for c in df.columns:
            if any(w in c.lower() for w in ["part","sku","mã","code"]):
                part_col = c; break
    if not part_col:
        raise ValueError("Không tìm thấy cột chứa Part trong Excel.")
    url_col = next((lower[c] for c in lower if c in ("product_url","url","link")), None)
    df = pd.DataFrame({
        "Part": df[part_col].astype(str).str.strip(),
        "product_url": df[url_col].astype(str).str.strip() if url_col else ""
    })
    df = df[df["Part"].str.len() > 0].reset_index(drop=True)
    return df

# ================== PLAYWRIGHT CORE ==================
def wait_dom_ready(page: Page):
    page.wait_for_load_state("domcontentloaded", timeout=20000)
    page.wait_for_timeout(300)
    try: page.wait_for_load_state("networkidle", timeout=8000)
    except PlaywrightTimeoutError: pass

def perform_search(page: Page, part: str):
    page.goto(BASE, wait_until="domcontentloaded", timeout=40000)
    box = page.locator("input#uiaas-search-input.search-bar__input, input.search-bar__input, input[type='search']").first
    box.fill(part)
    try:
        btn = page.locator("button#submitSearchButton, button.search-bar__button")
        if btn.count(): btn.first.click()
        else: box.press("Enter")
    except Exception:
        box.press("Enter")
    try: page.wait_for_url(re.compile(r"/search/\?q="), timeout=10000)
    except PlaywrightTimeoutError: pass
    page.wait_for_selector(".result-card__sections, .no-results", timeout=15000)

def pick_result_card(page: Page, part: str):
    target = normalize_part(part)
    cards = page.locator(".result-card__sections")
    n = cards.count()
    if n == 0: return None
    for i in range(n):
        try:
            sku = cards.nth(i).locator(".result-card__sku .highlight").inner_text(timeout=300)
            if normalize_part(sku) == target:
                return cards.nth(i)
        except Exception: pass
    return cards.first

def open_card(page: Page, card) -> bool:
    try:
        with page.expect_navigation(timeout=15000):
            card.locator("a, qds-standalone-link.result-card__title-standalone-link").first.click()
        return "/product/" in page.url
    except Exception:
        return "/product/" in page.url

def search_and_open_pdp(page: Page, part: str) -> bool:
    perform_search(page, part)
    card = pick_result_card(page, part)
    if not card: return False
    return open_card(page, card)

# ================== EXTRACT NAME ==================
def extract_product_name(page: Page) -> str:
    try:
        node = page.locator("qds-title.main-product-info__description")
        if node.count() == 0: return ""
        txt = node.first.inner_text(timeout=3000).strip()
        if not txt:
            txt = node.first.evaluate("(el)=>el.textContent||''").strip()
        return txt
    except Exception:
        return ""

# ================== DATA CLASS ==================
@dataclass
class PartResult:
    Part: str
    product_name: str
    product_url: str
    status: str
    note: str

# ================== MAIN ==================
def main():
    excel_path = pick_excel_file()
    if not excel_path:
        print("Chưa chọn file Excel. Thoát.")
        sys.exit(1)

    df = read_parts_from_excel(excel_path)
    outdir = Path(OUTDIR); ensure_dir(outdir)
    results: list[PartResult] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(locale="vi-VN")
        page = context.new_page()

        for idx, row in df.iterrows():
            part = str(row["Part"]).strip()
            hint = str(row.get("product_url") or "")
            log(f"\n({idx+1}/{len(df)}) → Xử lý Part: {part}")
            url = ""; pname = ""; status = "ok"; note = ""
            try:
                if hint and hint.startswith("http"):
                    page.goto(hint, wait_until="domcontentloaded", timeout=30000)
                    wait_dom_ready(page)
                    url = page.url
                else:
                    ok = search_and_open_pdp(page, part)
                    if not ok:
                        status = "not_found"; note = "Không tìm thấy PDP"
                        results.append(PartResult(part,"","",status,note))
                        continue
                    url = page.url

                pname = extract_product_name(page)
                if not pname:
                    status = "no_name"; note = "Không đọc được tên"
                log(f" → {pname or '(không có)'}")
                results.append(PartResult(part,pname,url,status,note))
            except Exception as e:
                results.append(PartResult(part,"",url,"error",str(e)))
                continue

    # Lưu Excel
    out_xlsx = outdir / f"product_names_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
    rows = [asdict(r) for r in results]
    pd.DataFrame(rows).to_excel(out_xlsx, index=False)
    log(f"\n✅ Hoàn tất. File kết quả: {out_xlsx}")

if __name__ == "__main__":
    main()
