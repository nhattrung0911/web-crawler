"""
emin.vn — thiet bi do/dung cu cong nghiep crawler (httpx + selectolax.lexbor, NO Selenium).

RECON (2026-08-05)
-------------------
API      : KHONG co JSON API rieng. `?view=json` tren /search bi lo — van tra text/html
           (khong phai Sapo). Site render HTML server-side => httpx + selectolax du dung.
robots   : "Allow: /" (chi Disallow /x6tr5dtx-*). Sitemap khai bao trong robots.txt.
sitemap  : https://emin.vn/sitemap.xml (index) -> .../s/product/vi/sitemap.xml = **INDEX
           CUA INDEX**: 2156 file sitemap-N.xml, MOI file ~4557 URL /pr.html. Catalog
           UOC TINH ~9.8 TRIEU SAN PHAM (id tang dan xuyen file — nhieu kha nang site
           gom/dropship tu nhieu hang: search "MT-1220" tra ca Omron/Panasonic/Molex
           trung so). => mode "all" KHONG duoc quet het, PHAI cap cung (MAX_ITEMS/--limit),
           chi lay vai file sitemap-N dau cho toi khi du so luong can.
slug     : /<slug-tu-do>/pr.html — slug KHONG map 1-1 ro rang tu SKU (vd MT-1220 ->
           /proskitmt1220-dong-ho-do-proskit-mt-1220-88153/pr.html) => KHONG the build_url
           truc tiep tu ma nhu proskit.com. Ma sai/khong ton tai -> 404 THAT (khong sneaky
           redirect ve trang chu).
search   : GET /search?q=<tu khoa> -> HTML, KHONG PHAN TRANG THAT (page=1 va page=2 tra
           y HET nhau, da test) — cap ~16-20 ket qua/query. Nhieu ket qua KHONG khop ma
           that (trung substring so, vd "MT-1220" tra ca "D4N-1220", "ML-1220-BE") => PHAI
           mo hit that so sanh Model/SKU that tren trang detail, KHONG duoc tin mu ket qua
           dau (xem AGENTS.md §4.3, §10 bay #5).
detail   : - meta[property=og:title] / og:image
           - div label hang/model/xuat xu/bao hanh (class grid 2 cot, span.text-slate-500
             = nhan, span.text-slate-900 = gia tri)
           - gia: div.p-3/p-4 chua "chua bao gom vat" / "da bao gom vat" + .text-2xl
           - ton kho: div badge px-2 py-1 uppercase (mau emerald/amber/rose) + so trong text
           - thong so: selector cu the trong tab "San pham thuyet minh", fallback <table>
           - breadcrumb: nav ol li
category : 1 link danh muc/hang (?page=N) -> quet SACH toan bo category (stuck-page:
           dung khi trang khong sinh link moi). Day la kieu nhap RIENG cua site nay
           (khac cac site khac trong workspace).

Run
---
  emin.bat                                          -> GUI
  python emin.py --mode code --code MT-1220
  python emin.py --mode keyword --keyword "dong ho van nang"
  python emin.py --mode link --link https://emin.vn/.../pr.html
  python emin.py --mode excel --excel parts.xlsx --sheet Sheet1 --col SKU
  python emin.py --mode category --link "https://emin.vn/thiet-bi-do.../pc.html"
  python emin.py --mode all --limit 200                -> quet vai file sitemap dau (CO CAP)
"""
from __future__ import annotations

import argparse
import asyncio
import gc
import json
import os
import pathlib
import queue
import re
import threading
import time
from typing import Any
from urllib.parse import quote

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser

HERE = pathlib.Path(__file__).parent
BASE = "https://emin.vn"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")

SITEMAP_INDEX = f"{BASE}/sitemap.xml"
CIRCUIT_BREAK = 25        # N loi LIEN TIEP -> nghi bi chan IP -> dung han
MAX_ITEMS = 5000          # tran work-list
MAX_PAGES = 600           # tran phan trang category/search
MAX_SITEMAP_FILES = 20    # tran so file sitemap-N.xml doc trong mode "all" (catalog ~9.8tr, KHONG quet het)
MAX_REDIRECTS = 5
TIMEOUT = 30.0
GC_EVERY = 500
SEARCH_CANDIDATES = 3     # so ket qua search mo ra de so khop ma CHINH XAC truoc khi fallback


# ================================================================ CORE (khong dung GUI)
def norm_code(s: Any) -> str:
    """Chuan hoa ma de so khop. URL thi GIU NGUYEN hoa/thuong (path phan biet chu)."""
    v = re.sub(r"\s+", "", str(s or ""))
    return v if v.lower().startswith("http") else v.upper()


def norm_alnum(s: str) -> str:
    """Chi giu chu+so, viet thuong — dung de so khop tho trong slug (vd 'MT-1220' -> 'mt1220')."""
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def search_url(key: str) -> str:
    return f"{BASE}/search?q={quote(key)}"


def parse_sitemap(xml: str) -> list[str]:
    return re.findall(r"<loc>([^<]+)</loc>", xml)


def parse_search(html: str) -> list[str]:
    """Link san pham trong trang search — giu thu tu, bo trung."""
    t = LexborHTMLParser(html)
    seen: set[str] = set()
    out: list[str] = []
    for a in t.css('a[href*="/pr.html"]'):
        href = a.attributes.get("href") or ""
        if href.startswith("/"):
            href = BASE + href
        href = href.split("?")[0].split("#")[0]
        if href.startswith(BASE) and href not in seen:
            seen.add(href)
            out.append(href)
    return out


def parse_page(html: str, url: str) -> dict | None:
    """HTML trang detail -> 1 record (1 trang = 1 SKU)."""
    t = LexborHTMLParser(html)

    title_meta = t.css_first('meta[property="og:title"]')
    name = (title_meta.attributes.get("content") or "").strip() if title_meta else ""
    if not name:
        h1 = t.css_first("h1")
        name = h1.text(strip=True) if h1 else ""

    img_meta = t.css_first('meta[property="og:image"]')
    img_url = (img_meta.attributes.get("content") or "").strip() if img_meta else ""

    if t.body is None:
        return None
    text_content = t.body.text()

    price_no_vat = price_vat = currency = vat_percent = ""
    for div in t.css("div.p-3, div.p-4"):
        low = div.text(strip=True).lower()
        if "chưa bao gồm vat" in low:
            el = div.css_first(".text-2xl")
            if el:
                m = re.search(r"(\d{1,3}(?:,\d{3})+)\s*(VND|VNĐ|₫|USD)", el.text(strip=True), re.I)
                if m:
                    price_no_vat, currency = m.group(1).replace(",", ""), m.group(2).upper()
        elif "đã bao gồm vat" in low:
            el = div.css_first(".text-2xl")
            if el:
                m = re.search(r"(\d{1,3}(?:,\d{3})+)\s*(VND|VNĐ|₫|USD)", el.text(strip=True), re.I)
                if m:
                    price_vat = m.group(1).replace(",", "")
                    if not currency:
                        currency = m.group(2).upper()
    if currency in ("VNĐ", "₫"):
        currency = "VND"
    if price_no_vat and price_vat:
        try:
            p_no, p_vat = float(price_no_vat), float(price_vat)
            if p_no > 0:
                vat_percent = f"{round(((p_vat - p_no) / p_no) * 100)}%"
        except Exception:
            pass
    elif not price_no_vat and not price_vat:
        m = re.search(r"(\d{1,3}(?:,\d{3})+)\s*(VND|VNĐ|₫|USD)", text_content, re.I)
        if m:
            price_vat, currency = m.group(1).replace(",", ""), m.group(2).upper()
            if currency in ("VNĐ", "₫"):
                currency = "VND"

    brand = model = origin = warranty = ""
    for div in t.css("div"):
        cls = div.attributes.get("class", "") or ""
        if "md:grid-cols-2" in cls and "bg-white" in cls and "p-4" in cls:
            for col in div.css("div.flex-col"):
                lbl_el = col.css_first("span.text-slate-500")
                val_el = col.css_first("span.text-slate-900")
                if lbl_el and val_el:
                    lbl = lbl_el.text(strip=True).lower()
                    val = re.sub(r"\(Xem thêm.*?\)", "", val_el.text(separator=" ", strip=True)).strip()
                    if "hãng" in lbl:
                        brand = val
                    elif "model" in lbl:
                        model = val
                    elif "xuất" in lbl:
                        origin = val
                    elif "bảo hành" in lbl:
                        warranty = val
            break
    if not brand:
        a = t.css_first('a[href*="/pa.html"]')
        if a:
            brand = a.text(strip=True)
    if not warranty:
        m = re.search(r"Bảo hành[:\s]+([^\n<,]+)", text_content, re.I)
        if m:
            warranty = m.group(1).strip()

    stock_qty = ""
    for div in t.css("div"):
        cls = div.attributes.get("class", "") or ""
        if "px-2 py-1" in cls and "uppercase tracking-wide" in cls and (
                "bg-emerald-100" in cls or "bg-amber-100" in cls or "bg-rose-100" in cls):
            m = re.search(r"(\d+)", div.text(separator=" ", strip=True))
            if m:
                stock_qty = m.group(1)
            break

    spec_str = ""
    spec_sel = ('body > div.flex.flex-col.min-h-screen > main > div > '
                'div.bg-white.border.border-slate-200.rounded-xl.shadow-sm.mb-2 > '
                'div.p-4 > div.relative.overflow-hidden.ease-in-out > div')
    spec_el = t.css_first(spec_sel)
    if spec_el:
        frag = (spec_el.html or "").replace("<br>", "\n").replace("<br/>", "\n").replace("</p>", "\n</p>")
        body = LexborHTMLParser(frag).body
        if body:
            lines = [ln.strip() for ln in body.text(separator="\n", strip=True).split("\n") if ln.strip()]
            spec_str = "\n".join(lines)
    if not spec_str:
        specs = []
        table = t.css_first("table")
        if table:
            for tr in table.css("tr"):
                tds = tr.css("td, th")
                if len(tds) >= 2:
                    specs.append(f"{tds[0].text(strip=True)}: {tds[1].text(strip=True)}")
        spec_str = "\n".join(specs)

    cats: list[str] = []
    nav = t.css_first("nav ol")
    if nav:
        cats = [li.text(strip=True) for li in nav.css("li") if li.text(strip=True)]

    sku = model
    if not sku:
        m = re.search(r"([a-zA-Z0-9\-.]+)", name)
        sku = m.group(1) if m else ""
    if not sku and not name:
        return None

    return {"Input": "", "Item No": norm_code(sku), "Page": name, "Status": "OK",
            "_cats": cats, "Thuong hieu": brand.split("(")[0].strip(), "Xuat xu": origin,
            "Gia chua VAT": price_no_vat, "Gia co VAT": price_vat, "VAT (%)": vat_percent,
            "Don vi": currency, "Bao hanh": warranty, "Ton kho": stock_qty,
            "Thong so": spec_str, "Image": img_url, "URL": url}


def export_excel(results: list[dict], path: pathlib.Path) -> pathlib.Path:
    """Thong so da la 1 chuoi text san (khong phai dict roi hoc) -> ghi thang 1 COT.
    Xem AGENTS.md §4.6."""
    from openpyxl.utils import get_column_letter

    max_cat = max((len(r.get("_cats") or []) for r in results), default=0)
    fixed_head = ["Input", "Item No", "Page", "Status"]
    fixed_tail = ["Thuong hieu", "Xuat xu", "Gia chua VAT", "Gia co VAT", "VAT (%)", "Don vi",
                  "Bao hanh", "Ton kho", "Thong so", "Image", "URL"]
    rows = []
    for r in results:
        row = {k: r.get(k, "") for k in fixed_head}
        cats = r.get("_cats") or []
        for i in range(max_cat):
            row[f"Category {i + 1}"] = cats[i] if i < len(cats) else ""
        row.update({k: r.get(k, "") for k in fixed_tail})
        rows.append(row)
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, index=False)
        ws = xw.sheets["Sheet1"]
        col = get_column_letter(df.columns.get_loc("Thong so") + 1)
        ws.column_dimensions[col].width = 60
        for c in ws[col][1:]:
            c.alignment = c.alignment.copy(wrap_text=True, vertical="top")
    return path


class Engine:
    """Crawler core — chay headless/CLI duoc, khong phu thuoc GUI."""

    def __init__(self, conc: int = 10, delay: float = 0.05, log=print,
                 stop_evt: threading.Event | None = None,
                 pause_evt: threading.Event | None = None) -> None:
        self.conc, self.delay, self.log = max(1, conc), max(0.0, delay), log
        self.stop_evt = stop_evt or threading.Event()
        self.pause_evt = pause_evt or threading.Event()
        self.fails = 0
        self.results: list[dict] = []
        self.done = 0
        self.total = 0
        self.key = ""
        self._ckpt_loaded = False

    # ---- resume: JSONL sidecar (append tung item; chay lai = doc lai, bo qua trang da xong)
    def ckpt(self, key: str) -> pathlib.Path:
        return HERE / f"_raw_{re.sub(r'[^A-Za-z0-9]+', '_', key)[:40]}.jsonl"

    def _load_ckpt(self, key: str) -> set[str]:
        p = self.ckpt(key)
        if self._ckpt_loaded:
            return {r.get("URL", "") for r in self.results}
        self._ckpt_loaded = True
        seen: set[str] = set()
        if not p.exists():
            return seen
        for line in p.read_text("utf-8", errors="ignore").splitlines():
            try:
                r = json.loads(line)
            except Exception:
                continue
            self.results.append(r)
            seen.add(r.get("URL", ""))
        if seen:
            self.log(f"[SYS] Resume: {len(self.results)} record da xong -> bo qua")
        return seen

    def _save(self, key: str, recs: list[dict]) -> None:
        with self.ckpt(key).open("a", encoding="utf-8") as f:
            for r in recs:
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
            f.flush()

    def cleanup(self) -> None:
        """Xong sach -> xoa checkpoint. Bi Stop/crash -> GIU lai de lan sau resume."""
        if not self.key:
            return
        p = self.ckpt(self.key)
        if p.exists() and not self.stop_evt.is_set():
            p.unlink()

    async def _gate(self) -> bool:
        while self.pause_evt.is_set() and not self.stop_evt.is_set():
            await asyncio.sleep(0.2)
        return not self.stop_evt.is_set()

    async def fetch(self, cli: httpx.AsyncClient, url: str, tries: int = 4) -> str | None:
        for i in range(tries):
            if not await self._gate():
                return None
            try:
                r = await cli.get(url)
                if r.status_code in (403, 429, 503):
                    wait = min(2 ** i * 5, 60)
                    self.log(f"[WARN] {r.status_code} -> ngu {wait}s | {url[-60:]}")
                    await asyncio.sleep(wait)
                    continue
                if r.status_code == 404:
                    return None
                r.raise_for_status()
                self.fails = 0
                return r.text
            except Exception as e:
                if i < tries - 1:
                    await asyncio.sleep(min(2 ** i, 20))
                    continue
                self.fails += 1
                self.log(f"[ERR] {type(e).__name__} | {url[-60:]}")
                if self.fails >= CIRCUIT_BREAK:
                    self.log(f"[STOP] {CIRCUIT_BREAK} loi LIEN TIEP - nghi bi chan IP. DUNG.")
                    self.stop_evt.set()
                return None
        return None

    async def _crawl_urls(self, cli: httpx.AsyncClient, urls: list[str], key: str,
                          inputs: dict[str, str] | None = None) -> None:
        seen = self._load_ckpt(key)
        urls = [u for u in dict.fromkeys(urls) if u not in seen][:MAX_ITEMS]
        self.total = len(urls) + len(seen)
        self.done = len(seen)
        sem = asyncio.Semaphore(self.conc)

        async def one(u: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                html = await self.fetch(cli, u)
                if self.delay:
                    await asyncio.sleep(self.delay)
                rec = parse_page(html, u) if html else None
                if rec and inputs:
                    rec["Input"] = inputs.get(u, inputs.get(rec["Item No"], ""))
                if rec:
                    self.results.append(rec)
                    self._save(key, [rec])
                self.done += 1
                self.log(f"[{'OK ' if rec else 'MISS'}] {self.done}/{self.total} | {u[-58:]}")
                if self.done % GC_EVERY == 0:
                    gc.collect()

        try:
            await asyncio.gather(*(one(u) for u in urls))
        except (KeyboardInterrupt, Exception) as e:
            self.log(f"[FATAL] {type(e).__name__}: {e}")

    async def _resolve_code(self, cli: httpx.AsyncClient, code: str) -> tuple[dict | None, bool]:
        """Search -> mo TOI DA SEARCH_CANDIDATES hit dau, so mã CHINH XAC voi Item No
        that tren trang detail (AGENTS §4.3 — khong tin mu ket qua dau). Khop -> (rec, True).
        Khong khop hit nao nhung parse duoc -> (rec dau tien, False)."""
        shtml = await self.fetch(cli, search_url(code))
        if self.delay:
            await asyncio.sleep(self.delay)
        hits = parse_search(shtml) if shtml else []
        if not hits:
            return None, False
        target = norm_alnum(code)
        hits.sort(key=lambda u: 0 if target in norm_alnum(u.rsplit("/pr.html", 1)[0]) else 1)

        fallback: dict | None = None
        for cand in hits[:SEARCH_CANDIDATES]:
            if not await self._gate():
                break
            html = await self.fetch(cli, cand)
            if self.delay:
                await asyncio.sleep(self.delay)
            if not html:
                continue
            rec = parse_page(html, cand)
            if not rec:
                continue
            if fallback is None:
                fallback = rec
            if norm_code(rec["Item No"]) == code:
                return rec, True
        return fallback, False

    async def _crawl_codes(self, cli: httpx.AsyncClient, codes: list[str], key: str) -> None:
        seen = self._load_ckpt(key)
        have = {r.get("Input") for r in self.results if r.get("Input")}
        todo = [c for c in dict.fromkeys(codes) if c not in have][:MAX_ITEMS]
        self.total = len(todo) + len(have)
        self.done = len(have)
        sem = asyncio.Semaphore(self.conc)

        async def one(code: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                rec, exact = await self._resolve_code(cli, code)
                self.done += 1
                if not rec:
                    self.log(f"[MISS] {self.done}/{self.total} khong tim thay: {code}")
                    return
                rec["Input"] = code
                if not exact:
                    rec["Status"] = "OK (khong khop chinh xac ma)"
                self.results.append(rec)
                self._save(key, [rec])
                self.log(f"[{'OK ' if exact else 'HIT'}] {self.done}/{self.total} {code} -> {rec['Item No']}")
                if self.done % GC_EVERY == 0:
                    gc.collect()

        try:
            await asyncio.gather(*(one(c) for c in todo))
        except (KeyboardInterrupt, Exception) as e:
            self.log(f"[FATAL] {type(e).__name__}: {e}")

        have = {r.get("Input") for r in self.results}
        for c in codes:                      # ma miss van co dong Not Found
            if c not in have:
                self.results.append({"Input": c, "Item No": c, "Page": "", "Status": "Not Found",
                                     "_cats": [], "Thuong hieu": "", "Xuat xu": "", "Gia chua VAT": "",
                                     "Gia co VAT": "", "VAT (%)": "", "Don vi": "", "Bao hanh": "",
                                     "Ton kho": "", "Thong so": "", "Image": "", "URL": ""})

    async def _crawl_category(self, cli: httpx.AsyncClient, cat_url: str, key: str) -> None:
        """1 link danh muc/hang -> phan trang ?page=N -> quet SACH (stuck-page = dung)."""
        base_url = cat_url.split("?")[0].strip()
        urls: list[str] = []
        page = 1
        while page <= MAX_PAGES:
            if not await self._gate():
                break
            html = await self.fetch(cli, f"{base_url}?page={page}")
            new = [u for u in (parse_search(html) if html else []) if u not in urls]
            if not new:
                break
            urls += new
            self.log(f"[SYS] Danh muc trang {page} -> {len(urls)} SP")
            page += 1
        await self._crawl_urls(cli, urls[:MAX_ITEMS], key)

    async def _crawl_all(self, cli: httpx.AsyncClient, key: str, limit: int) -> None:
        """Catalog UOC ~9.8 trieu SP (xem docstring) -> CAP CUNG, chi doc vai file
        sitemap-N.xml dau cho toi khi du so luong can (MAX_SITEMAP_FILES / cap)."""
        cap = limit or MAX_ITEMS
        idx_xml = await self.fetch(cli, SITEMAP_INDEX)
        top_locs = parse_sitemap(idx_xml) if idx_xml else []
        prod_index = next((u for u in top_locs if "/product/" in u and u.endswith("sitemap.xml")), None)
        if not prod_index:
            self.log("[ERR] Khong tim thay sitemap san pham trong sitemap.xml")
            return
        sub_xml = await self.fetch(cli, prod_index)
        sub_sitemaps = parse_sitemap(sub_xml) if sub_xml else []
        self.log(f"[SYS] Canh bao: catalog rat lon ({len(sub_sitemaps)} file sitemap con). "
                 f"Mode ALL chi lay toi da {cap} SP (dung --limit de chinh).")

        urls: list[str] = []
        for i, sm in enumerate(sub_sitemaps[:MAX_SITEMAP_FILES], 1):
            if not await self._gate() or len(urls) >= cap:
                break
            xml = await self.fetch(cli, sm)
            got = [u for u in (parse_sitemap(xml) if xml else []) if u.endswith("/pr.html")]
            urls += got
            self.log(f"[SYS] sitemap-{i} -> +{len(got)} (tong {len(urls)}/{cap})")
        await self._crawl_urls(cli, urls[:cap], key)

    async def run(self, mode: str, target: str = "", excel: str = "", sheet: str = "",
                  col: str = "", limit: int = 0) -> list[dict]:
        headers = {"User-Agent": UA, "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8"}
        key = self.key = f"{mode}_{target or excel}_{col}"
        async with httpx.AsyncClient(headers=headers, timeout=TIMEOUT, follow_redirects=True,
                                     max_redirects=MAX_REDIRECTS,
                                     limits=httpx.Limits(max_connections=self.conc + 4)) as cli:
            if mode == "all":
                await self._crawl_all(cli, key, limit)

            elif mode in ("code", "excel"):
                raw = [norm_code(target)] if mode == "code" else read_excel_col(excel, sheet, col)
                raw = [c for c in dict.fromkeys(raw) if c][:MAX_ITEMS]
                links = [c for c in raw if c.lower().startswith("http")]
                codes = [c for c in raw if c not in links]
                if links:
                    await self._crawl_urls(cli, links, key, inputs={u: u for u in links})
                if codes:
                    await self._crawl_codes(cli, codes, key)

            elif mode == "keyword":
                urls, page = [], 1
                while page <= MAX_PAGES:
                    if not await self._gate():
                        break
                    html = await self.fetch(cli, f"{search_url(target)}&page={page}")
                    new = [u for u in (parse_search(html) if html else []) if u not in urls]
                    if not new:                       # stuck-page (search khong phan trang that -> dung o page 1)
                        break
                    urls += new
                    self.log(f"[SYS] '{target}' trang {page} -> {len(urls)} URL")
                    if limit and len(urls) >= limit:
                        break
                    page += 1
                await self._crawl_urls(cli, urls[:limit or MAX_ITEMS], key)

            elif mode == "link":
                await self._crawl_urls(
                    cli, [u.strip() for u in target.splitlines() if u.strip().startswith("http")], key)

            elif mode == "category":
                await self._crawl_category(cli, target, key)

            else:
                self.log(f"[ERR] mode khong hop le: {mode}")
        return self.results


def read_excel_col(path: str, sheet: str, col: str) -> list[str]:
    """Loi -> ValueError (KHONG dung SystemExit: worker GUI bat `except Exception` se
    KHONG bat duoc SystemExit -> thread chet cam, GUI treo running=True)."""
    if not path:
        raise ValueError("Chua chon file Excel")
    df = pd.read_excel(path, sheet_name=sheet or 0, dtype=str)
    df.columns = df.columns.str.strip()
    if col not in df.columns:
        raise ValueError(f"Khong thay cot '{col}'. Cot co san: {list(df.columns)}")
    return [norm_code(v) for v in df[col].dropna().tolist()]


# ================================================================ GUI
def run_gui() -> None:
    import sys
    import tkinter as tk
    from tkinter import messagebox, ttk

    root_dir = next(p for p in (HERE, *HERE.parents) if (p / "gui_common.py").exists())
    sys.path.insert(0, str(root_dir))
    from gui_common import ask_export_path, attach_excel_picker

    root = tk.Tk()
    root.title("EMIN.VN Crawler — emin.vn")
    root.geometry("1080x680")
    ttk.Style().theme_use("clam")

    q: queue.Queue[str] = queue.Queue()
    stop_evt, pause_evt = threading.Event(), threading.Event()
    state: dict[str, Any] = {"running": False, "results": [], "eng": None}

    left = ttk.Frame(root, padding=10)
    left.pack(side="left", fill="y")
    right = ttk.Frame(root, padding=(0, 10, 10, 10))
    right.pack(side="right", fill="both", expand=True)

    mode = tk.StringVar(value="code")
    ttk.Label(left, text="Pipeline", font=("Segoe UI", 10, "bold")).pack(anchor="w")
    for val, txt in [("code", "1. Theo MA"), ("keyword", "2. Theo TU KHOA"),
                     ("link", "3. Theo LINK (moi dong 1 link)"), ("excel", "4. Theo cot EXCEL"),
                     ("category", "5. Theo LINK DANH MUC/HANG (quet sach)"),
                     ("all", "6. Crawl ALL (sitemap, CO CAP — catalog rat lon)")]:
        ttk.Radiobutton(left, text=txt, value=val, variable=mode).pack(anchor="w")

    ttk.Label(left, text="\nMa / Tu khoa / Link danh muc:").pack(anchor="w")
    ent = ttk.Entry(left, width=36)
    ent.pack(anchor="w")
    ttk.Label(left, text="Link (moi dong 1 link):").pack(anchor="w")
    box = tk.Text(left, width=36, height=5)
    box.pack(anchor="w")

    xls, sheet, col = tk.StringVar(), tk.StringVar(), tk.StringVar()
    outdir = tk.StringVar(value="images")
    attach_excel_picker(left, xls, sheet, col, outdir)

    conc, delay, limit = tk.IntVar(value=10), tk.DoubleVar(value=0.1), tk.IntVar(value=0)
    for lbl, var in [("Concurrency:", conc), ("Delay (s):", delay), ("Limit (0=all):", limit)]:
        f3 = ttk.Frame(left)
        f3.pack(anchor="w", pady=1)
        ttk.Label(f3, text=lbl, width=13).pack(side="left")
        ttk.Entry(f3, textvariable=var, width=8).pack(side="left")

    console = tk.Text(right, bg="#12161c", fg="#d6e2ee", insertbackground="#fff", wrap="none")
    console.pack(fill="both", expand=True)
    pb = ttk.Progressbar(right, mode="determinate")
    pb.pack(fill="x", pady=4)
    stat = ttk.Label(right, text="0/0")
    stat.pack(anchor="w")

    def drain() -> None:
        while not q.empty():
            console.insert("end", q.get() + "\n")
            console.see("end")
        eng: Engine | None = state["eng"]
        if eng and eng.total:
            pb["maximum"], pb["value"] = eng.total, eng.done
            stat.config(text=f"{eng.done}/{eng.total} | {len(eng.results)} record")
        root.after(100, drain)

    def worker() -> None:
        m = mode.get()
        target = box.get("1.0", "end").strip() if m == "link" else ent.get().strip()
        eng = Engine(conc.get(), delay.get(), q.put, stop_evt, pause_evt)
        state["eng"] = eng
        try:
            asyncio.run(eng.run(m, target, xls.get(), sheet.get(), col.get(), limit.get()))
            eng.cleanup()
        except BaseException as e:
            q.put(f"[FATAL] {type(e).__name__}: {e}")
        finally:
            state["results"] = eng.results
            state["running"] = False
            q.put(f"[SYS] XONG — {len(eng.results)} record. Bam EXPORT de xuat Excel.")

    def start() -> None:
        if state["running"]:
            messagebox.showinfo("Dang chay", "Dang crawl. Bam Stop truoc.")
            return
        m = mode.get()
        if m == "excel" and not (xls.get() and col.get()):
            messagebox.showwarning("Thieu du lieu", "Chon file Excel va cot truoc.")
            return
        if m in ("code", "keyword", "category") and not ent.get().strip():
            messagebox.showwarning("Thieu du lieu", "Nhap ma / tu khoa / link danh muc truoc.")
            return
        if m == "link" and not box.get("1.0", "end").strip():
            messagebox.showwarning("Thieu du lieu", "Dan link vao o (moi dong 1 link).")
            return
        stop_evt.clear()
        pause_evt.clear()
        state["running"] = True
        console.delete("1.0", "end")
        threading.Thread(target=worker, daemon=True).start()

    def export() -> None:
        if not state["results"]:
            messagebox.showinfo("Export", "Chua co du lieu")
            return
        p = ask_export_path(xls.get(), HERE, "eminvn.xlsx")
        if p:
            export_excel(state["results"], pathlib.Path(p))
            messagebox.showinfo("Export", f"Da luu:\n{p}")

    fb = ttk.Frame(left)
    fb.pack(anchor="w", pady=8)
    ttk.Button(fb, text="START", command=start).pack(side="left")
    ttk.Button(fb, text="Pause", command=lambda: (pause_evt.clear() if pause_evt.is_set()
                                                  else pause_evt.set())).pack(side="left")
    ttk.Button(fb, text="Stop", command=stop_evt.set).pack(side="left")
    ttk.Button(fb, text="EXPORT", command=export).pack(side="left")

    def on_close() -> None:
        if state["running"] and not messagebox.askyesno("Thoat", "Dang crawl. Thoat va tat luon?"):
            return
        stop_evt.set()
        try:
            root.destroy()
        finally:
            os._exit(0)

    root.protocol("WM_DELETE_WINDOW", on_close)
    drain()
    root.mainloop()


# ================================================================ CLI
def main() -> None:
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")   # console cp1252 -> vo dau tieng Viet
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="EMIN.VN crawler (emin.vn)")
    ap.add_argument("--mode", choices=["code", "keyword", "link", "excel", "category", "all"])
    ap.add_argument("--code")
    ap.add_argument("--keyword")
    ap.add_argument("--link")
    ap.add_argument("--excel")
    ap.add_argument("--sheet", default="")
    ap.add_argument("--col", default="")
    ap.add_argument("--out", default="eminvn.xlsx")
    ap.add_argument("--conc", type=int, default=10)
    ap.add_argument("--delay", type=float, default=0.1)
    ap.add_argument("--limit", type=int, default=0)
    a = ap.parse_args()

    if not a.mode:
        run_gui()
        return

    target = a.code or a.keyword or a.link or ""
    eng = Engine(a.conc, a.delay)
    t0 = time.time()
    try:
        asyncio.run(eng.run(a.mode, target, a.excel or "", a.sheet, a.col, a.limit))
    finally:
        if eng.results:
            p = export_excel(eng.results, pathlib.Path(a.out))
            print(f"[SYS] {len(eng.results)} record -> {p} ({time.time() - t0:.1f}s)")
        eng.cleanup()


if __name__ == "__main__":
    main()
