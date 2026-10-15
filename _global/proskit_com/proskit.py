"""
proskit.com — Pro'sKit tool catalog crawler (httpx + selectolax, NO browser, NO Selenium).

RECON (2026-08-05)
-------------------
API      : KHONG co JSON API (khong thay XHR/fetch nao tra JSON). Site render HTML
           server-side (ASP.NET) => httpx + selectolax la du.
robots   : "Allow: /". Sitemap khai bao trong robots.txt.
sitemap  : https://www.proskit.com/sitemap.xml (sitemapindex) ->
           https://www.proskit.com/sitemap/sitemap-tools.xml = **2709 URL san pham**,
           lastmod MOI (2026-07-29) -> dung TRUC TIEP cho mode "all", khoi BFS/phan trang.
slug=ma  : URL detail = /Product/<CODE> **CHINH XAC = ma san pham** (vd /Product/MT-1220).
           Case-INSENSITIVE: /Product/mt-1220 -> 302 -> /Product/MT-1220 (200).
           Ma KHONG ton tai -> KHONG 404, redirect ve /Products (trang danh sach chung, 200!)
           => PHAI kiem final URL (redirect guard): path phai bat dau "/product/" (thuong,
           khong phai "/products" so nhieu) moi coi la hit that.
detail   : - script[type=application/ld+json] @type=Product -> sku, name, image[], description,
             category (chuan, day du nhat, DUNG CAI NAY thay vi scrape DOM).
           - script[type=application/ld+json] @type=BreadcrumbList -> full category path.
           - div#productInfo (tab "San pham / San pham thuyet minh") = mo ta + anh + <table>
             thong so ky thuat dang 2 cot don gian (KHONG rowspan/colspan trong mau da xem).
           - div.productDescription-positiveNotes = bullet highlight (<br> ngan dong).
           - 1 TRANG = 1 MODEL DUY NHAT (khac toptul.com — khong co bang nhieu ma/trang).
search   : GET /SearchProducts?keyword=<tu khoa hoac ma>&page=<N> -> HTML server-side,
           ket qua <a class="proList-link" href="/Product/<CODE>">; phan trang stuck-page
           (het ket qua thi trang khong sinh link moi).
gia/ton  : KHONG co gia/ton kho cong khai tren site catalog nay (B2B, khong ban le truc tuyen)
           => KHONG crawl gia/ton, KHONG bia ra cot gia.
toc do   : Slug = ma nen **direct lookup 1 request/ma** (build_url(code) roi fetch thang),
           KHONG can search truoc nhu toptul.com (search cham hon, chi dung khi direct MISS
           do sai chinh ta / ma khong nam trong URL slug).

Run
---
  proskit.bat                                       -> GUI
  python proskit.py --mode all [--limit 50]          -> quet sach sitemap-tools.xml (~2.7k SP)
  python proskit.py --mode code --code MT-1220
  python proskit.py --mode keyword --keyword "digital multimeter"
  python proskit.py --mode link --link https://www.proskit.com/Product/MT-1220
  python proskit.py --mode excel --excel parts.xlsx --sheet Sheet1 --col Part
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
BASE = "https://www.proskit.com"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")

SITEMAP_URL = f"{BASE}/sitemap/sitemap-tools.xml"
CIRCUIT_BREAK = 25        # N loi LIEN TIEP -> nghi bi chan IP -> dung han
MAX_ITEMS = 5000          # tran work-list
MAX_PAGES = 600           # tran phan trang search
MAX_REDIRECTS = 5
TIMEOUT = 30.0
GC_EVERY = 500

PROD_PATH_RE = re.compile(r"^/product/", re.I)   # guard: /products (so nhieu) KHONG khop


# ================================================================ CORE (khong dung GUI)
def norm_code(s: Any) -> str:
    """Chuan hoa ma de so khop. URL thi GIU NGUYEN hoa/thuong (path phan biet chu)."""
    v = re.sub(r"\s+", "", str(s or ""))
    return v if v.lower().startswith("http") else v.upper()


def build_url(code: str) -> str:
    return f"{BASE}/Product/{quote(code)}"


def search_url(key: str) -> str:
    return f"{BASE}/SearchProducts?keyword={quote(key)}"


def parse_sitemap(xml: str) -> list[str]:
    return [u for u in re.findall(r"<loc>([^<]+)</loc>", xml) if "/Product/" in u]


def parse_search(html: str) -> list[str]:
    """Link san pham trong trang search — giu thu tu, bo trung."""
    t = LexborHTMLParser(html)
    seen: set[str] = set()
    out: list[str] = []
    for a in t.css("a.proList-link"):
        href = a.attributes.get("href") or ""
        if href.startswith("/"):
            href = BASE + href
        if href.startswith(BASE) and "/Product/" in href and href not in seen:
            seen.add(href)
            out.append(href)
    return out


def slug_code(url: str) -> str:
    return norm_code(url.rstrip("/").rsplit("/", 1)[-1])


def build_grid(table) -> list[list[str]]:
    """Bang HTML -> luoi text chu nhat; giai rowspan/colspan (o gop -> lap gia tri ra cac
    o con). Bang thong so proskit hay dung rowspan gop 2 ma dung chung 1 mo ta (vd PK-2091M:
    9HW-002A/9HW-002B) — khong giai se MAT ma dong sau (thieu td do bi che boi rowspan)."""
    grid: list[list[str | None]] = []
    for r, tr in enumerate(table.css("tr")):
        while len(grid) <= r:
            grid.append([])
        col = 0
        for cell in tr.css("td, th"):
            while col < len(grid[r]) and grid[r][col] is not None:
                col += 1
            try:
                cs = max(1, int(cell.attributes.get("colspan") or 1))
                rs = max(1, int(cell.attributes.get("rowspan") or 1))
            except ValueError:
                cs = rs = 1
            txt = re.sub(r"\s+", " ", cell.text(separator=" ", strip=True) or "")
            for dr in range(rs):
                while len(grid) <= r + dr:
                    grid.append([])
                row = grid[r + dr]
                for dc in range(cs):
                    while len(row) <= col + dc:
                        row.append(None)
                    if row[col + dc] is None:
                        row[col + dc] = txt
            col += cs
    width = max((len(r) for r in grid), default=0)
    return [[(c or "") for c in r] + [""] * (width - len(r)) for r in grid]


def parse_page(html: str, url: str) -> dict | None:
    """HTML trang detail -> 1 record (1 trang = 1 model tren site nay)."""
    t = LexborHTMLParser(html)

    ld_product: dict = {}
    cats: list[str] = []
    for sc in t.css('script[type="application/ld+json"]'):
        try:
            data = json.loads(sc.text())
        except Exception:
            continue
        if not isinstance(data, dict):
            continue
        if data.get("@type") == "BreadcrumbList":
            names = [e.get("name", "") for e in data.get("itemListElement", [])]
            cats = [n for n in names[2:-1] if n]   # bo Home + "San pham tong quan" + ten SP (leaf)
        elif data.get("@type") == "Product":
            ld_product = data

    code = norm_code(ld_product.get("sku") or ld_product.get("mpn") or slug_code(url))
    name = re.sub(r"\s+", " ", str(ld_product.get("name") or "")).strip()
    if not name:
        h1 = t.css_first("h1.productDescription-name")
        name = h1.text(strip=True) if h1 else ""
    if not code and not name:
        return None

    images = [u for u in (ld_product.get("image") or []) if u]
    if not images:
        for img in t.css("div.swiper-BigImg img.swiper-slide-img"):
            src = img.attributes.get("src")
            if src and src not in images:
                images.append(src)

    desc = ""
    notes = t.css_first("div.productDescription-positiveNotes")
    if notes:
        frag = (notes.html or "").replace("<br>", "\n").replace("<br/>", "\n")
        body = LexborHTMLParser(frag).body
        if body:
            lines = [ln.strip() for ln in body.text(separator="\n", strip=True).split("\n") if ln.strip()]
            desc = "\n".join(lines)
    if not desc:
        desc = str(ld_product.get("description") or "")

    spec: dict[str, str] = {}
    info = t.css_first("div#productInfo")
    if info:
        for table in info.css("table"):
            for row in build_grid(table):
                cells = [c for c in row if c]
                if len(cells) >= 2:
                    k, v = cells[0], " ".join(cells[1:])
                    if k and v and k not in spec:
                        spec[k] = v

    return {"Input": "", "Item No": code, "Page": name, "Status": "OK",
            "_cats": cats, "_desc": desc, "_spec": spec,
            "Image": " | ".join(images[:5]), "URL": url}


def export_excel(results: list[dict], path: pathlib.Path) -> pathlib.Path:
    """Spec gop 1 COT 'Thong so' (moi dong 'Ten: Gia tri') — khong trai union key thanh
    hang tram cot rong. Xem AGENTS.md §4.6."""
    from openpyxl.utils import get_column_letter

    max_cat = max((len(r.get("_cats") or []) for r in results), default=0)
    rows = []
    for r in results:
        row = {"Input": r.get("Input", ""), "Item No": r.get("Item No", ""),
               "Page": r.get("Page", ""), "Status": r.get("Status", "")}
        cats = r.get("_cats") or []
        for i in range(max_cat):
            row[f"Category {i + 1}"] = cats[i] if i < len(cats) else ""
        row["Mo ta"] = r.get("_desc", "")
        row["Thong so"] = "\n".join(f"{k}: {v}" for k, v in (r.get("_spec") or {}).items())
        row["Image"] = r.get("Image", "")
        row["URL"] = r.get("URL", "")
        rows.append(row)
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, index=False)
        ws = xw.sheets["Sheet1"]
        for name in ("Mo ta", "Thong so"):
            col = get_column_letter(df.columns.get_loc(name) + 1)
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

    async def fetch(self, cli: httpx.AsyncClient, url: str, tries: int = 4) -> tuple[str | None, str | None]:
        """Tra (html, final_url_str). final_url dung de redirect-guard o noi goi."""
        for i in range(tries):
            if not await self._gate():
                return None, None
            try:
                r = await cli.get(url)
                if r.status_code in (403, 429, 503):
                    wait = min(2 ** i * 5, 60)
                    self.log(f"[WARN] {r.status_code} -> ngu {wait}s | {url[-60:]}")
                    await asyncio.sleep(wait)
                    continue
                if r.status_code == 404:
                    return None, None
                r.raise_for_status()
                self.fails = 0
                return r.text, str(r.url)
            except Exception as e:
                if i < tries - 1:
                    await asyncio.sleep(min(2 ** i, 20))
                    continue
                self.fails += 1
                self.log(f"[ERR] {type(e).__name__} | {url[-60:]}")
                if self.fails >= CIRCUIT_BREAK:
                    self.log(f"[STOP] {CIRCUIT_BREAK} loi LIEN TIEP - nghi bi chan IP. DUNG.")
                    self.stop_evt.set()
                return None, None
        return None, None

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
                html, final_url = await self.fetch(cli, u)
                if self.delay:
                    await asyncio.sleep(self.delay)
                rec = None
                if html and final_url and PROD_PATH_RE.match(httpx.URL(final_url).path):
                    rec = parse_page(html, final_url)
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

    async def _crawl_codes(self, cli: httpx.AsyncClient, codes: list[str], key: str) -> None:
        """Slug = ma -> direct lookup 1 request/ma (build_url). MISS (redirect ve /Products
        hoac loi) -> fallback /SearchProducts?keyword=<ma>, uu tien khop CHINH XAC slug."""
        seen = self._load_ckpt(key)
        have = {r.get("Input") for r in self.results if r.get("Input")}
        todo = [c for c in dict.fromkeys(codes) if c not in have][:MAX_ITEMS]
        self.total = len(todo) + len(have)
        self.done = len(have)
        sem = asyncio.Semaphore(self.conc)
        genuine_miss: set[str] = set()   # ma da fetch xong that (khong bi Stop giua chung) va khong co ket qua

        async def one(code: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                html, final_url = await self.fetch(cli, build_url(code))
                exact = True
                if not (html and final_url and PROD_PATH_RE.match(httpx.URL(final_url).path)):
                    # direct MISS -> fallback search, uu tien khop slug CHINH XAC (khong tin mu ket qua dau)
                    shtml, _ = await self.fetch(cli, search_url(code))
                    hits = parse_search(shtml) if shtml else []
                    pick = next((h for h in hits if slug_code(h) == code), None)
                    exact = pick is not None
                    if not pick and hits:
                        pick = hits[0]
                    if not pick:
                        self.done += 1
                        if self.stop_evt.is_set():
                            # bi Stop/circuit-breaker giua chung -> chua kiem tra that, KHONG duoc bao Not Found gia
                            self.log(f"[STOP] {self.done}/{self.total} bi ngat giua chung, chua kiem tra xong: {code}")
                        else:
                            genuine_miss.add(code)
                            self.log(f"[MISS] {self.done}/{self.total} khong tim thay: {code}")
                        return
                    html, final_url = await self.fetch(cli, pick)
                if self.delay:
                    await asyncio.sleep(self.delay)
                rec = parse_page(html, final_url) if html and final_url else None
                if not rec:
                    self.done += 1
                    if self.stop_evt.is_set():
                        self.log(f"[STOP] {self.done}/{self.total} bi ngat giua chung, chua kiem tra xong: {code}")
                    else:
                        genuine_miss.add(code)
                        self.log(f"[MISS] {self.done}/{self.total} khong parse duoc: {code}")
                    return
                rec["Input"] = code
                if not exact:
                    rec["Status"] = "OK (khong khop chinh xac ma)"
                self.results.append(rec)
                self._save(key, [rec])
                self.done += 1
                self.log(f"[{'OK ' if exact else 'HIT'}] {self.done}/{self.total} {code} -> {rec['Item No']}")
                if self.done % GC_EVERY == 0:
                    gc.collect()

        try:
            await asyncio.gather(*(one(c) for c in todo))
        except (KeyboardInterrupt, Exception) as e:
            self.log(f"[FATAL] {type(e).__name__}: {e}")

        have = {r.get("Input") for r in self.results}
        for c in codes:                      # ma miss THAT (da kiem tra) -> Not Found; ma bi ngat giua chung -> Unverified de lan sau retry
            if c not in have:
                status = "Not Found" if c in genuine_miss else "Unverified - catalog scan interrupted"
                self.results.append({"Input": c, "Item No": c, "Page": "", "Status": status,
                                     "_cats": [], "_desc": "", "_spec": {}, "Image": "", "URL": ""})

    async def run(self, mode: str, target: str = "", excel: str = "", sheet: str = "",
                  col: str = "", limit: int = 0) -> list[dict]:
        headers = {"User-Agent": UA, "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8"}
        key = self.key = f"{mode}_{target or excel}_{col}"
        async with httpx.AsyncClient(headers=headers, timeout=TIMEOUT, follow_redirects=True,
                                     max_redirects=MAX_REDIRECTS,
                                     limits=httpx.Limits(max_connections=self.conc + 4)) as cli:
            if mode == "all":
                xml, _ = await self.fetch(cli, SITEMAP_URL)
                urls = parse_sitemap(xml) if xml else []
                urls = list(dict.fromkeys(urls))[:limit or MAX_ITEMS]
                self.log(f"[SYS] sitemap-tools.xml -> {len(urls)} URL san pham")
                await self._crawl_urls(cli, urls, key)

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
                    html, _ = await self.fetch(cli, f"{search_url(target)}&page={page}")
                    new = [u for u in (parse_search(html) if html else []) if u not in urls]
                    if not new:                       # stuck-page = het ket qua
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
    root.title("Pro'sKit Crawler — proskit.com")
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
                     ("all", "5. Crawl ALL (sitemap ~2.7k SP)")]:
        ttk.Radiobutton(left, text=txt, value=val, variable=mode).pack(anchor="w")

    ttk.Label(left, text="\nMa / Tu khoa:").pack(anchor="w")
    ent = ttk.Entry(left, width=36)
    ent.pack(anchor="w")
    ttk.Label(left, text="Link (moi dong 1 link):").pack(anchor="w")
    box = tk.Text(left, width=36, height=5)
    box.pack(anchor="w")

    xls, sheet, col = tk.StringVar(), tk.StringVar(), tk.StringVar()
    outdir = tk.StringVar(value="images")
    attach_excel_picker(left, xls, sheet, col, outdir)

    conc, delay, limit = tk.IntVar(value=10), tk.DoubleVar(value=0.05), tk.IntVar(value=0)
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
        if m in ("code", "keyword") and not ent.get().strip():
            messagebox.showwarning("Thieu du lieu", "Nhap ma / tu khoa truoc.")
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
        p = ask_export_path(xls.get(), HERE, "proskit.xlsx")
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
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")   # console cp1252 -> vo CJK/VN
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Pro'sKit crawler (proskit.com)")
    ap.add_argument("--mode", choices=["code", "keyword", "link", "excel", "all"])
    ap.add_argument("--code")
    ap.add_argument("--keyword")
    ap.add_argument("--link")
    ap.add_argument("--excel")
    ap.add_argument("--sheet", default="")
    ap.add_argument("--col", default="")
    ap.add_argument("--out", default="proskit.xlsx")
    ap.add_argument("--conc", type=int, default=10)
    ap.add_argument("--delay", type=float, default=0.05)
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
