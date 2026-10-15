"""
toptul.com — TOPTUL hand-tool catalog crawler (httpx + selectolax, NO browser, NO Selenium).

RECON (2026-07-13)
------------------
API      : KHONG co JSON API. Site render HTML server-side => httpx + selectolax la du.
robots   : "Allow: /" (khong cam gi). sitemap.xml = sitemapindex, lastmod 2026-07-01 (MOI).
sitemap  : /sitemapen_site_product_group_<1..25>.xml -> ~100 URL/file (~2.5k trang EN)
           (ho `sitemapenproduct*.xml` la URL kieu cu /2-2179-<c>/product/...idNNN.html -> bo qua)
detail   : /en/product-<pid>/<slug>.html
           - h1.product-name                  -> ten trang (thuong la 1 SERIES)
           - script[application/ld+json]      -> BreadcrumbList (category) + Product (anh absolute)
           - div.des-txt                      -> "ITEM NO: XXXX" + bullet mo ta + cac bang
           - bang SPEC co ROWSPAN/COLSPAN, header 2 tang (vd "ITEM NO." colspan=2 -> Mirror|Satin)
           - 1 TRANG = NHIEU MODEL: moi dong bang co cot "ITEM NO." = 1 ma
             (product-325055 -> BAEA1608..BAEA1634). Trang kit/set -> 1 ma nam trong mo ta.
search   : GET /en/search?key=<ma>   (param la `key`, KHONG phai `keyword`) -> HTML server-side,
           link ket qua /en/product-<pid>/... . Khop ma chinh xac lai o buoc loc dong bang.
gia      : KHONG co gia cong khai (chi "Inquire Now") -> khong crawl gia.

Run
---
  toptul.bat                                   -> GUI
  python toptul.py --mode all [--limit 50]     -> quet sach sitemap
  python toptul.py --mode code --code BAEA1611
  python toptul.py --mode keyword --keyword "torque wrench"
  python toptul.py --mode link --link https://www.toptul.com/en/product-325055/....html
  python toptul.py --mode excel --excel parts.xlsx --sheet Sheet1 --col Part
  python toptul.py --mode images --excel parts.xlsx --col Link --outdir images
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
from urllib.parse import quote, urljoin

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser

HERE = pathlib.Path(__file__).parent
BASE = "https://www.toptul.com"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")

CODE_INDEX = HERE / "_code_index.json"   # ma -> URL (search toptul cham 3.4s/ma, xem _resolve_codes)
INDEX_TTL_DAYS = 30
BUILD_INDEX_FROM = 40     # >= N ma chua biet -> quet sitemap dung index thay vi search tung ma

SITEMAP_GROUPS = 25       # /sitemapen_site_product_group_1..25.xml
CIRCUIT_BREAK = 25        # N loi LIEN TIEP -> nghi bi chan IP -> dung han
MAX_ITEMS = 5000          # tran work-list
MAX_PAGES = 600           # tran phan trang search
MAX_REDIRECTS = 5
TIMEOUT = 30.0
GC_EVERY = 500

PROD_RE = re.compile(r"/en/product-(\d+)/[^\"']*?\.html")
CODE_RE = re.compile(r"^[A-Z]{2,6}[-_]?[0-9][A-Z0-9._/-]*$")   # BAEA1611, KPEA1236, GFE-0501
ITEMNO_TXT_RE = re.compile(r"ITEM\s*NO[.:]?\s*([A-Z0-9][A-Z0-9._/-]{3,})", re.I)


# ================================================================ CORE (khong dung GUI)
def norm_code(s: Any) -> str:
    """Chuan hoa ma de so khop. URL thi GIU NGUYEN hoa/thuong (path phan biet chu)."""
    v = re.sub(r"\s+", "", str(s or ""))
    return v if v.lower().startswith("http") else v.upper()


def search_url(key: str) -> str:
    return f"{BASE}/en/search?key={quote(key)}"


def sitemap_urls() -> list[str]:
    """File goc + group_2..25 (group_1 KHONG ton tai - 404)."""
    return ([f"{BASE}/sitemapen_site_product.xml"] +
            [f"{BASE}/sitemapen_site_product_group_{i}.xml" for i in range(2, SITEMAP_GROUPS + 1)])


def parse_sitemap(xml: str) -> list[str]:
    return [u for u in re.findall(r"<loc>([^<]+)</loc>", xml) if PROD_RE.search(u)]


def parse_search(html: str) -> list[str]:
    """Link san pham trong trang search — giu thu tu, bo trung."""
    seen: set[str] = set()
    out: list[str] = []
    for m in PROD_RE.finditer(html):
        u = urljoin(BASE, m.group(0))
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def build_grid(table) -> list[list[str]]:
    """Bang HTML -> luoi text chu nhat; giai rowspan/colspan (o gop -> lap gia tri ra cac o con)."""
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
            txt = re.sub(r"\s+", " ", cell.text(strip=True) or "")
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


def split_header(grid: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    """Tach header (1..n tang, gop ten) khoi dong du lieu; data = tu dong DAU co ma san pham."""
    first_data = next((i for i, row in enumerate(grid) if any(CODE_RE.match(c) for c in row)), None)
    if not first_data:                       # None (khong co ma) hoac 0 (khong co header)
        return [], grid
    head_rows = grid[:first_data]
    header = []
    for c in range(len(grid[0])):
        parts, seen = [], set()
        for hr in head_rows:
            v = hr[c] if c < len(hr) else ""
            if v and v not in seen:
                seen.add(v)
                parts.append(v)
        header.append(" ".join(parts).strip())
    return header, grid[first_data:]


def _record(code: str, page: str, cats: list[str], spec: dict, images: list[str], url: str) -> dict:
    return {"Input": "", "Item No": code, "Page": page, "Status": "OK",
            "_cats": cats, "_spec": spec, "Image": " | ".join(images[:5]), "URL": url}


def parse_spec_tables(root, page: str, cats: list[str], images: list[str], url: str) -> list[dict]:
    """Moi dong bang co cot 'ITEM NO.' -> 1 record rieng (1 trang = nhieu model)."""
    out: list[dict] = []
    seen: set[str] = set()
    for table in root.css("table"):
        grid = build_grid(table)
        if len(grid) < 2:
            continue
        header, rows = split_header(grid)
        item_cols = [i for i, h in enumerate(header) if "ITEM" in h.upper()]
        if not item_cols:
            continue
        for row in rows:
            for ic in item_cols:
                code = row[ic] if ic < len(row) else ""
                if not CODE_RE.match(code) or code in seen:
                    continue
                seen.add(code)
                spec = {header[i]: row[i] for i in range(len(header))
                        if i not in item_cols and header[i] and i < len(row) and row[i]}
                out.append(_record(code, page, cats, spec, images, url))
    return out


def parse_page(html: str, url: str) -> list[dict]:
    """HTML trang detail -> list record (1 record / ITEM NO)."""
    t = LexborHTMLParser(html)
    n = t.css_first("h1.product-name")
    page = re.sub(r"\s+", " ", n.text(strip=True)) if n else ""

    cats: list[str] = []
    images: list[str] = []
    for sc in t.css('script[type="application/ld+json"]'):
        try:
            data = json.loads(sc.text())
        except Exception:
            continue
        for blk in (data if isinstance(data, list) else [data]):
            if not isinstance(blk, dict):
                continue
            if blk.get("@type") == "BreadcrumbList":       # [home, PRODUCT, cat.., ten SP]
                cats = [e.get("name", "") for e in blk.get("itemListElement", [])][2:-1]
            elif blk.get("@type") == "Product":
                for im in blk.get("image") or []:
                    u = im.get("contentUrl") if isinstance(im, dict) else im
                    if u and u not in images:
                        images.append(u)
    if not images:
        for s in t.css("picture source"):
            ss = (s.attributes.get("srcset") or "").split()
            if ss:
                u = urljoin(BASE + "/", ss[0])
                if u not in images:
                    images.append(u)

    if t.body is None:
        return []
    # Bang spec nam NGOAI div.des-txt (tab thong so) -> quet bang tren ca trang
    recs = parse_spec_tables(t.body, page, cats, images, url)
    if recs:
        return recs

    # Trang kit/set: khong co bang ma -> ma nam trong text mo ta, spec = bang key/value 2 cot
    des = t.css_first("div.des-txt") or t.body
    codes = [norm_code(c) for c in ITEMNO_TXT_RE.findall(des.text())]
    codes = [c for c in dict.fromkeys(codes) if CODE_RE.match(c)]
    spec: dict[str, str] = {}
    for table in des.css("table"):
        for row in build_grid(table):
            if len(row) == 2 and row[0] and row[1] and not CODE_RE.match(row[0]):
                spec.setdefault(row[0], row[1])
    if not codes:
        return [_record("", page, cats, spec, images, url)] if page else []
    return [_record(c, page, cats, spec, images, url) for c in codes]


def export_excel(results: list[dict], path: pathlib.Path) -> pathlib.Path:
    """Spec gop 1 COT 'Thong so' (moi dong 'Ten: Gia tri') - khong trai union key thanh
    hang tram cot rong. Xem AGENTS.md 4.6."""
    from openpyxl.utils import get_column_letter

    max_cat = max((len(r.get("_cats") or []) for r in results), default=0)
    rows = []
    for r in results:
        row = {"Input": r.get("Input", ""), "Item No": r.get("Item No", ""),
               "Page": r.get("Page", ""), "Status": r.get("Status", "")}
        cats = r.get("_cats") or []
        for i in range(max_cat):
            row[f"Category {i + 1}"] = cats[i] if i < len(cats) else ""
        row["Thong so"] = "\n".join(f"{k}: {v}" for k, v in (r.get("_spec") or {}).items())
        row["Image"] = r.get("Image", "")
        row["URL"] = r.get("URL", "")
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

    def __init__(self, conc: int = 8, delay: float = 0.05, log=print,
                 stop_evt: threading.Event | None = None,
                 pause_evt: threading.Event | None = None) -> None:
        self.conc, self.delay, self.log = max(1, conc), max(0.0, delay), log
        self.stop_evt = stop_evt or threading.Event()
        self.pause_evt = pause_evt or threading.Event()
        self.fails = 0
        self.results: list[dict] = []
        self.done = 0
        self.total = 0
        self.key = ""            # dat trong run(); cleanup() dung lai -> khong lech ten checkpoint
        self._ckpt_loaded = False
        self._index: dict[str, str] | None = None   # code-index (ma -> URL), lazy load
        self._index_built_at: float | None = None   # tuoi cua DU LIEU index (khong phai mtime file)

    # ---- resume: JSONL sidecar (append tung item; chay lai = doc lai, bo qua trang da xong)
    def ckpt(self, key: str) -> pathlib.Path:
        return HERE / f"_raw_{re.sub(r'[^A-Za-z0-9]+', '_', key)[:40]}.jsonl"

    def _load_ckpt(self, key: str) -> set[str]:
        p = self.ckpt(key)
        seen: set[str] = set()
        if self._ckpt_loaded:                 # goi 2 lan (link + code) -> KHONG nap lai, tranh nhan doi
            return {r.get("URL", "") for r in self.results}
        self._ckpt_loaded = True
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
            self.log(f"[SYS] Resume: {len(self.results)} record / {len(seen)} trang da xong -> bo qua")
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
                if r.status_code in (403, 429, 503):          # WAF / rate-limit -> lui dan
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
                          codes: set[str] | None = None,
                          inputs: dict[str, str] | None = None) -> None:
        """Crawl cac URL detail. codes != None -> chi giu dong bang khop ma (trang = nhieu model)."""
        seen = self._load_ckpt(key)
        urls = [u for u in dict.fromkeys(urls) if u not in seen][:MAX_ITEMS]
        self.total = len(urls) + len(seen)
        self.done = len(seen)
        sem = asyncio.Semaphore(self.conc)

        async def one(u: str) -> None:
            async with sem:
                if not await self._gate():           # re-check SAU khi gianh slot
                    return
                html = await self.fetch(cli, u)
                if self.delay:
                    await asyncio.sleep(self.delay)
                recs = parse_page(html, u) if html else []
                if codes is not None:
                    recs = [r for r in recs if norm_code(r["Item No"]) in codes]
                if inputs:
                    for r in recs:
                        r["Input"] = inputs.get(norm_code(r["Item No"]), inputs.get(u, ""))
                if recs:
                    self.results.extend(recs)
                    self._save(key, recs)
                    self.index_records(recs)     # lam giau code-index cho lan sau
                self.done += 1
                self.log(f"[{'OK ' if recs else 'MISS'}] {self.done}/{self.total} "
                         f"+{len(recs)} | {u[-58:]}")
                if self.done % GC_EVERY == 0:
                    gc.collect()

        try:
            await asyncio.gather(*(one(u) for u in urls))
        except (KeyboardInterrupt, Exception) as e:   # crash-proof: record da nam trong JSONL
            self.log(f"[FATAL] {type(e).__name__}: {e}")

    # ---- CODE INDEX: search cua toptul cham ~3.4s/ma, nhung 1 TRANG chua toi 26 ma
    # (ca series BAEA1608..1634). Nen: tai 1 trang -> ghi nho MOI ma trong do -> cac ma cung
    # series khoi phai search. Cache ra dia, dung lai cho lan chay sau.
    def _load_index(self) -> dict[str, str]:
        if self._index is not None:
            return self._index
        self._index = {}
        if CODE_INDEX.exists():
            try:
                data = json.loads(CODE_INDEX.read_text("utf-8"))
                # built_at = tuoi THUC cua du lieu, khong phai mtime file (mtime bi lam moi
                # moi lan save_index() ghi de -> TTL se khong bao gio het han). File dinh dang
                # cu (bare mapping, khong co built_at) duoc coi la cu -> xay lai, khong crash.
                built_at = data.get("built_at") if isinstance(data, dict) else None
                if built_at is not None and (time.time() - built_at) / 86400 < INDEX_TTL_DAYS:
                    self._index = data.get("codes", {})
                    self._index_built_at = built_at   # nho lai de save_index KHONG vo tinh lam tre no
                    self.log(f"[SYS] Code-index: {len(self._index)} ma da biet (khoi search)")
                # qua han (hoac khong co built_at) -> coi nhu rong, KHONG giu built_at cu: cac
                # ma hoc duoc trong lan chay nay se duoc luu voi built_at MOI (mo cua so TTL moi)
            except Exception:
                self._index = {}
        return self._index

    def index_records(self, recs: list[dict]) -> None:
        """Moi record crawl duoc -> nap vao index (ma -> URL)."""
        idx = self._load_index()
        for r in recs:
            c = norm_code(r.get("Item No"))
            if c and r.get("URL"):
                idx[c] = r["URL"]

    def save_index(self) -> None:
        if self._index:
            if self._index_built_at is None:   # index moi/xay lai -> danh dau tuoi tu bay gio
                self._index_built_at = time.time()
            data = {"built_at": self._index_built_at, "codes": self._index}
            CODE_INDEX.write_text(json.dumps(data, ensure_ascii=False), "utf-8")

    async def build_index(self, cli: httpx.AsyncClient) -> dict[str, str]:
        """Quet SACH sitemap -> index ma->URL cho CA catalog. Trang detail rat nhanh (0.04s/trang
        o conc 8) nen ~2.5k trang chi ton ~100s, doi lai TRA MA TUC THI mai ve sau (TTL 30 ngay).
        Rẻ hơn nhiều so voi search 3.4s/ma khi list dai."""
        sem = asyncio.Semaphore(self.conc)

        async def sm(u: str) -> list[str]:
            async with sem:
                xml = await self.fetch(cli, u)
                return parse_sitemap(xml) if xml else []

        urls = list(dict.fromkeys(
            [u for part in await asyncio.gather(*(sm(s) for s in sitemap_urls())) for u in part]))
        self.log(f"[SYS] Dung code-index: quet {len(urls)} trang (~{len(urls) * 0.05:.0f}s)...")
        idx = self._load_index()
        self._index_built_at = time.time()   # quet lai TOAN BO sitemap -> xay lai tu dau, built_at moi
        done = 0

        async def one(u: str) -> None:
            nonlocal done
            async with sem:
                if not await self._gate():
                    return
                html = await self.fetch(cli, u)
                if html:
                    self.index_records(parse_page(html, u))
                done += 1
                if done % 200 == 0:
                    self.log(f"[SYS] index: {done}/{len(urls)} trang | {len(idx)} ma")

        await asyncio.gather(*(one(u) for u in urls))
        self.save_index()
        self.log(f"[SYS] Code-index xong: {len(idx)} ma -> tu gio tra ma KHONG can search")
        return idx

    async def _resolve_codes(self, cli: httpx.AsyncClient, codes: list[str]) -> tuple[list[str], dict]:
        """Ma -> URL. Uu tien code-index (0 request); con lai moi search (3.4s/ma, song song).
        Sau MOI lan search ra trang, trang do duoc parse -> mo khoa luon cac ma cung trang."""
        idx = self._load_index()
        urls: list[str] = []
        inputs: dict[str, str] = {}
        hit_idx = 0
        todo: list[str] = []
        for c in codes:
            if c in idx:
                urls.append(idx[c])
                inputs[c] = c
                hit_idx += 1
            else:
                todo.append(c)
        if hit_idx:
            self.log(f"[SYS] {hit_idx}/{len(codes)} ma lay tu code-index -> tiet kiem "
                     f"{hit_idx * 3.4:.0f}s search")
        if not todo:
            return urls, inputs

        # List dai: search ton ~3.4s/ma (server toptul cham), trong khi quet sach sitemap dung
        # index chi ~100s va dung duoc mai -> tu dong doi chien thuat.
        if len(todo) >= BUILD_INDEX_FROM:
            self.log(f"[SYS] {len(todo)} ma chua biet — search se ton ~{len(todo) * 1.5:.0f}s. "
                     f"Quet sitemap dung index (~100s, dung lai mai) thay the.")
            idx = await self.build_index(cli)
            still: list[str] = []
            for c in todo:
                if c in idx:
                    urls.append(idx[c])
                    inputs[c] = c
                else:
                    still.append(c)
            self.log(f"[SYS] Index phu {len(todo) - len(still)}/{len(todo)} ma")
            todo = still
            if not todo:
                return urls, inputs

        lock = asyncio.Lock()

        async def one(c: str) -> None:
            if not await self._gate():
                return
            async with lock:                     # ma nay co the vua duoc trang khac mo khoa
                if c in idx:
                    urls.append(idx[c])
                    inputs[c] = c
                    return
            html = await self.fetch(cli, search_url(c))
            await asyncio.sleep(self.delay)
            hits = parse_search(html) if html else []
            if not hits:
                self.log(f"[MISS] search khong ra ma: {c}")
                return
            urls.append(hits[0])
            inputs[c] = c
            page = await self.fetch(cli, hits[0])   # trang chua CA SERIES -> nap het vao index
            if page:
                async with lock:
                    self.index_records(parse_page(page, hits[0]))

        async def wave(cs: list[str]) -> None:
            sem = asyncio.Semaphore(self.conc)

            async def guarded(c: str) -> None:
                async with sem:
                    await one(c)

            await asyncio.gather(*(guarded(c) for c in cs))

        # Search cham (~3.4s) nhung SONG SONG tot => toi da 2 VONG CHO, khong chia dot tuan tu:
        #  vong 1: search MOT MAU RAI DEU (moi trang toptul chua toi 26 ma -> 1 mau mo khoa ca cum)
        #  vong 2: bung song song nhung ma con lai chua co trong index
        # List dang series -> vong 2 gan nhu rong. List roi rac -> chi ton them 1 vong.
        if len(todo) > self.conc:
            step = max(1, len(todo) // max(2, self.conc // 3))
            sample = todo[::step][: self.conc]
            await wave(sample)
            rest = [c for c in todo if c not in idx and c not in inputs]
            unlocked = sum(1 for c in todo if c not in sample and c in idx)
            self.log(f"[SYS] Mau {len(sample)} ma -> code-index mo khoa them {unlocked} ma "
                     f"| con {len(rest)} ma phai search")
        else:
            rest = todo
        if rest:
            await wave(rest)
        self.save_index()
        return urls, inputs

    async def run(self, mode: str, target: str = "", excel: str = "", sheet: str = "",
                  col: str = "", limit: int = 0, outdir: str = "images") -> list[dict]:
        headers = {"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"}
        key = self.key = f"{mode}_{target or excel}_{col}"
        async with httpx.AsyncClient(headers=headers, timeout=TIMEOUT, follow_redirects=True,
                                     max_redirects=MAX_REDIRECTS,
                                     limits=httpx.Limits(max_connections=self.conc + 4)) as cli:
            if mode == "all":
                sems = asyncio.Semaphore(self.conc)   # sitemap cham (~7-11s/file) -> keo song song

                async def sm_one(sm: str) -> list[str]:
                    async with sems:
                        if not await self._gate():
                            return []
                        xml = await self.fetch(cli, sm)
                        got = parse_sitemap(xml) if xml else []
                        self.log(f"[SYS] {sm.rsplit('/', 1)[-1]} -> {len(got)} URL")
                        return got

                urls = [u for part in await asyncio.gather(*(sm_one(s) for s in sitemap_urls()))
                        for u in part]
                urls = list(dict.fromkeys(urls))[:limit or MAX_ITEMS]
                self.log(f"[SYS] Crawl ALL: {len(urls)} trang san pham")
                await self._crawl_urls(cli, urls, key)

            elif mode in ("code", "excel"):
                raw = [norm_code(target)] if mode == "code" else read_excel_col(excel, sheet, col)
                raw = [c for c in dict.fromkeys(raw) if c][:MAX_ITEMS]
                links = [c for c in raw if c.lower().startswith("http")]
                codes = [c for c in raw if c not in links]
                # LINK crawl RIENG: 1 trang = nhieu model, neu tron chung voi codes thi bo loc
                # `codes` se xoa sach cac dong cua link (Item No cua no dau co trong danh sach ma)
                if links:
                    await self._crawl_urls(cli, links, key, inputs={u: u for u in links})
                if codes:
                    urls, inputs = await self._resolve_codes(cli, codes)
                    await self._crawl_urls(cli, urls, key, codes=set(codes), inputs=inputs or None)
                have = {norm_code(r["Item No"]) for r in self.results}
                for c in codes:                       # ma miss van co dong Not Found
                    if c not in have:
                        self.results.append({"Input": c, "Item No": c, "Page": "",
                                             "Status": "Not Found", "_cats": [], "_spec": {},
                                             "Image": "", "URL": ""})

            elif mode == "keyword":
                urls, page = [], 1
                while page <= MAX_PAGES:
                    if not await self._gate():
                        break
                    html = await self.fetch(cli, f"{search_url(target)}&page={page}")
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

            elif mode == "images":
                await self._download_images(cli, excel, sheet, col, outdir)

            else:
                self.log(f"[ERR] mode khong hop le: {mode}")
        self.save_index()          # index tich luy qua moi lan chay -> lan sau khoi search
        return self.results

    async def _download_images(self, cli: httpx.AsyncClient, excel: str, sheet: str,
                               col: str, outdir: str) -> None:
        """Excel (cot = URL trang SP hoac ma) -> anh tu ld+json -> luu outdir/<ItemNo>_N.jpg."""
        out = pathlib.Path(outdir)
        out.mkdir(parents=True, exist_ok=True)
        keys = read_excel_col(excel, sheet, col)[:MAX_ITEMS]
        self.total = len(keys)
        sem = asyncio.Semaphore(self.conc)

        async def one(k: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                url = k if k.lower().startswith("http") else None
                if not url:
                    html = await self.fetch(cli, search_url(k))
                    hits = parse_search(html) if html else []
                    url = hits[0] if hits else None
                recs = []
                if url:
                    html = await self.fetch(cli, url)
                    recs = parse_page(html, url) if html else []
                self.done += 1
                if not recs:
                    self.log(f"[MISS] {self.done}/{self.total} khong ra anh: {k}")
                    return
                imgs = [u for u in (recs[0]["Image"] or "").split(" | ") if u]
                # dat ten theo MA NHAP (1 trang co nhieu ma -> tranh lay nham ma dong dau)
                name = (norm_code(k) if not k.lower().startswith("http")
                        else norm_code(recs[0]["Item No"]) or re.sub(r"\W+", "_", k)[:40])
                ok = 0
                for i, iu in enumerate(imgs, 1):
                    try:
                        r = await cli.get(iu)
                        r.raise_for_status()
                        (out / f"{name}_{i}{pathlib.Path(iu).suffix or '.jpg'}").write_bytes(r.content)
                        ok += 1
                    except Exception as e:
                        self.log(f"[ERR] tai anh {iu[-50:]}: {type(e).__name__}")
                self.log(f"[OK ] {self.done}/{self.total} {name} -> {ok} anh")

        await asyncio.gather(*(one(k) for k in keys))


def read_excel_col(path: str, sheet: str, col: str) -> list[str]:
    """Loi -> ValueError (KHONG dung SystemExit: no la BaseException, worker GUI bat
    `except Exception` se KHONG bat duoc -> thread chet cam, GUI treo running=True)."""
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

    # ha tang GUI chung o root (di nguoc len cay thu muc -> song sot khi doi nhom folder)
    root_dir = next(p for p in (HERE, *HERE.parents) if (p / "gui_common.py").exists())
    sys.path.insert(0, str(root_dir))
    from gui_common import ask_export_path, attach_excel_picker

    root = tk.Tk()
    root.title("TOPTUL Crawler — toptul.com")
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
                     ("all", "5. Crawl ALL (sitemap ~2.5k trang)"), ("images", "6. Tai ANH tu Excel")]:
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

    conc, delay, limit = tk.IntVar(value=8), tk.DoubleVar(value=0.05), tk.IntVar(value=0)
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
            asyncio.run(eng.run(m, target, xls.get(), sheet.get(), col.get(),
                                limit.get(), outdir.get()))
            eng.cleanup()
        except BaseException as e:      # BaseException: SystemExit/KeyboardInterrupt cung phai bat,
            q.put(f"[FATAL] {type(e).__name__}: {e}")   # khong thi thread chet cam, GUI treo
        finally:                        # finally: LUON nha co, khong khoa START vinh vien
            state["results"] = eng.results
            state["running"] = False
            q.put(f"[SYS] XONG — {len(eng.results)} record. Bam EXPORT de xuat Excel.")

    def start() -> None:
        if state["running"]:
            messagebox.showinfo("Dang chay", "Dang crawl. Bam Stop truoc.")
            return
        m = mode.get()                       # chan dau vao thieu -> BAO RO, khong im lang
        if m in ("excel", "images") and not (xls.get() and col.get()):
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
        p = ask_export_path(xls.get(), HERE, "toptul.xlsx")   # mac dinh canh file Excel dau vao
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
    ap = argparse.ArgumentParser(description="TOPTUL crawler (toptul.com)")
    ap.add_argument("--mode", choices=["code", "keyword", "link", "excel", "all", "images"])
    ap.add_argument("--code")
    ap.add_argument("--keyword")
    ap.add_argument("--link")
    ap.add_argument("--excel")
    ap.add_argument("--sheet", default="")
    ap.add_argument("--col", default="")
    ap.add_argument("--outdir", default="images")
    ap.add_argument("--out", default="toptul.xlsx")
    ap.add_argument("--conc", type=int, default=8)   # do duoc: conc>=16 bi server drop request
    ap.add_argument("--delay", type=float, default=0.05)
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--build-index", action="store_true",
                    help="quet sach sitemap dung code-index (~100s) -> tra ma tuc thi ve sau")
    a = ap.parse_args()

    if a.build_index:
        async def _bi() -> None:
            eng = Engine(a.conc, a.delay)
            async with httpx.AsyncClient(headers={"User-Agent": UA}, timeout=TIMEOUT,
                                         follow_redirects=True) as cli:
                await eng.build_index(cli)
        asyncio.run(_bi())
        return

    if not a.mode:
        run_gui()
        return

    target = a.code or a.keyword or a.link or ""
    eng = Engine(a.conc, a.delay)
    t0 = time.time()
    try:
        asyncio.run(eng.run(a.mode, target, a.excel or "", a.sheet, a.col, a.limit, a.outdir))
    finally:
        if eng.results and a.mode != "images":
            p = export_excel(eng.results, pathlib.Path(a.out))
            print(f"[SYS] {len(eng.results)} record -> {p} ({time.time() - t0:.1f}s)")
        eng.cleanup()


if __name__ == "__main__":
    main()
