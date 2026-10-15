"""
Trung Tam Dung Cu crawler (trungtamdungcu.com)
===============================================
Recon (2026-08-11):
  * Nen tang: Haravan (banner /robots.txt: "we use Haravan as our ecommerce
    platform"). CDN anh: cdn.hstatic.net (Haravan storage).
  * KHONG co JSON search API dung duoc: /search?q=..&view=json van tra ve
    FULL HTML page (~670KB), khong phai JSON -- view=json bi bo qua boi theme.
    /products.json cung tra rong (route bi tat). => KHONG dung duoc pattern
    Shopify/Haravan search-API chuan.
  * CO detail JSON: GET /products/<slug>.js -> JSON day du (variants, price,
    images, vendor, type, tags, description HTML). Day la nguon du lieu chinh.
  * Sitemap DAY DU va MOI (cap nhat real-time theo robots.txt comment):
    /sitemap.xml -> 78 file sitemap_products_N.xml (~1000 URL/file, tong
    ~78,000 san pham) + sitemap_collections_1.xml (502 category/brand hub).
    => Dung sitemap de dung CODE-INDEX (mo hinh toptul 5.1): tai 78 file
    sitemap 1 lan, cache ra dia (TTL 7 ngay), tach "duoi" slug (vd
    m21-id13100cp / dcd709n-kr) de khop mã -> tra cuu O(1), KHOI search.
  * URL slug KHONG PHAI la ma thuan (vd "13mm-khoan-100n-m-gen-4-turbo-xam-
    dekton-m21-id13100cp") -- ma san pham nam o CUOI slug, noi voi ten mo ta
    bang dau '-'. Khop bang cach noi N token cuoi (bo dau '-') va so khop
    CHINH XAC (khong substring mu) voi ma da chuan hoa.
  * Gia trong .js la GIA*100 (vd hien thi 1,120,460d nhung .js tra
    "price": 112046000) -- da xac minh doi chieu voi trang HTML (og:price).
    Chia 100 de ra VND that.
  * Spec ky thuat nam trong "description" (HTML) dang
    <ul><li><strong>Ten:</strong> Gia tri</li></ul> -- khong phai <table>.
    Parse bullet -> gop 1 cot "Thong so" (§4.6).
  * 404 tren .js -> HTTP 404 sach (khong redirect ve trang loi), de detect.

Pipeline (GUI 4 lua chon, xem AGENTS.md §4.2):
  1) Theo MA        -> tra CODE-INDEX (tu sitemap) -> fetch .js
  2) Theo LINK/URL   -> fetch .js truc tiep tu URL/slug dan vao
  3) Theo cot Excel  -> chon file -> ma hoac link
  4) Crawl ALL       -> toan bo CODE-INDEX (co --limit / MAX_ITEMS an toan)

Resume: JSONL sidecar (_raw_<key>.jsonl) - xoa khi xong sach, giu khi crash.
"""
from __future__ import annotations

import argparse
import asyncio
import hashlib
import html
import json
import os
import queue
import re
import sys
import threading
import time
from pathlib import Path
from tkinter import *
from tkinter import filedialog, messagebox, ttk

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

HERE = Path(__file__).parent
ROOT = next(p for p in (HERE, *HERE.parents) if (p / "gui_common.py").exists())
sys.path.insert(0, str(ROOT))
from gui_common import ask_export_path, attach_excel_picker  # noqa: E402

BASE = "https://trungtamdungcu.com"
HOST = "trungtamdungcu.com"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

SITEMAP_INDEX = f"{BASE}/sitemap.xml"
INDEX_CACHE = HERE / ".sitemap_index_cache.json"
INDEX_TTL_SEC = 7 * 24 * 3600

# ---- safety consts (AGENTS.md §4.5) ----
MAX_REDIRECTS = 5
CIRCUIT_BREAK = 25
MAX_ITEMS = 5000
MAX_PAGES = 600          # tran so file sitemap_products_N.xml doc (thuc te 78)


# ============================================================================
# Core helpers (pure, no GUI) -> testable headless
# ============================================================================
def norm_code(s: str) -> str:
    """Chuan hoa ma de so khop: bo trang, upper. KHONG dung cho URL (§10 bay 1)."""
    return re.sub(r"\s+", "", str(s)).upper()


def _alnum(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", s).upper()


def slug_from_url(url: str) -> str:
    return url.rstrip("/").rsplit("/", 1)[-1]


def detail_url(slug_or_url: str) -> str:
    """Chap nhan slug hoac URL day du -> URL trang san pham (khong .js)."""
    e = slug_or_url.strip()
    if e.startswith("http"):
        return e.split("?")[0].rstrip("/")
    if e.startswith("/"):
        return BASE + e
    return f"{BASE}/products/{e}"


def js_url(page_url: str) -> str:
    return page_url if page_url.endswith(".js") else page_url + ".js"


_TAG = re.compile(r"<[^>]+>")


def extract_spec_from_description(desc_html: str | None) -> str:
    """description HTML -> cot 'Thong so' gop 1 cot (§4.6).

    Format thuc te: <ul><li><strong>Ten:</strong> Gia tri</li></ul>.
    Dong khong co dau ':' (bullet mo ta) noi luon vao cung cot.
    """
    if not desc_html:
        return ""
    tree = LexborHTMLParser(desc_html)
    lines = []
    for li in tree.css("li"):
        txt = re.sub(r"\s+", " ", li.text(separator=" ", strip=True)).strip()
        if txt:
            lines.append(txt)
    if not lines:
        # fallback: khong co <li> -> lay text tho, bo tag
        txt = html.unescape(_TAG.sub(" ", desc_html))
        txt = re.sub(r"[ \t]+", " ", txt).strip()
        lines = [t.strip() for t in txt.split("\n") if t.strip()]
    return "\n".join(lines)


def parse_detail_json(data: dict, entry: str, page_url: str) -> dict:
    """Haravan /products/<slug>.js JSON -> 1 record phang (§4.6: scalar moi o)."""
    variants = data.get("variants") or [{}]
    v0 = variants[0]
    prices = [round((vv.get("price") or 0) / 100) for vv in variants if vv.get("price")]
    images = data.get("images") or []
    return {
        "Input": entry,
        "Page": data.get("title") or "",
        "Item No": v0.get("sku") or "",
        "Status": "Success",
        "Thuong hieu": data.get("vendor") or "",
        "Loai": data.get("type") or "",
        "Gia": min(prices) if prices else None,
        "Gia goc": round((v0.get("compare_at_price") or 0) / 100) or None,
        "Con hang": "Con hang" if data.get("available") else "Het hang",
        "So bien the": len(variants),
        "Tags": ", ".join(data.get("tags") or []),
        "Thong so": extract_spec_from_description(data.get("description")),
        "Image": images[0] if images else "",
        "So anh": len(images),
        "URL": page_url,
    }


def not_found(entry: str, url: str = "", reason: str = "Not Found") -> dict:
    return {"Input": entry, "Page": "", "Item No": "", "Status": reason, "URL": url}


# ============================================================================
# Sitemap CODE-INDEX (build once, cache to disk with TTL) -- §5.1 ky thuat 3
# ============================================================================
async def _get_text(client: httpx.AsyncClient, url: str) -> str:
    r = await client.get(url)
    r.raise_for_status()
    return r.text


_LOC_RE = re.compile(r"<loc>(.*?)</loc>")
_TITLE_RE = re.compile(r'<image:title>(.*?)</image:title>')
_URL_BLOCK_RE = re.compile(r"<url>(.*?)</url>", re.S)


def _parse_products_sitemap(xml_text: str) -> list[dict]:
    out = []
    for block in _URL_BLOCK_RE.findall(xml_text):
        m = _LOC_RE.search(block)
        if not m:
            continue
        url = m.group(1).strip()
        tm = _TITLE_RE.search(block)
        title = html.unescape(tm.group(1)).strip() if tm else ""
        out.append({"url": url, "slug": slug_from_url(url), "title": title})
    return out


async def build_index(log=None, conc: int = 8) -> list[dict]:
    log = log or (lambda s: None)
    limits = httpx.Limits(max_connections=conc, max_keepalive_connections=conc)
    async with httpx.AsyncClient(headers={"User-Agent": UA}, timeout=30.0,
                                  limits=limits, follow_redirects=True,
                                  max_redirects=MAX_REDIRECTS) as client:
        idx_xml = await _get_text(client, SITEMAP_INDEX)
        sitemaps = [u for u in _LOC_RE.findall(idx_xml) if "sitemap_products_" in u]
        sitemaps = sitemaps[:MAX_PAGES]
        log(f"[SYS] Tai {len(sitemaps)} file sitemap san pham...\n")
        sem = asyncio.Semaphore(conc)
        items: list[dict] = []
        done = 0

        async def one(u):
            nonlocal done
            async with sem:
                try:
                    xml = await _get_text(client, u)
                except Exception as e:
                    log(f"[WARN] loi tai {u}: {e}\n")
                    return
            items.extend(_parse_products_sitemap(xml))
            done += 1
            if done % 10 == 0 or done == len(sitemaps):
                log(f"[SYS] sitemap {done}/{len(sitemaps)} -> {len(items)} san pham\n")

        await asyncio.gather(*(one(u) for u in sitemaps))
    log(f"[SYS] Index xong: {len(items)} san pham.\n")
    return items


def _build_code_map(items: list[dict]) -> dict[str, list[dict]]:
    """Ma da chuan hoa (alnum, upper) -> list item khop CHINH XAC.

    Khop bang cach noi 1..3 token cuoi cua slug (tach boi '-'), bo ky tu
    khong phai chu/so, so sanh BANG (khong substring mu -- §10 bay 5)."""
    code_map: dict[str, list[dict]] = {}
    for it in items:
        toks = it["slug"].split("-")
        seen_keys = set()
        for k in (1, 2, 3):
            if k > len(toks):
                break
            key = _alnum("".join(toks[-k:]))
            if len(key) < 3 or key in seen_keys:
                continue
            seen_keys.add(key)
            code_map.setdefault(key, []).append(it)
    return code_map


def _load_cached_index(log, force: bool) -> list[dict] | None:
    if force or not INDEX_CACHE.exists():
        return None
    age = time.time() - INDEX_CACHE.stat().st_mtime
    if age >= INDEX_TTL_SEC:
        return None
    try:
        items = json.loads(INDEX_CACHE.read_text(encoding="utf-8"))
        log(f"[SYS] Dung index cache ({len(items)} sp, {age/3600:.1f}h tuoi).\n")
        return items
    except Exception:
        return None


async def load_or_build_index_async(log=None, force: bool = False) -> tuple[list[dict], dict]:
    """Dung trong coroutine (GUI da o trong event loop) -- KHONG goi asyncio.run()."""
    log = log or (lambda s: None)
    items = _load_cached_index(log, force)
    if items is None:
        items = await build_index(log)
        if items:
            INDEX_CACHE.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")
    return items, _build_code_map(items)


def load_or_build_index_sync(log=None, force: bool = False) -> tuple[list[dict], dict]:
    """Dung tu CLI / thread rieng (chua co event loop dang chay)."""
    log = log or (lambda s: None)
    items = _load_cached_index(log, force)
    if items is None:
        items = asyncio.run(build_index(log))
        if items:
            INDEX_CACHE.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")
    return items, _build_code_map(items)


def code_lookup(code_map: dict[str, list[dict]], code: str) -> list[dict]:
    key = _alnum(code)
    return code_map.get(key, [])


# ============================================================================
# Async crawl engine (usable headless; GUI wraps it)
# ============================================================================
class TTDCEngine:
    def __init__(self, log=None, prog=None, stop_evt=None, pause_evt=None):
        self.log = log or (lambda s: None)
        self.prog = prog or (lambda d, t, m: None)
        self.stop_evt = stop_evt or threading.Event()
        self.pause_evt = pause_evt or threading.Event()
        self.results: list[dict] = []

    def _client(self, conc):
        limits = httpx.Limits(max_connections=conc, max_keepalive_connections=conc)
        transport = httpx.AsyncHTTPTransport(retries=2)
        return httpx.AsyncClient(headers={"User-Agent": UA}, timeout=20.0,
                                  limits=limits, transport=transport,
                                  follow_redirects=True, max_redirects=MAX_REDIRECTS)

    async def _fetch_detail(self, client, entry: str, page_url: str, tries: int = 4) -> dict:
        # BUG 3: retry loi tam thoi voi backoff leo thang; 429/503 = rate-limit, khong
        # phai loi vinh vien (AGENTS.md §6)
        url = js_url(page_url)
        for i in range(tries):
            while self.pause_evt.is_set() and not self.stop_evt.is_set():
                await asyncio.sleep(0.2)
            if self.stop_evt.is_set():
                return not_found(entry, page_url, "Error: Stopped")
            try:
                r = await client.get(url)
                if r.status_code == 404:
                    return not_found(entry, page_url)
                if r.status_code in (429, 503):
                    wait = min(2 ** i * 5, 60)
                    self.log(f"[WARN] {r.status_code} -> ngu {wait}s | {entry}\n")
                    await asyncio.sleep(wait)
                    continue
                r.raise_for_status()
                data = r.json()
                return parse_detail_json(data, entry, page_url)
            except httpx.TooManyRedirects:
                return not_found(entry, page_url, "Redirect loop")
            except Exception as e:
                if i < tries - 1:
                    await asyncio.sleep(min(2 ** i, 20))
                    continue
                return not_found(entry, page_url, f"Error: {type(e).__name__}")
        return not_found(entry, page_url, "Error: RateLimited")

    async def _run_targets(self, client, targets: list[tuple], conc, delay, ckpt_key: str | None):
        """targets: list of (entry, page_url). Resume via JSONL sidecar."""
        seen_input = set()
        jsonl_path = HERE / f"_raw_{ckpt_key}.jsonl" if ckpt_key else None
        if jsonl_path and jsonl_path.exists():
            with jsonl_path.open(encoding="utf-8") as f:
                for line in f:
                    try:
                        rec = json.loads(line)
                    except Exception:
                        continue
                    status = rec.get("Status", "")
                    if status == "Not Found" or status.startswith("Error:"):
                        # BUG 2: mieng/loi tam thoi khong tinh la "da xong" -> phai thu lai
                        continue
                    self.results.append(rec)
                    seen_input.add(rec.get("Input"))
            if seen_input:
                self.log(f"[SYS] Resume: {len(seen_input)} muc da xong tu lan truoc.\n")

        work = [(e, u) for e, u in targets if e not in seen_input]
        total = len(targets)
        if not work:
            self.prog(total, total, sum(1 for r in self.results if r.get("Status") == "Success"))
            return

        if len(work) > MAX_ITEMS:
            self.log(f"[WARN] {len(work)} muc vuot {MAX_ITEMS}, cat bot.\n")
            work = work[:MAX_ITEMS]

        self.log(f"[SYS] Crawl {len(work)} muc (tong {total}) | concurrency={conc}\n")
        sem = asyncio.Semaphore(conc)
        done = len(seen_input)
        models = sum(1 for r in self.results if r.get("Status") == "Success")
        consec_fail = 0
        lock = asyncio.Lock()
        jf = jsonl_path.open("a", encoding="utf-8") if jsonl_path else None

        async def one(entry, url):
            nonlocal done, models, consec_fail
            while self.pause_evt.is_set() and not self.stop_evt.is_set():
                await asyncio.sleep(0.2)
            if self.stop_evt.is_set():
                return
            async with sem:
                if self.stop_evt.is_set():
                    return
                rec = await self._fetch_detail(client, entry, url)
            async with lock:
                self.results.append(rec)
                if jf:
                    jf.write(json.dumps(rec, ensure_ascii=False) + "\n")
                    jf.flush()
                done += 1
                ok = rec.get("Status") == "Success"
                models += 1 if ok else 0
                consec_fail = 0 if ok else consec_fail + 1
                self.log(f"[{'OK' if ok else rec.get('Status', 'MISS')}] {entry}\n")
                self.prog(done, total, models)
                if consec_fail >= CIRCUIT_BREAK and not self.stop_evt.is_set():
                    self.stop_evt.set()
                    self.log(f"[STOP] {CIRCUIT_BREAK} loi lien tiep -> ngat an toan (nghi block IP).\n")
            if delay:
                await asyncio.sleep(delay)

        try:
            await asyncio.gather(*(one(e, u) for e, u in work))
        finally:
            if jf:
                jf.close()
        if jsonl_path and not self.stop_evt.is_set():
            try:
                jsonl_path.unlink()
            except OSError:
                pass

    # -- pipelines ------------------------------------------------------
    async def run_codes(self, code_map, entries: list[str], conc, delay):
        targets = []
        for c in entries:
            hits = code_lookup(code_map, c)
            if hits:
                targets.append((c, hits[0]["url"]))
                if len(hits) > 1:
                    self.log(f"[WARN] {c} khop {len(hits)} trang, lay trang dau: {hits[0]['url']}\n")
            else:
                self.results.append(not_found(c))
        real = [(e, u) for e, u in targets]
        async with self._client(conc) as client:
            # md5 thay cho hash(): hash() bi salt ngau nhien moi process (PYTHONHASHSEED)
            # -> ten checkpoint doi moi lan chay, resume khong bao gio khop (BUG 1)
            digest = hashlib.md5("|".join(sorted(entries)).encode("utf-8")).hexdigest()[:8]
            key = f"codes_{digest}"
            await self._run_targets(client, real, conc, delay, key)

    async def run_links(self, entries: list[str], conc, delay):
        targets = [(e, detail_url(e)) for e in entries]
        async with self._client(conc) as client:
            digest = hashlib.md5("|".join(sorted(entries)).encode("utf-8")).hexdigest()[:8]
            key = f"links_{digest}"
            await self._run_targets(client, targets, conc, delay, key)

    async def run_all(self, items: list[dict], conc, delay, limit: int = 0):
        pool = items[:limit] if limit else items
        targets = [(it["slug"], it["url"]) for it in pool]
        async with self._client(conc) as client:
            await self._run_targets(client, targets, conc, delay, "all")


# ============================================================================
# Excel export (§4.6: cot Thong so gop 1 cot, khong noi cot dong)
# ============================================================================
def export_excel(results: list[dict], path: str):
    lead = ["Input", "Page", "Item No", "Status", "Thuong hieu", "Loai",
            "Gia", "Gia goc", "Con hang", "So bien the", "Tags"]
    mid = ["Thong so"]
    tail = ["Image", "So anh", "URL"]
    ordered = lead + mid + tail
    df = pd.DataFrame(results).reindex(columns=ordered)
    from openpyxl.styles import Alignment
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Data")
        ws = xw.sheets["Data"]
        spec_col = ordered.index("Thong so") + 1
        for row in ws.iter_rows(min_row=2, min_col=spec_col, max_col=spec_col):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True)
        ws.column_dimensions[ws.cell(row=1, column=spec_col).column_letter].width = 60
    return df


# ============================================================================
# GUI
# ============================================================================
class TTDCApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Trung Tam Dung Cu Crawler  |  trungtamdungcu.com")
        self.root.geometry("1000x720")
        self.root.minsize(900, 620)

        self.results: list[dict] = []
        self.log_q: queue.Queue = queue.Queue()
        self.stop_evt = threading.Event()
        self.pause_evt = threading.Event()
        self.is_running = False
        self.index_items: list[dict] = []
        self.code_map: dict = {}

        self.bg = "#f4f5f7"
        self.primary = "#0a6e3a"
        self.accent = "#0f9950"
        self.text = "#222222"

        self._styles()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._drain_logs)

    def _on_close(self):
        if self.is_running and not messagebox.askyesno("Thoat", "Dang crawl. Thoat va tat luon?"):
            return
        self.stop_evt.set()
        try:
            self.root.destroy()
        finally:
            os._exit(0)

    def _styles(self):
        st = ttk.Style()
        st.theme_use("clam")
        st.configure("TFrame", background=self.bg)
        st.configure("TLabel", background=self.bg, foreground=self.text, font=("Segoe UI", 10))
        st.configure("Title.TLabel", background=self.bg, foreground=self.primary,
                     font=("Segoe UI Semibold", 16, "bold"))
        st.configure("TRadiobutton", background=self.bg, font=("Segoe UI", 9))
        st.configure("Primary.TButton", font=("Segoe UI Semibold", 10, "bold"),
                     background=self.primary, foreground="white")
        st.map("Primary.TButton", background=[("active", self.accent)])

    def _build_ui(self):
        self.root.configure(bg=self.bg)
        header = ttk.Frame(self.root, padding=(15, 10))
        header.pack(fill=X)
        ttk.Label(header, text="TRUNG TAM DUNG CU CRAWLER", style="Title.TLabel").pack(side=LEFT)
        ttk.Label(header, text="  Haravan .js API  +  sitemap code-index",
                  font=("Segoe UI", 9, "italic"), foreground="#666").pack(side=LEFT, pady=5)

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=BOTH, expand=YES)

        left = ttk.LabelFrame(main, text=" Pipeline & Input ", padding=12, width=380)
        left.pack(side=LEFT, fill=Y, padx=(0, 10))
        left.pack_propagate(False)

        self.pipeline = StringVar(value="codes")
        ttk.Label(left, text="Chon pipeline:", font=("Segoe UI Semibold", 9)).pack(anchor=W)
        for val, txt in (("codes", "1. Theo MA san pham"),
                         ("links", "2. Theo link / URL"),
                         ("excel", "3. Theo cot Excel"),
                         ("all", "4. Crawl ALL (toan catalog)")):
            ttk.Radiobutton(left, text=txt, value=val, variable=self.pipeline,
                            command=self._on_pipeline).pack(anchor=W, pady=1)

        self.codes_frame = ttk.Frame(left)
        ttk.Label(self.codes_frame, text="Ma san pham (moi dong 1 ma):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.codes_text = Text(self.codes_frame, height=10, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.codes_text.pack(fill=BOTH, expand=YES)
        ttk.Label(self.codes_frame, text="vd: M21-ID13100CP  DCD709N-KR",
                  font=("Segoe UI", 8, "italic"), foreground="#888").pack(anchor=W)
        ttk.Button(self.codes_frame, text="Nap/lam moi index (sitemap)",
                   command=self._ensure_index_async).pack(anchor=W, pady=(6, 0))
        self.index_lbl = ttk.Label(self.codes_frame, text="Index: chua nap",
                                   foreground="#888")
        self.index_lbl.pack(anchor=W)

        self.links_frame = ttk.Frame(left)
        ttk.Label(self.links_frame, text="Link hoac slug (moi dong 1):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.links_text = Text(self.links_frame, height=10, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.links_text.pack(fill=BOTH, expand=YES)
        ttk.Label(self.links_frame, text="vd: https://trungtamdungcu.com/products/...",
                  font=("Segoe UI", 8, "italic"), foreground="#888").pack(anchor=W)

        self.excel_frame = ttk.Frame(left)
        self.xls = StringVar(); self.sheet = StringVar(); self.col = StringVar(); self.outdir = StringVar()
        attach_excel_picker(self.excel_frame, self.xls, self.sheet, self.col, self.outdir)
        self.excel_as_code = BooleanVar(value=True)
        Checkbutton(self.excel_frame, text="Cot la MA (khong tick = cot la LINK)",
                    variable=self.excel_as_code, bg=self.bg, fg=self.text,
                    font=("Segoe UI", 9), activebackground=self.bg, anchor=W).pack(fill=X, pady=(6, 0))

        self.all_frame = ttk.Frame(left)
        ttk.Label(self.all_frame, text="Quet toan bo catalog qua sitemap index.\n"
                  "Catalog RAT LON (~78,000 sp) -- gioi han so luong ben duoi.",
                  font=("Segoe UI", 9), foreground="#555").pack(anchor=W, pady=(10, 4))
        lim_row = ttk.Frame(self.all_frame); lim_row.pack(anchor=W)
        ttk.Label(lim_row, text="Gioi han (0 = het, toi da MAX_ITEMS):").pack(side=LEFT)
        self.limit = Spinbox(lim_row, from_=0, to=MAX_ITEMS, width=7, justify=CENTER)
        self.limit.delete(0, END); self.limit.insert(0, "500")
        self.limit.pack(side=LEFT, padx=6)

        speed = ttk.Frame(left); speed.pack(fill=X, pady=(12, 8), side=BOTTOM)
        ttk.Label(speed, text="Concurrency:", font=("Segoe UI Semibold", 9)).grid(row=0, column=0, sticky=W)
        self.conc = Spinbox(speed, from_=1, to=32, width=5, justify=CENTER)
        self.conc.delete(0, END); self.conc.insert(0, "8")
        self.conc.grid(row=0, column=1, sticky=E, padx=6, pady=3)
        ttk.Label(speed, text="Delay (s):", font=("Segoe UI Semibold", 9)).grid(row=1, column=0, sticky=W)
        self.delay = Spinbox(speed, from_=0.0, to=3.0, increment=0.05, width=5, justify=CENTER)
        self.delay.delete(0, END); self.delay.insert(0, "0.05")
        self.delay.grid(row=1, column=1, sticky=E, padx=6, pady=3)

        ctrl = ttk.Frame(left); ctrl.pack(fill=X, side=BOTTOM)
        self.start_btn = ttk.Button(ctrl, text="START CRAWL", style="Primary.TButton", command=self._start)
        self.start_btn.pack(fill=X, pady=(0, 6))
        sub = ttk.Frame(ctrl); sub.pack(fill=X)
        self.pause_btn = ttk.Button(sub, text="Pause", state=DISABLED, command=self._toggle_pause)
        self.pause_btn.pack(side=LEFT, fill=X, expand=YES, padx=(0, 4))
        self.stop_btn = ttk.Button(sub, text="Stop", state=DISABLED, command=self._stop)
        self.stop_btn.pack(side=RIGHT, fill=X, expand=YES, padx=(4, 0))
        self.export_btn = ttk.Button(ctrl, text="EXPORT EXCEL", state=DISABLED, command=self._export)
        self.export_btn.pack(fill=X, pady=(6, 0))

        right = ttk.LabelFrame(main, text=" Console & Progress ", padding=12)
        right.pack(side=RIGHT, fill=BOTH, expand=YES)
        info = ttk.Frame(right); info.pack(fill=X)
        self.prog_lbl = ttk.Label(info, text="Idle", font=("Segoe UI Semibold", 10))
        self.prog_lbl.pack(side=LEFT)
        self.stats_lbl = ttk.Label(info, text="0 / 0  |  0 OK", foreground="#555")
        self.stats_lbl.pack(side=RIGHT)
        self.pbar = ttk.Progressbar(right, mode="determinate")
        self.pbar.pack(fill=X, pady=(4, 10))
        self.console = Text(right, wrap=WORD, font=("Consolas", 9), bg="#1e1e1e",
                            fg="#d4d4d4", insertbackground="white", relief=FLAT)
        self.console.pack(fill=BOTH, expand=YES)
        sb = ttk.Scrollbar(self.console, command=self.console.yview)
        sb.pack(side=RIGHT, fill=Y)
        self.console.config(yscrollcommand=sb.set)

        self._on_pipeline()
        self._log("[SYS] San sang crawl trungtamdungcu.com.\n")

    def _on_pipeline(self):
        for f in (self.codes_frame, self.links_frame, self.excel_frame, self.all_frame):
            f.pack_forget()
        p = self.pipeline.get()
        {"codes": self.codes_frame, "links": self.links_frame,
         "excel": self.excel_frame, "all": self.all_frame}[p].pack(fill=BOTH, expand=YES)

    def _log(self, text):
        self.console.config(state=NORMAL)
        self.console.insert(END, text); self.console.see(END)
        self.console.config(state=DISABLED)

    def _drain_logs(self):
        while not self.log_q.empty():
            kind, data = self.log_q.get()
            if kind == "log":
                self._log(data)
            elif kind == "prog":
                done, total, models = data
                self.pbar["maximum"] = max(total, 1); self.pbar["value"] = done
                pct = int(done / total * 100) if total else 0
                self.prog_lbl.config(text=f"Crawling {pct}%")
                self.stats_lbl.config(text=f"{done} / {total}  |  {models} OK")
            elif kind == "index_done":
                self.index_items, self.code_map = data
                self.index_lbl.config(text=f"Index: {len(self.index_items)} san pham")
            elif kind == "done":
                self._finish(data)
        self.root.after(100, self._drain_logs)

    def _finish(self, aborted):
        self.is_running = False
        self.start_btn.config(state=NORMAL)
        self.pause_btn.config(state=DISABLED, text="Pause")
        self.stop_btn.config(state=DISABLED)
        self.export_btn.config(state=NORMAL if self.results else DISABLED)
        self.prog_lbl.config(text="Da dung" if aborted else "Hoan tat")
        if not aborted:
            n_ok = sum(1 for r in self.results if r.get("Status") == "Success")
            messagebox.showinfo("Xong", f"Crawl xong {n_ok}/{len(self.results)}. Bam EXPORT EXCEL.")

    def _toggle_pause(self):
        if self.pause_evt.is_set():
            self.pause_evt.clear(); self.pause_btn.config(text="Pause")
            self.log_q.put(("log", "[SYS] Tiep tuc.\n"))
        else:
            self.pause_evt.set(); self.pause_btn.config(text="Resume")
            self.log_q.put(("log", "[SYS] Tam dung.\n"))

    def _stop(self):
        if messagebox.askyesno("Dung", "Dung crawl? Ket qua hien co se duoc giu lai."):
            self.stop_evt.set(); self.pause_evt.clear()

    def _ensure_index_async(self):
        threading.Thread(target=self._build_index_thread, daemon=True).start()

    def _build_index_thread(self):
        self.log_q.put(("log", "[SYS] Dang nap sitemap index...\n"))
        items, cmap = load_or_build_index_sync(log=lambda s: self.log_q.put(("log", s)))
        self.log_q.put(("index_done", (items, cmap)))

    def _gather(self, widget):
        raw = widget.get("1.0", "end-1c")
        return [x.strip() for x in re.split(r"[\r\n,;\t]+", raw) if x.strip()]

    def _gather_excel(self):
        try:
            df = pd.read_excel(self.xls.get(), sheet_name=self.sheet.get())
        except Exception as e:
            self.log_q.put(("log", f"[ERR] {e}\n")); return []
        col = self.col.get()
        if col not in df.columns:
            self.log_q.put(("log", f"[WARN] Khong thay cot '{col}'.\n")); return []
        return [str(v).strip() for v in df[col].dropna() if str(v).strip()]

    def _start(self):
        try:
            conc = int(self.conc.get()); delay = float(self.delay.get())
        except ValueError:
            messagebox.showerror("Loi", "Concurrency phai la so nguyen, Delay la so thuc."); return
        pipeline = self.pipeline.get()
        if pipeline == "codes" and not self.code_map:
            messagebox.showwarning("Chua co index", "Bam 'Nap/lam moi index' truoc.")
            return
        self.results = []
        self.stop_evt.clear(); self.pause_evt.clear()
        self.is_running = True
        self.start_btn.config(state=DISABLED)
        self.pause_btn.config(state=NORMAL); self.stop_btn.config(state=NORMAL)
        self.export_btn.config(state=DISABLED)
        self.console.config(state=NORMAL); self.console.delete("1.0", END); self.console.config(state=DISABLED)
        threading.Thread(target=self._run_async, args=(pipeline, conc, delay), daemon=True).start()

    def _run_async(self, pipeline, conc, delay):
        try:
            asyncio.run(self._crawl(pipeline, conc, delay))
        except Exception as e:
            self.log_q.put(("log", f"[FATAL] {e}\n"))
            self.log_q.put(("done", True))
            return
        self.log_q.put(("done", self.stop_evt.is_set()))

    async def _crawl(self, pipeline, conc, delay):
        eng = TTDCEngine(log=lambda s: self.log_q.put(("log", s)),
                         prog=lambda d, t, m: self.log_q.put(("prog", (d, t, m))),
                         stop_evt=self.stop_evt, pause_evt=self.pause_evt)
        if pipeline == "codes":
            codes = self._gather(self.codes_text)
            if not codes:
                self.log_q.put(("log", "[WARN] Chua nhap ma nao.\n")); return
            await eng.run_codes(self.code_map, codes, conc, delay)
        elif pipeline == "links":
            links = self._gather(self.links_text)
            if not links:
                self.log_q.put(("log", "[WARN] Chua nhap link nao.\n")); return
            await eng.run_links(links, conc, delay)
        elif pipeline == "excel":
            vals = self._gather_excel()
            if not vals:
                return
            if self.excel_as_code.get():
                if not self.code_map:
                    self.log_q.put(("log", "[SYS] Chua co index -> tu nap...\n"))
                    items, cmap = await load_or_build_index_async(
                        log=lambda s: self.log_q.put(("log", s)))
                    self.index_items, self.code_map = items, cmap
                    self.log_q.put(("index_done", (items, cmap)))
                await eng.run_codes(self.code_map, vals, conc, delay)
            else:
                await eng.run_links(vals, conc, delay)
        elif pipeline == "all":
            if not self.index_items:
                self.log_q.put(("log", "[SYS] Chua co index -> tu nap...\n"))
                items, cmap = await load_or_build_index_async(
                    log=lambda s: self.log_q.put(("log", s)))
                self.index_items, self.code_map = items, cmap
                self.log_q.put(("index_done", (items, cmap)))
            limit = int(self.limit.get() or 0)
            await eng.run_all(self.index_items, conc, delay, limit)
        self.results = eng.results

    def _export(self):
        if not self.results:
            messagebox.showwarning("Trong", "Chua co du lieu."); return
        path = ask_export_path(self.xls.get(), HERE, "trungtamdungcu_ketqua.xlsx")
        if not path:
            return
        try:
            export_excel(self.results, path)
            self._log(f"[SYS] Da xuat -> {path}\n")
            messagebox.showinfo("Xong", f"Da luu:\n{path}")
        except Exception as e:
            messagebox.showerror("Loi xuat", str(e))


# ============================================================================
# CLI / headless entry
# ============================================================================
def _cli_codes(codes: list[str], conc: int, delay: float) -> list[dict]:
    items, cmap = load_or_build_index_sync(log=lambda s: print(s, end="", flush=True))
    eng = TTDCEngine(log=lambda s: print(s, end="", flush=True))
    asyncio.run(eng.run_codes(cmap, codes, conc, delay))
    return eng.results


def _cli_links(links: list[str], conc: int, delay: float) -> list[dict]:
    eng = TTDCEngine(log=lambda s: print(s, end="", flush=True))
    asyncio.run(eng.run_links(links, conc, delay))
    return eng.results


def main():
    ap = argparse.ArgumentParser(description="Trung Tam Dung Cu crawler (GUI mac dinh)")
    ap.add_argument("--code", action="append", help="ma san pham (lap lai de nhieu ma)")
    ap.add_argument("--link", action="append", help="link/slug san pham")
    ap.add_argument("--conc", type=int, default=8)
    ap.add_argument("--delay", type=float, default=0.05)
    ap.add_argument("--export", help="duong dan xlsx xuat ket qua")
    a = ap.parse_args()

    if a.code or a.link:
        results = []
        if a.code:
            results += _cli_codes(a.code, a.conc, a.delay)
        if a.link:
            results += _cli_links(a.link, a.conc, a.delay)
        for r in results:
            print(f"{r.get('Status')}: {r.get('Input')} -> {r.get('Page')}")
        if a.export:
            export_excel(results, a.export)
            print(f"Exported -> {a.export}")
        return

    root = Tk()
    TTDCApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
