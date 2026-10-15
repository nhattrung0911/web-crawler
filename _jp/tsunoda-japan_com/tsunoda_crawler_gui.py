"""
Tsunoda Tools Product Crawler  (tsunoda-japan.com)
==================================================
Modern async crawler: httpx (pooled HTTP) + selectolax (Lexbor parser).
Same engine family as the ARM Sangyo / AnexTool crawlers.

Site recon (tsunoda-japan.com):
  * Static HTML site. Bilingual: JP under /itempage/*.html, EN under
    /EN/itempage/*.html.  Product page URL slug == the product code
    (TP-8 -> /EN/itempage/TP-8.html)  -> O(1) direct lookup (Anex-style).
  * sitemap.xml exists but is STALE (2017) and incomplete -> do NOT rely on
    it for discovery.  Instead BFS the category / brand hub pages, which link
    to every /itempage/*.html  (ARM-style discovery).
  * A single itempage often lists SEVERAL models as spec-table ROWS
    (CN-125.html title "CN-125/150" -> rows CN-125, CN-150).  So a requested
    code may live as a row inside another product's page -> crawl + filter.
  * Spec tables are Bootstrap <table class="table table-sm table-bordered">
    with MERGED headers (rowspan=2 code col, colspan=6 "Size[mm]" over
    sub-headers).  Parsed via a proper rowspan/colspan grid builder.
  * Multiple bordered tables per page describe the same models from different
    angles (dimensions / cutting-capacity) -> merged by Item No.
  * Pages are UTF-8.

Four user-selectable pipelines (mirrors the ARM crawler):
  1) Model code  - paste codes -> direct lookup + BFS-and-filter (finds codes
                   that only exist as a table row on another page)
  2) Slug / URL  - paste product codes or full /itempage/ URLs (all models)
  3) Crawl ALL   - BFS-discover every itempage, extract all models
  4) Excel column- pick file -> sheet -> column of codes / URLs
"""

import os
import re
import sys
import queue
import asyncio
import threading
from urllib.parse import urljoin, urlparse
from tkinter import *
from tkinter import ttk, filedialog, messagebox

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser

BASE = "https://www.tsunoda-japan.com"
HOST = "www.tsunoda-japan.com"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# hub pages that link out to /itempage/*.html (both languages).  Seed set for BFS.
HUB_SEEDS = [
    "/", "/index.html", "/newproducts.html",
    "/crimp.html", "/nippers.html", "/nippers_slant.html", "/sidecutpl.html",
    "/metalcutter.html", "/powercutter.html", "/longnosepliers.html",
    "/plastic_nippers.html", "/cablecutter.html", "/wireropecutter.html",
    "/pliers.html", "/snapring.html", "/hexkey.html", "/crafttools.html",
    "/specialtools.html", "/supplies.html",
    "/brand_tsunoda.html", "/brand_power.html", "/brand_stainless.html",
    "/brand_trinity.html", "/brand_kingttc.html", "/brand_ttc.html",
    "/brand_mini.html", "/brand_craft.html", "/brand_hexkey.html",
    "/brand_perfect.html",
]

# ---- anti-loop / safety limits ----
MAX_REDIRECTS = 5
CIRCUIT_BREAK = 25          # abort after N consecutive page failures (IP block)
MAX_ITEMS = 5000            # runaway guard on work-list size
MAX_BFS_PAGES = 1500        # runaway guard on hub discovery


# ----------------------------------------------------------------------------
# Core helpers (pure, no GUI)
# ----------------------------------------------------------------------------
def norm_code(s: str) -> str:
    """Normalize a model code for matching: drop whitespace, upper-case."""
    return re.sub(r"\s+", "", str(s)).upper()


def code_from_url(url: str) -> str:
    """/EN/itempage/CN-125.html -> CN-125"""
    base = url.rstrip("/").rsplit("/", 1)[-1]
    return re.sub(r"\.html?$", "", base, flags=re.I)


def item_url_candidates(entry: str, lang: str = "EN") -> list:
    """Accept a code, an *.html name, an /itempage/ path or full URL ->
    ordered list of candidate URLs (preferred language first)."""
    e = entry.strip()
    if e.startswith("http"):
        return [e]
    if e.startswith("/"):
        return [BASE + e]
    code = re.sub(r"\.html?$", "", e, flags=re.I)
    en = f"{BASE}/EN/itempage/{code}.html"
    jp = f"{BASE}/itempage/{code}.html"
    return [en, jp] if lang == "EN" else [jp, en]


def _full_img(src: str, page_url: str) -> str:
    if not src:
        return ""
    return urljoin(page_url, src)


def _pick_product_image(tree, page_url: str) -> str:
    """First /img/products/ photo that is NOT the /size/ dimension diagram."""
    fallback = ""
    for img in tree.css("img"):
        src = img.attributes.get("src") or img.attributes.get("data-src") or ""
        low = src.lower()
        if "/img/products/" not in low:
            continue
        url = _full_img(src, page_url)
        if "/products/size/" in low:      # dimension diagram -> only a fallback
            fallback = fallback or url
            continue
        return url
    return fallback


def page_name(tree) -> str:
    """Product name from <title> ('CODE<full-width space>Name：：：Tsunoda...')."""
    t = tree.css_first("title")
    if not t:
        h1 = tree.css_first("h1")
        return h1.text(strip=True) if h1 else ""
    txt = t.text(strip=True)
    txt = txt.split("：：：")[0].strip()          # drop the ：：：Tsunoda... suffix
    txt = re.sub(r"\s+", " ", txt.replace("　", " ")).strip()
    return txt


# ---- spec-table grid parsing ------------------------------------------------
def _cell_text(node) -> str:
    return re.sub(r"\s+", " ", node.text(separator=" ", strip=True)).strip()


def build_grid(table) -> list:
    """Expand an HTML table (honouring rowspan/colspan) into a rectangular
    matrix of cell strings."""
    grid = []
    carry = {}          # col_index -> (text, remaining_rows)
    for tr in table.css("tr"):
        cells = list(tr.css("td, th"))
        if not cells and not carry:
            continue
        row = {}
        cq = list(cells)
        c = 0
        while cq or any(k >= c for k in carry):
            if c in carry:
                text, rem = carry[c]
                row[c] = text
                if rem - 1 > 0:
                    carry[c] = (text, rem - 1)
                else:
                    del carry[c]
                c += 1
                continue
            if cq:
                cell = cq.pop(0)
                cs = int(cell.attributes.get("colspan") or 1)
                rs = int(cell.attributes.get("rowspan") or 1)
                txt = _cell_text(cell)
                for k in range(cs):
                    row[c + k] = txt
                    if rs > 1:
                        carry[c + k] = (txt, rs - 1)
                c += cs
                continue
            break
        width = (max(row) + 1) if row else 0
        grid.append([row.get(i, "") for i in range(width)])
    w = max((len(r) for r in grid), default=0)
    return [r + [""] * (w - len(r)) for r in grid]


def _is_data_row(cells: list) -> bool:
    """A row is data if it has a non-empty first cell and most other cells
    carry digits (measurements / EAN)."""
    if not cells or not cells[0].strip():
        return False
    rest = cells[1:]
    if not rest:
        return False
    nums = sum(1 for c in rest if re.search(r"\d", c))
    return nums >= max(1, len(rest) // 2)


def parse_spec_table(grid: list) -> list:
    """Grid -> list of (item_no, {label: value}) records."""
    if len(grid) < 2:
        return []
    first_data = next((i for i, r in enumerate(grid) if _is_data_row(r)), None)
    if first_data is None or first_data == 0:
        header, data = grid[:1], grid[1:]      # fallback: row 0 is the header
    else:
        header, data = grid[:first_data], grid[first_data:]

    width = len(grid[0])
    labels = []
    for j in range(width):
        parts = []
        for hr in header:
            v = hr[j].strip() if j < len(hr) else ""
            if v and v not in parts:
                parts.append(v)
        label = " ".join(parts)
        if j == 0 and not label:
            label = "Item No"
        labels.append(label or f"Col{j+1}")
    labels[0] = "Item No"

    out = []
    for r in data:
        item = r[0].strip()
        if not item:
            continue
        specs = {}
        for j, val in enumerate(r):
            if j == 0:
                continue
            lab = labels[j]
            v = val.strip()
            if v:
                specs.setdefault(lab, v)
        out.append((item, specs))
    return out


def parse_page(html: str, entry: str, url: str) -> list:
    """Parse a Tsunoda itempage -> one record per model (rows merged across the
    bordered spec tables by Item No)."""
    tree = LexborHTMLParser(html)
    name = page_name(tree)
    img = _pick_product_image(tree, url)

    order = []                 # keep model order stable
    models = {}                # norm_code -> record
    for table in tree.css("table.table.table-bordered, table.table-bordered"):
        grid = build_grid(table)
        for item, specs in parse_spec_table(grid):
            key = norm_code(item)
            if key not in models:
                models[key] = {"Input": entry, "Page": name, "Item No": item,
                               "Image": img, "URL": url, "Status": "Success"}
                order.append(key)
            for k, v in specs.items():
                models[key].setdefault(k, v)

    if not models:
        return [{"Input": entry, "Page": name, "Item No": "",
                 "Image": img, "URL": url, "Status": "No models"}]
    return [models[k] for k in order]


def extract_itempage_links(html: str, page_url: str) -> tuple:
    """Return (itempage_urls, hub_urls) found on a page.

    itempage_urls -> /itempage/*.html or /EN/itempage/*.html product pages
    hub_urls      -> other same-host *.html (category/brand) pages to follow
    """
    tree = LexborHTMLParser(html)
    items, hubs = [], []
    for a in tree.css("a"):
        href = a.attributes.get("href") or ""
        if not href or href.startswith("#") or href.startswith("mailto:") \
                or href.startswith("tel:") or href.startswith("javascript:"):
            continue
        full = urljoin(page_url, href)
        p = urlparse(full)
        if p.netloc != HOST or not p.path.lower().endswith((".html", "/")):
            continue
        clean = f"{BASE}{p.path}"
        if "/itempage/" in p.path.lower():
            items.append(clean)
        elif p.path.lower().endswith(".html"):
            hubs.append(clean)
    return items, hubs


# ----------------------------------------------------------------------------
# Async crawl engine  (usable head-less; GUI wraps it)
# ----------------------------------------------------------------------------
class TsunodaEngine:
    """Pure async crawl logic.  `log`/`prog` are callables (default: silent)."""

    def __init__(self, log=None, prog=None, stop_evt=None, pause_evt=None):
        self.log = log or (lambda s: None)
        self.prog = prog or (lambda d, t, m: None)
        self.stop_evt = stop_evt or threading.Event()
        self.pause_evt = pause_evt or threading.Event()
        self.results = []

    def _client(self, conc):
        limits = httpx.Limits(max_connections=conc, max_keepalive_connections=conc)
        transport = httpx.AsyncHTTPTransport(retries=2)
        return httpx.AsyncClient(headers={"User-Agent": UA}, timeout=20.0,
                                 limits=limits, transport=transport,
                                 follow_redirects=True, max_redirects=MAX_REDIRECTS)

    async def _get(self, client, url):
        r = await client.get(url)
        final = str(r.url)
        if r.status_code != 200:
            raise httpx.HTTPStatusError(f"HTTP {r.status_code}",
                                        request=r.request, response=r)
        return r.content.decode("utf-8", "replace"), final

    async def discover(self, client, conc, lang):
        """BFS the hub pages -> every itempage URL (preferred lang per code)."""
        self.log("[SYS] BFS phat hien trang san pham (category/brand)...\n")
        seeds = [BASE + h for h in HUB_SEEDS] + [BASE + "/EN" + h for h in HUB_SEEDS]
        frontier = {s for s in seeds}
        visited = set()
        found = {}                       # norm_code -> url (prefer EN)
        sem = asyncio.Semaphore(conc)

        async def grab(url):
            async with sem:
                try:
                    html, final = await self._get(client, url)
                    return extract_itempage_links(html, final)
                except Exception:
                    return [], []

        hubs_seen = 0
        while frontier and hubs_seen < MAX_BFS_PAGES and not self.stop_evt.is_set():
            batch = [u for u in frontier if u not in visited]
            visited.update(batch)
            hubs_seen += len(batch)
            frontier = set()
            prefer_en = (lang == "EN")
            for items, hubs in await asyncio.gather(*(grab(u) for u in batch)):
                for it in items:
                    key = norm_code(code_from_url(it))
                    is_en = "/EN/itempage/" in it
                    cur = found.get(key)
                    if cur is None:
                        found[key] = it
                    else:
                        cur_en = "/EN/itempage/" in cur
                        if prefer_en and is_en and not cur_en:
                            found[key] = it       # upgrade JP -> EN
                        elif not prefer_en and not is_en and cur_en:
                            found[key] = it       # upgrade EN -> JP
                for hb in hubs:
                    if hb not in visited and "/itempage/" not in hb:
                        frontier.add(hb)
        pages = sorted(set(found.values()))
        self.log(f"[SYS] Tim thay {len(pages)} trang san pham.\n")
        return pages

    async def _process(self, client, url, entry=None):
        entry = entry or url
        try:
            html, final = await self._get(client, url)
            if "/itempage/" not in final:
                return [{"Input": entry, "URL": final, "Status": "Redirect off /itempage/"}], False
            recs = parse_page(html, entry, final)
            ok = any(r.get("Status") == "Success" for r in recs)
            return recs, ok
        except httpx.TooManyRedirects:
            return [{"Input": entry, "URL": url, "Status": "Redirect loop"}], False
        except Exception as e:
            return [{"Input": entry, "URL": url, "Status": f"Error: {type(e).__name__}"}], False

    async def _try_candidates(self, client, entry, lang):
        """Direct lookup: first candidate URL that yields models wins."""
        last = None
        for url in item_url_candidates(entry, lang):
            recs, ok = await self._process(client, url, entry)
            if ok:
                return recs, True
            last = recs
        return last or [{"Input": entry, "Status": "Not Found"}], False

    # -- pipelines -----------------------------------------------------------
    async def run(self, pipeline, codes, conc, delay, lang="EN"):
        self.results = []
        async with self._client(conc) as client:
            if pipeline == "model":
                await self._run_model(client, codes, conc, delay, lang)
            elif pipeline == "codes":
                await self._run_direct(client, codes, conc, delay, lang, filter_set=None)
            elif pipeline == "all":
                pages = await self.discover(client, conc, lang)
                await self._crawl_pages(client, pages, conc, delay, filter_set=None)
        return self.results

    async def _run_direct(self, client, codes, conc, delay, lang, filter_set):
        work = []
        for c in codes:
            work.append(c)
        total = len(work)
        if not total:
            self.log("[WARN] Khong co ma/URL nao.\n")
            return
        self.log(f"[SYS] Direct lookup {total} muc | concurrency={conc}\n")
        sem = asyncio.Semaphore(conc)
        done = models = 0
        lock = asyncio.Lock()

        async def one(entry):
            nonlocal done, models
            while self.pause_evt.is_set() and not self.stop_evt.is_set():
                await asyncio.sleep(0.2)
            if self.stop_evt.is_set():
                return
            async with sem:
                if self.stop_evt.is_set():
                    return
                recs, ok = await self._try_candidates(client, entry, lang)
            if filter_set is not None:
                recs = [r for r in recs if norm_code(r.get("Item No", "")) in filter_set]
            async with lock:
                self.results.extend(recs)
                done += 1
                n = sum(1 for r in recs if r.get("Status") == "Success")
                models += n
                tag = f"OK  {n} model" if n else (recs[0].get("Status", "MISS") if recs else "MISS")
                self.log(f"[{tag}] {entry}\n")
                self.prog(done, total, models)
            if delay:
                await asyncio.sleep(delay)

        await asyncio.gather(*(one(e) for e in work))

    async def _crawl_pages(self, client, pages, conc, delay, filter_set):
        total = len(pages)
        if not total:
            self.log("[WARN] Khong co trang nao de crawl.\n")
            return
        if total > MAX_ITEMS:
            self.log(f"[WARN] {total} trang vuot {MAX_ITEMS}, cat bot.\n")
            pages = pages[:MAX_ITEMS]; total = MAX_ITEMS
        self.log(f"[SYS] Crawl {total} trang | concurrency={conc}\n")
        sem = asyncio.Semaphore(conc)
        done = models = consec_fail = 0
        lock = asyncio.Lock()

        async def one(url):
            nonlocal done, models, consec_fail
            while self.pause_evt.is_set() and not self.stop_evt.is_set():
                await asyncio.sleep(0.2)
            if self.stop_evt.is_set():
                return
            async with sem:
                if self.stop_evt.is_set():
                    return
                recs, ok = await self._process(client, url)
            if filter_set is not None:
                recs = [r for r in recs if norm_code(r.get("Item No", "")) in filter_set]
            async with lock:
                if recs:
                    self.results.extend(recs)
                done += 1
                n = sum(1 for r in recs if r.get("Status") == "Success")
                models += n
                consec_fail = 0 if ok else consec_fail + 1
                if filter_set is None:
                    tag = f"OK  {n} model" if ok else (recs[0].get("Status", "MISS") if recs else "MISS")
                    self.log(f"[{tag}] {code_from_url(url)}\n")
                elif n:
                    self.log(f"[HIT {n}] {code_from_url(url)}\n")
                self.prog(done, total, models)
                if consec_fail >= CIRCUIT_BREAK and not self.stop_evt.is_set():
                    self.stop_evt.set()
                    self.log(f"[STOP] {CIRCUIT_BREAK} loi lien tiep -> ngat an toan.\n")
            if delay:
                await asyncio.sleep(delay)

        await asyncio.gather(*(one(u) for u in pages))

    async def _run_model(self, client, codes, conc, delay, lang):
        """Model search: direct lookup first, then BFS-and-filter for the rest."""
        match_set = {}
        for c in codes:
            match_set.setdefault(norm_code(c), c)
        self.log(f"[SYS] Tim {len(match_set)} ma model.\n")

        # 1) direct per-code lookup (fast path; a code with its own page)
        await self._run_direct(client, list(match_set.values()), conc, delay, lang,
                               filter_set=match_set)
        found = {norm_code(r.get("Item No", "")) for r in self.results
                 if r.get("Status") == "Success"}
        missing = {k: v for k, v in match_set.items() if k not in found}

        # 2) codes still missing may live as a ROW inside another page -> BFS
        if missing and not self.stop_evt.is_set():
            self.log(f"[SYS] Con {len(missing)} ma chua thay -> BFS toan site.\n")
            pages = await self.discover(client, conc, lang)
            before = len(self.results)
            await self._crawl_pages(client, pages, conc, delay, filter_set=missing)
            for r in self.results[before:]:
                found.add(norm_code(r.get("Item No", "")))

        # drop the placeholder "Not Found" rows the direct pass produced for
        # codes that were later matched via BFS
        self.results = [r for r in self.results
                        if not (r.get("Status") == "Not Found"
                                and norm_code(r.get("Item No", "")) in found)]

        for k, orig in match_set.items():
            if k not in found:
                self.results.append({"Input": orig, "Item No": orig, "Status": "Not Found"})
                self.log(f"[MISS] {orig} khong thay tren site.\n")
        self.log(f"[SYS] Khop {len(found & set(match_set))}/{len(match_set)} ma.\n")


# ----------------------------------------------------------------------------
# Excel export (shared by GUI + head-less test)
# ----------------------------------------------------------------------------
def export_excel(results, path):
    lead = ["Input", "Page", "Item No", "Status"]
    tail = ["Image", "URL"]
    spec_cols = []
    for r in results:
        for k in r:
            if k not in lead and k not in tail and k not in spec_cols:
                spec_cols.append(k)
    ordered = lead + spec_cols + tail
    df = pd.DataFrame(results).reindex(columns=ordered)
    df.to_excel(path, index=False)
    return df


# ----------------------------------------------------------------------------
# GUI application  (same skeleton as the ARM / Anex crawlers)
# ----------------------------------------------------------------------------
class TsunodaCrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tsunoda Tools Crawler  |  tsunoda-japan.com")
        self.root.geometry("980x700")
        self.root.minsize(880, 600)

        self.results = []
        self.log_q = queue.Queue()
        self.stop_evt = threading.Event()
        self.pause_evt = threading.Event()
        self.file_rows = []
        self.is_running = False

        self.bg = "#f4f5f7"
        self.primary = "#c8102e"       # Tsunoda red
        self.accent = "#e2453f"
        self.text = "#222222"

        self._styles()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._drain_logs)

    def _on_close(self):
        if self.is_running:
            if not messagebox.askyesno("Thoat", "Dang crawl. Thoat va tat luon?"):
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
        st.configure("TButton", font=("Segoe UI Semibold", 10))
        st.configure("Primary.TButton", font=("Segoe UI Semibold", 10, "bold"),
                     background=self.primary, foreground="white")
        st.map("Primary.TButton", background=[("active", self.accent)])

    def _build_ui(self):
        self.root.configure(bg=self.bg)
        header = ttk.Frame(self.root, padding=(15, 10))
        header.pack(fill=X)
        ttk.Label(header, text="TSUNODA TOOLS CRAWLER", style="Title.TLabel").pack(side=LEFT)
        ttk.Label(header, text="  httpx + selectolax  •  BFS + slug-direct",
                  font=("Segoe UI", 9, "italic"), foreground="#666").pack(side=LEFT, pady=5)

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=BOTH, expand=YES)

        left = ttk.LabelFrame(main, text=" Pipeline & Input ", padding=12, width=360)
        left.pack(side=LEFT, fill=Y, padx=(0, 10))
        left.pack_propagate(False)

        self.pipeline = StringVar(value="model")
        ttk.Label(left, text="Chon pipeline:", font=("Segoe UI Semibold", 9)).pack(anchor=W)
        for val, txt in (("model", "1. Theo MA MODEL (Item No)"),
                         ("codes", "2. Theo ma / URL trang item"),
                         ("all",   "3. Crawl ALL (BFS toan site)"),
                         ("excel", "4. Theo cot trong Excel")):
            ttk.Radiobutton(left, text=txt, value=val, variable=self.pipeline,
                            command=self._on_pipeline).pack(anchor=W, pady=1)

        # language
        lang_row = ttk.Frame(left); lang_row.pack(fill=X, pady=(6, 0))
        ttk.Label(lang_row, text="Ngon ngu:", font=("Segoe UI Semibold", 9)).pack(side=LEFT)
        self.lang = StringVar(value="EN")
        ttk.Radiobutton(lang_row, text="EN", value="EN", variable=self.lang).pack(side=LEFT, padx=4)
        ttk.Radiobutton(lang_row, text="JP", value="JP", variable=self.lang).pack(side=LEFT)

        self.model_frame = ttk.Frame(left)
        ttk.Label(self.model_frame, text="Ma model (moi dong 1 ma):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.model_text = Text(self.model_frame, height=10, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.model_text.pack(fill=BOTH, expand=YES)
        ttk.Label(self.model_frame, text="vd: CN-125  CN-150N  PCN-125TG\n(lookup + quet toan site roi loc dung ma)",
                  font=("Segoe UI", 8, "italic"), foreground="#888").pack(anchor=W)

        self.codes_frame = ttk.Frame(left)
        ttk.Label(self.codes_frame, text="Ma hoac URL (moi dong 1 muc):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.codes_text = Text(self.codes_frame, height=10, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.codes_text.pack(fill=BOTH, expand=YES)
        ttk.Label(self.codes_frame, text="vd: TP-8  /  /EN/itempage/CN-125.html",
                  font=("Segoe UI", 8, "italic"), foreground="#888").pack(anchor=W)

        self.excel_frame = ttk.Frame(left)
        top = ttk.Frame(self.excel_frame); top.pack(fill=X, pady=(8, 4))
        ttk.Label(top, text="File Excel:", font=("Segoe UI Semibold", 9)).pack(side=LEFT)
        ttk.Button(top, text="Chon file...", width=12, command=self._browse_excel).pack(side=RIGHT)
        self.excel_cfg = ttk.Frame(self.excel_frame); self.excel_cfg.pack(fill=X)
        ttk.Label(self.excel_cfg, text="Chua chon file",
                  font=("Segoe UI", 9, "italic"), foreground="#888").pack(pady=8)
        self.excel_as_model = BooleanVar(value=True)
        Checkbutton(self.excel_frame, text="Cot la MA MODEL (tim theo ma)",
                    variable=self.excel_as_model, bg=self.bg, fg=self.text,
                    font=("Segoe UI", 9), activebackground=self.bg,
                    anchor=W).pack(fill=X, pady=(4, 0))

        self.all_frame = ttk.Frame(left)
        ttk.Label(self.all_frame,
                  text="BFS toan bo trang /itempage/*.html\nroi boc moi model trong bang.",
                  font=("Segoe UI", 9), foreground="#555").pack(anchor=W, pady=10)

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
        self.stats_lbl = ttk.Label(info, text="0 / 0 trang  |  0 model", foreground="#555")
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
        self._log("[SYS] San sang crawl tsunoda-japan.com.\n")

    def _on_pipeline(self):
        for f in (self.model_frame, self.codes_frame, self.excel_frame, self.all_frame):
            f.pack_forget()
        p = self.pipeline.get()
        if p == "model":
            self.model_frame.pack(fill=BOTH, expand=YES)
        elif p == "codes":
            self.codes_frame.pack(fill=BOTH, expand=YES)
        elif p == "excel":
            self.excel_frame.pack(fill=X)
        else:
            self.all_frame.pack(fill=X)

    def _browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xls")])
        if not path:
            return
        for w in self.excel_cfg.winfo_children():
            w.destroy()
        self.file_rows = []
        try:
            sheets = pd.ExcelFile(path).sheet_names
        except Exception as e:
            self._log(f"[ERR] Khong doc duoc file: {e}\n")
            return
        ttk.Label(self.excel_cfg, text=os.path.basename(path),
                  font=("Segoe UI Semibold", 8), foreground=self.primary).pack(anchor=W)
        row = ttk.Frame(self.excel_cfg); row.pack(fill=X, pady=4)
        ttk.Label(row, text="Sheet:", width=6).pack(side=LEFT)
        sheet_cb = ttk.Combobox(row, state="readonly", width=16, values=sheets); sheet_cb.pack(side=LEFT, padx=4)
        row2 = ttk.Frame(self.excel_cfg); row2.pack(fill=X, pady=4)
        ttk.Label(row2, text="Cot:", width=6).pack(side=LEFT)
        col_cb = ttk.Combobox(row2, state="readonly", width=16); col_cb.pack(side=LEFT, padx=4)
        rowdata = {"path": path, "sheet": sheet_cb, "col": col_cb}
        self.file_rows = [rowdata]
        sheet_cb.bind("<<ComboboxSelected>>", lambda e: self._load_cols(rowdata))
        if sheets:
            sheet_cb.set(sheets[0]); self._load_cols(rowdata)
        self._log(f"[SYS] Da nap file: {os.path.basename(path)}\n")

    def _load_cols(self, rowdata):
        try:
            df = pd.read_excel(rowdata["path"], sheet_name=rowdata["sheet"].get(), nrows=5)
        except Exception as e:
            self._log(f"[ERR] Doc sheet loi: {e}\n"); return
        cols = list(df.columns)
        rowdata["col"]["values"] = cols
        pick = next((c for c in cols if re.search(r"ma|code|item|sku|品番|型式", str(c), re.I)),
                    cols[0] if cols else "")
        if pick:
            rowdata["col"].set(pick)

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
                self.pbar["maximum"] = total; self.pbar["value"] = done
                pct = int(done / total * 100) if total else 0
                self.prog_lbl.config(text=f"Crawling {pct}%")
                self.stats_lbl.config(text=f"{done} / {total} trang  |  {models} model")
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
            n_model = sum(1 for r in self.results if r.get("Status") == "Success")
            messagebox.showinfo("Xong", f"Crawl xong {n_model} model. Bam EXPORT EXCEL.")

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

    def _start(self):
        try:
            conc = int(self.conc.get()); delay = float(self.delay.get())
        except ValueError:
            messagebox.showerror("Loi", "Concurrency phai la so nguyen, Delay la so thuc."); return
        self.results = []
        self.stop_evt.clear(); self.pause_evt.clear()
        self.is_running = True
        self.start_btn.config(state=DISABLED)
        self.pause_btn.config(state=NORMAL); self.stop_btn.config(state=NORMAL)
        self.export_btn.config(state=DISABLED)
        self.console.config(state=NORMAL); self.console.delete("1.0", END); self.console.config(state=DISABLED)
        threading.Thread(target=self._run_async,
                         args=(self.pipeline.get(), conc, delay), daemon=True).start()

    def _gather_codes(self, widget):
        raw = widget.get("1.0", "end-1c")
        return [x.strip() for x in re.split(r"[\r\n,;\t]+", raw) if x.strip()]

    def _gather_excel(self):
        out = []
        for r in self.file_rows:
            try:
                df = pd.read_excel(r["path"], sheet_name=r["sheet"].get())
            except Exception as e:
                self.log_q.put(("log", f"[ERR] {e}\n")); continue
            col = r["col"].get()
            if col in df.columns:
                out += [str(v).strip() for v in df[col].dropna() if str(v).strip()]
            else:
                self.log_q.put(("log", f"[WARN] Khong thay cot '{col}'.\n"))
        return out

    def _run_async(self, pipeline, conc, delay):
        try:
            asyncio.run(self._crawl(pipeline, conc, delay))
        except Exception as e:
            self.log_q.put(("log", f"[FATAL] {e}\n"))
            self.log_q.put(("done", True))

    async def _crawl(self, pipeline, conc, delay):
        eng = TsunodaEngine(log=lambda s: self.log_q.put(("log", s)),
                            prog=lambda d, t, m: self.log_q.put(("prog", (d, t, m))),
                            stop_evt=self.stop_evt, pause_evt=self.pause_evt)
        lang = self.lang.get()

        if pipeline == "model":
            codes = self._gather_codes(self.model_text)
        elif pipeline == "codes":
            codes = self._gather_codes(self.codes_text)
        elif pipeline == "excel":
            vals = self._gather_excel()
            pipeline = "model" if self.excel_as_model.get() else "codes"
            codes = vals
        else:
            codes = []

        if pipeline in ("model", "codes") and not codes:
            self.log_q.put(("log", "[WARN] Chua nhap ma nao.\n"))
            self.log_q.put(("done", True)); return

        await eng.run(pipeline, codes, conc, delay, lang)
        self.results = eng.results
        self.log_q.put(("done", self.stop_evt.is_set()))

    def _export(self):
        if not self.results:
            messagebox.showwarning("Trong", "Chua co du lieu."); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")],
                                            initialfile="tsunoda_data.xlsx")
        if not path:
            return
        try:
            export_excel(self.results, path)
            self._log(f"[SYS] Da xuat -> {path}\n")
            messagebox.showinfo("Xong", f"Da luu:\n{path}")
        except Exception as e:
            messagebox.showerror("Loi xuat", str(e))


if __name__ == "__main__":
    root = Tk()
    TsunodaCrawlerApp(root)
    root.mainloop()
