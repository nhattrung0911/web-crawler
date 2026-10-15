"""
Pony / Jorgensen Clamp Crawler  (ponyjorgensen.com)
===================================================
Modern async crawler: httpx (pooled HTTP) + selectolax (Lexbor parser).
Same engine family as the ARM / Anex / Tsunoda crawlers.

Site recon (ponyjorgensen.com):
  * WordPress (Yoast) but NOT WooCommerce -> custom post type "product" is NOT
    exposed over the REST API (/wp-json/wp/v2/product = 404).  No product JSON.
  * JSON-LD present but only Yoast WebPage / BreadcrumbList -> no Product schema
    with sku/offers.  So specs MUST be parsed from HTML.
  * Discovery = the Yoast product sitemap (modern, complete):
        https://www.ponyjorgensen.com/product-sitemap.xml   (~162 products)
    Product URL:  /product/{slug}/   (slug = descriptive name, NOT the code).
  * Each product page = ONE spec table.  First column header "Product #" == the
    model / SKU (#102, #103 ...).  Remaining columns vary per product
    (Max Opening, Clamping Force, Band Length ...) -> dynamic spec columns.
    Several models per page (one table row each)  -> a requested code lives as a
    ROW, so model-search = crawl the catalog and filter by "Product #".
  * Name  = <h1> (contains a <span class="brand"> + the family name).
    Brand = span.brand ;  Category = last breadcrumb (JSON-LD BreadcrumbList).
    Image = first /wp-content/uploads/ photo that is not a logo / material-icon.
  * Pages are UTF-8.  Values use HTML entity &frasl; (fraction slash).

Four user-selectable pipelines (mirrors the ARM/Tsunoda crawler):
  1) Model code  - paste codes (#102 / 102) -> crawl catalog & filter by Product #
  2) Slug / URL  - paste product slugs or full /product/ URLs (all models on page)
  3) Crawl ALL   - every product in the sitemap, extract all models
  4) Excel column- pick file -> sheet -> column of codes / slugs / URLs
"""

import os
import re
import queue
import asyncio
import threading
from urllib.parse import urljoin, urlparse
from tkinter import *
from tkinter import ttk, filedialog, messagebox

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser

BASE = "https://www.ponyjorgensen.com"
HOST = "www.ponyjorgensen.com"
PRODUCT_SITEMAP = f"{BASE}/product-sitemap.xml"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

ITEM_LABEL = "Product #"
SKIP_IMG = ("logo", "/material-icons/", "icon", "placeholder", "og-")

# ---- anti-loop / safety limits ----
MAX_REDIRECTS = 5
CIRCUIT_BREAK = 25          # abort after N consecutive page failures (IP block)
MAX_ITEMS = 5000            # runaway guard on work-list size


# ----------------------------------------------------------------------------
# Core helpers (pure, no GUI)
# ----------------------------------------------------------------------------
def norm_code(s: str) -> str:
    """Normalize a Product # for matching: drop spaces + leading '#', upper."""
    return re.sub(r"[\s#]+", "", str(s)).upper()


def slug_from_url(url: str) -> str:
    m = re.search(r"/product/([^/]+)/?", url)
    return m.group(1) if m else url.rstrip("/").rsplit("/", 1)[-1]


def build_product_url(entry: str) -> str:
    """Accept a slug, a /product/ path or a full URL -> canonical product URL."""
    e = entry.strip()
    if e.startswith("http"):
        return e.rstrip("/") + "/"
    if e.startswith("/"):
        return BASE + "/" + e.strip("/") + "/"
    return f"{BASE}/product/{e.strip('/')}/"


def _clean(s: str) -> str:
    return re.sub(r"\s+", " ", s.replace("⁄", "/")).strip()


def _pick_image(tree, page_url: str) -> str:
    """First /uploads/ product photo that is not a logo / material icon."""
    for img in tree.css("img"):
        src = img.attributes.get("src") or img.attributes.get("data-src") or ""
        low = src.lower()
        if "/uploads/" not in low or any(s in low for s in SKIP_IMG):
            continue
        return urljoin(page_url, src)
    return ""


def _breadcrumb_category(tree) -> str:
    """Last breadcrumb item (JSON-LD BreadcrumbList), excluding 'Home'."""
    import json
    for sc in tree.css('script[type="application/ld+json"]'):
        try:
            data = json.loads(sc.text())
        except Exception:
            continue
        graph = data.get("@graph", [data]) if isinstance(data, dict) else data
        for g in (graph if isinstance(graph, list) else [graph]):
            if isinstance(g, dict) and g.get("@type") == "BreadcrumbList":
                names = [i.get("name", "") for i in g.get("itemListElement", [])]
                names = [n for n in names if n and n.lower() != "home"]
                if names:
                    return names[-1]
    return ""


def _page_name_brand(tree) -> tuple:
    h1 = tree.css_first("h1")
    if not h1:
        return "", ""
    name = _clean(h1.text(separator=" ", strip=True))
    brand_node = h1.css_first("span.brand") or tree.css_first("span.brand")
    brand = _clean(brand_node.text(strip=True)) if brand_node else ""
    return name, brand


# ---- spec-table grid parsing (handles rowspan/colspan, works on simple too) --
def _cell_text(node) -> str:
    return _clean(node.text(separator=" ", strip=True))


def build_grid(table) -> list:
    grid = []
    carry = {}
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
    if not cells or not cells[0].strip():
        return False
    first = cells[0].strip()
    # a model code: starts with '#' or contains a digit
    return first.startswith("#") or bool(re.search(r"\d", first))


def parse_spec_table(grid: list) -> list:
    if len(grid) < 2:
        return []
    first_data = next((i for i, r in enumerate(grid) if _is_data_row(r)), None)
    if not first_data:
        header, data = grid[:1], grid[1:]
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
        labels.append(" ".join(parts) or (ITEM_LABEL if j == 0 else f"Col{j+1}"))
    labels[0] = ITEM_LABEL

    out = []
    for r in data:
        item = r[0].strip()
        if not item:
            continue
        specs = {}
        for j, val in enumerate(r):
            if j == 0:
                continue
            v = val.strip()
            if v:
                specs.setdefault(labels[j], v)
        out.append((item, specs))
    return out


def parse_page(html: str, entry: str, url: str) -> list:
    tree = LexborHTMLParser(html)
    name, brand = _page_name_brand(tree)
    category = _breadcrumb_category(tree)
    img = _pick_image(tree, url)

    order, models = [], {}
    for table in tree.css("table"):
        grid = build_grid(table)
        rows = parse_spec_table(grid)
        # only accept tables that look like a product-spec table
        if not any(k[0].startswith("#") or re.match(r"^[\w./-]+$", k[0]) for k in rows):
            continue
        for item, specs in rows:
            key = norm_code(item)
            if not key:
                continue
            if key not in models:
                models[key] = {"Input": entry, "Name": name, "Brand": brand,
                               "Category": category, ITEM_LABEL: item,
                               "Image": img, "URL": url, "Status": "Success"}
                order.append(key)
            for k, v in specs.items():
                models[key].setdefault(k, v)

    if not models:
        return [{"Input": entry, "Name": name, "Brand": brand, "Category": category,
                 ITEM_LABEL: "", "Image": img, "URL": url, "Status": "No models"}]
    return [models[k] for k in order]


# ----------------------------------------------------------------------------
# Async crawl engine  (usable head-less; GUI wraps it)
# ----------------------------------------------------------------------------
class PonyEngine:
    def __init__(self, log=None, prog=None, stop_evt=None, pause_evt=None):
        self.log = log or (lambda s: None)
        self.prog = prog or (lambda d, t, m: None)
        self.stop_evt = stop_evt or threading.Event()
        self.pause_evt = pause_evt or threading.Event()
        self.results = []

    def _client(self, conc):
        limits = httpx.Limits(max_connections=conc, max_keepalive_connections=conc)
        transport = httpx.AsyncHTTPTransport(retries=2)
        return httpx.AsyncClient(headers={"User-Agent": UA}, timeout=25.0,
                                 limits=limits, transport=transport,
                                 follow_redirects=True, max_redirects=MAX_REDIRECTS)

    async def _get(self, client, url):
        r = await client.get(url)
        final = str(r.url)
        if r.status_code != 200:
            raise httpx.HTTPStatusError(f"HTTP {r.status_code}",
                                        request=r.request, response=r)
        return r.content.decode("utf-8", "replace"), final

    async def load_sitemap(self, client):
        self.log("[SYS] Tai product-sitemap.xml...\n")
        try:
            html, _ = await self._get(client, PRODUCT_SITEMAP)
        except Exception as e:
            self.log(f"[ERR] sitemap: {e}\n")
            return []
        urls = re.findall(r"<loc>\s*(https?://[^<\s]+/product/[^<\s]+?)\s*</loc>", html)
        seen, clean = set(), []
        for u in urls:
            u = u.rstrip("/") + "/"
            if u not in seen:
                seen.add(u); clean.append(u)
        self.log(f"[SYS] Sitemap: {len(clean)} san pham.\n")
        return clean

    async def _process(self, client, url, entry=None):
        entry = entry or url
        try:
            html, final = await self._get(client, url)
            if "/product/" not in final:
                return [{"Input": entry, "URL": final, "Status": "Redirect off /product/"}], False
            recs = parse_page(html, entry, final)
            ok = any(r.get("Status") == "Success" for r in recs)
            return recs, ok
        except httpx.TooManyRedirects:
            return [{"Input": entry, "URL": url, "Status": "Redirect loop"}], False
        except Exception as e:
            return [{"Input": entry, "URL": url, "Status": f"Error: {type(e).__name__}"}], False

    async def run(self, pipeline, items, conc, delay):
        self.results = []
        async with self._client(conc) as client:
            if pipeline == "codes":
                work = [build_product_url(e) for e in items]
                await self._crawl(client, work, conc, delay, filter_set=None, entries=items)
            elif pipeline == "all":
                work = await self.load_sitemap(client)
                await self._crawl(client, work, conc, delay, filter_set=None)
            elif pipeline == "model":
                match = {}
                for c in items:
                    match.setdefault(norm_code(c), c)
                self.log(f"[SYS] Tim {len(match)} ma model -> quet toan catalog.\n")
                work = await self.load_sitemap(client)
                await self._crawl(client, work, conc, delay, filter_set=match)
                found = {norm_code(r.get(ITEM_LABEL, "")) for r in self.results
                         if r.get("Status") == "Success"}
                for k, orig in match.items():
                    if k not in found:
                        self.results.append({"Input": orig, ITEM_LABEL: orig, "Status": "Not Found"})
                        self.log(f"[MISS] {orig} khong thay tren site.\n")
                self.log(f"[SYS] Khop {len(found & set(match))}/{len(match)} ma.\n")
        return self.results

    async def _crawl(self, client, work, conc, delay, filter_set, entries=None):
        # dedup keep order
        seen, uniq, uentries = set(), [], []
        for i, u in enumerate(work):
            if u not in seen:
                seen.add(u); uniq.append(u)
                uentries.append(entries[i] if entries else None)
        work, entries = uniq, uentries

        total = len(work)
        if not total:
            self.log("[WARN] Khong co trang nao de crawl.\n")
            return
        if total > MAX_ITEMS:
            self.log(f"[WARN] {total} trang vuot {MAX_ITEMS}, cat bot.\n")
            work = work[:MAX_ITEMS]; entries = entries[:MAX_ITEMS]; total = MAX_ITEMS
        self.log(f"[SYS] Crawl {total} trang | concurrency={conc}\n")

        sem = asyncio.Semaphore(conc)
        done = models = consec_fail = 0
        lock = asyncio.Lock()

        async def one(url, entry):
            nonlocal done, models, consec_fail
            while self.pause_evt.is_set() and not self.stop_evt.is_set():
                await asyncio.sleep(0.2)
            if self.stop_evt.is_set():
                return
            async with sem:
                if self.stop_evt.is_set():
                    return
                recs, ok = await self._process(client, url, entry or url)
            if filter_set is not None:
                recs = [r for r in recs if norm_code(r.get(ITEM_LABEL, "")) in filter_set]
            async with lock:
                if recs:
                    self.results.extend(recs)
                done += 1
                n = sum(1 for r in recs if r.get("Status") == "Success")
                models += n
                consec_fail = 0 if ok else consec_fail + 1
                if filter_set is None:
                    tag = f"OK  {n} model" if ok else (recs[0].get("Status", "MISS") if recs else "MISS")
                    self.log(f"[{tag}] {slug_from_url(url)}\n")
                elif n:
                    self.log(f"[HIT {n}] {slug_from_url(url)}\n")
                self.prog(done, total, models)
                if consec_fail >= CIRCUIT_BREAK and not self.stop_evt.is_set():
                    self.stop_evt.set()
                    self.log(f"[STOP] {CIRCUIT_BREAK} loi lien tiep -> ngat an toan.\n")
            if delay:
                await asyncio.sleep(delay)

        await asyncio.gather(*(one(u, e) for u, e in zip(work, entries)))


# ----------------------------------------------------------------------------
# Excel export
# ----------------------------------------------------------------------------
def export_excel(results, path):
    lead = ["Input", "Name", "Brand", "Category", ITEM_LABEL, "Status"]
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
# GUI application  (same skeleton as the ARM / Tsunoda crawlers)
# ----------------------------------------------------------------------------
class PonyCrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pony / Jorgensen Crawler  |  ponyjorgensen.com")
        self.root.geometry("980x700")
        self.root.minsize(880, 600)

        self.results = []
        self.log_q = queue.Queue()
        self.stop_evt = threading.Event()
        self.pause_evt = threading.Event()
        self.file_rows = []
        self.is_running = False

        self.bg = "#f4f5f7"
        self.primary = "#c8102e"       # Pony red
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
        ttk.Label(header, text="PONY / JORGENSEN CRAWLER", style="Title.TLabel").pack(side=LEFT)
        ttk.Label(header, text="  httpx + selectolax  •  sitemap + table-rows",
                  font=("Segoe UI", 9, "italic"), foreground="#666").pack(side=LEFT, pady=5)

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=BOTH, expand=YES)

        left = ttk.LabelFrame(main, text=" Pipeline & Input ", padding=12, width=360)
        left.pack(side=LEFT, fill=Y, padx=(0, 10))
        left.pack_propagate(False)

        self.pipeline = StringVar(value="model")
        ttk.Label(left, text="Chon pipeline:", font=("Segoe UI Semibold", 9)).pack(anchor=W)
        for val, txt in (("model", "1. Theo MA MODEL (Product #)"),
                         ("codes", "2. Theo slug / URL san pham"),
                         ("all",   "3. Crawl ALL (toan bo sitemap)"),
                         ("excel", "4. Theo cot trong Excel")):
            ttk.Radiobutton(left, text=txt, value=val, variable=self.pipeline,
                            command=self._on_pipeline).pack(anchor=W, pady=1)

        self.model_frame = ttk.Frame(left)
        ttk.Label(self.model_frame, text="Ma model (moi dong 1 ma):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.model_text = Text(self.model_frame, height=10, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.model_text.pack(fill=BOTH, expand=YES)
        ttk.Label(self.model_frame, text="vd: #102  103  #3712-HT\n(quet toan catalog roi loc dung ma)",
                  font=("Segoe UI", 8, "italic"), foreground="#888").pack(anchor=W)

        self.codes_frame = ttk.Frame(left)
        ttk.Label(self.codes_frame, text="Slug hoac URL (moi dong 1 muc):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.codes_text = Text(self.codes_frame, height=10, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.codes_text.pack(fill=BOTH, expand=YES)
        ttk.Label(self.codes_frame, text="vd: pony-carriage-c-clamp\n/product/pony-press-screws/",
                  font=("Segoe UI", 8, "italic"), foreground="#888").pack(anchor=W)

        self.excel_frame = ttk.Frame(left)
        top = ttk.Frame(self.excel_frame); top.pack(fill=X, pady=(8, 4))
        ttk.Label(top, text="File Excel:", font=("Segoe UI Semibold", 9)).pack(side=LEFT)
        ttk.Button(top, text="Chon file...", width=12, command=self._browse_excel).pack(side=RIGHT)
        self.excel_cfg = ttk.Frame(self.excel_frame); self.excel_cfg.pack(fill=X)
        ttk.Label(self.excel_cfg, text="Chua chon file",
                  font=("Segoe UI", 9, "italic"), foreground="#888").pack(pady=8)
        self.excel_as_model = BooleanVar(value=True)
        Checkbutton(self.excel_frame, text="Cot la MA MODEL (Product #)",
                    variable=self.excel_as_model, bg=self.bg, fg=self.text,
                    font=("Segoe UI", 9), activebackground=self.bg,
                    anchor=W).pack(fill=X, pady=(4, 0))

        self.all_frame = ttk.Frame(left)
        ttk.Label(self.all_frame,
                  text="Tai toan bo product-sitemap.xml\nroi boc moi model trong bang.",
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
        self._log("[SYS] San sang crawl ponyjorgensen.com.\n")

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
        pick = next((c for c in cols if re.search(r"product|ma|code|item|sku|slug|url", str(c), re.I)),
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
            n = sum(1 for r in self.results if r.get("Status") == "Success")
            messagebox.showinfo("Xong", f"Crawl xong {n} model. Bam EXPORT EXCEL.")

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

    def _gather(self, widget):
        raw = widget.get("1.0", "end-1c")
        return [x.strip() for x in re.split(r"[\r\n]+", raw) if x.strip()]

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
        eng = PonyEngine(log=lambda s: self.log_q.put(("log", s)),
                         prog=lambda d, t, m: self.log_q.put(("prog", (d, t, m))),
                         stop_evt=self.stop_evt, pause_evt=self.pause_evt)
        if pipeline == "model":
            items = self._gather(self.model_text)
        elif pipeline == "codes":
            items = self._gather(self.codes_text)
        elif pipeline == "excel":
            items = self._gather_excel()
            pipeline = "model" if self.excel_as_model.get() else "codes"
        else:
            items = []

        if pipeline in ("model", "codes") and not items:
            self.log_q.put(("log", "[WARN] Chua nhap gi.\n"))
            self.log_q.put(("done", True)); return

        await eng.run(pipeline, items, conc, delay)
        self.results = eng.results
        self.log_q.put(("done", self.stop_evt.is_set()))

    def _export(self):
        if not self.results:
            messagebox.showwarning("Trong", "Chua co du lieu."); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")],
                                            initialfile="ponyjorgensen_data.xlsx")
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
    PonyCrawlerApp(root)
    root.mainloop()
