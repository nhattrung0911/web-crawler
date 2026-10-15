"""
EIGHT Tool Crawler  (eight-tool.co.jp)
======================================
Modern async crawler: httpx (pooled HTTP) + embedded-JSON extraction.
Same engine family as the ARM / Anex / Tsunoda / Pony crawlers.

Site recon (eight-tool.co.jp):
  * Platform = WIX (server "Pepyaka", parastorage/wixstatic assets) with the
    Wix Stores app.  NOT WordPress/Shopify.
  * API-first: the product data ships as EMBEDDED JSON in every product page —
    the Wix "warmupData" block:
        <script type="application/json"> ... "appsWarmupData":
            {<storesAppId>: {"productPage_JPY_<slug>": {"catalog":
                {"product": { id, name, sku, price, formattedPrice,
                              description, media[], categories, inventory ... }}}}}
    -> parse that JSON, do NOT scrape the obfuscated Wix DOM.
  * Discovery = the Wix product sitemap (complete, modern):
        https://www.eight-tool.co.jp/store-products-sitemap.xml   (~256 products)
    Product URL:  /product-page/{slug}   and  slug == the product code, lower-cased
        TLS-9N -> /product-page/tls-9n ,  004-6 -> /product-page/004-6
    -> O(1) direct lookup by code (Anex-style).
  * EIGHT fills the store minimally: `name` == the product code, `price` is
    sometimes set (¥8,300) / often ¥0, `sku`/`description`/options usually empty.
    The spec sheet is baked into the product IMAGE (media[].url on wixstatic).
    So the reliably-extracted fields are: Code, Price, In-stock, Image(s),
    Category (when present).
  * media[].url = "f1ddc6_xxxx~mv2.png"  -> full-res =
        https://static.wixstatic.com/media/{url}   (drop the /v1/fit/... transform)

Three user-selectable pipelines:
  1) Code / URL   - paste product codes or full /product-page/ URLs (direct)
  2) Crawl ALL    - pull every product from the store sitemap
  3) Excel column - pick file -> sheet -> code / URL column
"""

import os
import re
import json
import queue
import asyncio
import threading
from tkinter import *
from tkinter import ttk, filedialog, messagebox

import httpx
import pandas as pd

BASE = "https://www.eight-tool.co.jp"
STORE_SITEMAP = f"{BASE}/store-products-sitemap.xml"
WIXMEDIA = "https://static.wixstatic.com/media/"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# ---- anti-loop / safety limits ----
MAX_REDIRECTS = 5
CIRCUIT_BREAK = 25
MAX_ITEMS = 5000


# ----------------------------------------------------------------------------
# Core helpers (pure, no GUI)
# ----------------------------------------------------------------------------
def norm_code(s: str) -> str:
    return re.sub(r"\s+", "", str(s)).upper()


def code_to_slug(code: str) -> str:
    s = code.strip().lower()
    s = re.sub(r"\s+", "-", s)
    return s


def build_url(entry: str) -> str:
    e = entry.strip()
    if e.startswith("http"):
        return e.split("?")[0].rstrip("/")
    if e.startswith("/"):
        return BASE + "/" + e.strip("/")
    return f"{BASE}/product-page/{code_to_slug(e)}"


def slug_from_url(url: str) -> str:
    m = re.search(r"/product-page/([^/?#]+)", url)
    return m.group(1) if m else url.rstrip("/").rsplit("/", 1)[-1]


def _clean_html(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r"<[^>]+>", " ", s)
    s = s.replace("&nbsp;", " ").replace("&amp;", "&")
    return re.sub(r"\s+", " ", s).strip()


def _full_img(u: str) -> str:
    if not u:
        return ""
    if u.startswith("http"):
        # strip Wix transform (/v1/fit/... /file.png) back to the original media
        m = re.search(r"static\.wixstatic\.com/media/([^/]+)", u)
        return WIXMEDIA + m.group(1) if m else u
    return WIXMEDIA + u


def find_product_json(html: str) -> dict:
    """Pull the Wix Stores product object out of the warmupData JSON blocks."""
    for blk in re.findall(r'<script[^>]*type="application/json"[^>]*>(.*?)</script>',
                          html, re.S):
        if "appsWarmupData" not in blk:
            continue
        try:
            data = json.loads(blk)
        except Exception:
            continue
        aw = data.get("appsWarmupData")
        if not isinstance(aw, dict):
            continue
        for app in aw.values():
            if not isinstance(app, dict):
                continue
            for k, v in app.items():
                if k.startswith("productPage") and isinstance(v, dict) \
                        and isinstance(v.get("catalog"), dict):
                    prod = v["catalog"].get("product")
                    if isinstance(prod, dict):
                        return prod
    return {}


def parse_product(html: str, entry: str, url: str) -> dict:
    p = find_product_json(html)
    if not p:
        return {"Input": entry, "URL": url, "Status": "No product JSON"}

    media = p.get("media") or []
    imgs = []
    for it in media:
        if not isinstance(it, dict):
            continue
        u = it.get("url") or it.get("fullUrl") or ""
        full = _full_img(u)
        if full and full not in imgs:
            imgs.append(full)

    cats = []
    for c in (p.get("categories") or []):
        if isinstance(c, dict) and c.get("name"):
            cats.append(c["name"])
    if not cats:
        for b in (p.get("breadcrumbs") or []):
            v = b.get("value") if isinstance(b, dict) else None
            if v and v.lower() not in ("home", "shop", "store"):
                cats.append(v)

    items = p.get("productItems") or [{}]
    inv = items[0].get("inventory") if isinstance(items[0], dict) else None
    in_stock = ""
    if isinstance(inv, dict):
        in_stock = "Yes" if inv.get("status") in ("in_stock", "partially_out_of_stock") \
            or inv.get("quantity") else ("No" if inv.get("status") == "out_of_stock" else "")

    return {
        "Input": entry,
        "Code": p.get("name", "") or "",
        "SKU": p.get("sku", "") or "",
        "Price": p.get("formattedPrice", "") or "",
        "In Stock": in_stock,
        "Category": " > ".join(cats),
        "Description": _clean_html(p.get("description", "")),
        "Image": imgs[0] if imgs else "",
        "All Images": " | ".join(imgs),
        "URL": url,
        "Status": "Success",
    }


# ----------------------------------------------------------------------------
# Async crawl engine  (usable head-less; GUI wraps it)
# ----------------------------------------------------------------------------
class EightEngine:
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
        """Fetch with Wix-aware 429 handling: respect Retry-After, back off, retry."""
        for attempt in range(5):
            if self.stop_evt.is_set():
                raise asyncio.CancelledError()
            r = await client.get(url)
            if r.status_code == 429:
                ra = r.headers.get("retry-after")
                wait = float(ra) if (ra and ra.replace(".", "").isdigit()) else min(2 ** attempt, 20)
                self.log(f"[429] {slug_from_url(url)} -> cho {wait:.0f}s roi thu lai\n")
                await asyncio.sleep(wait)
                continue
            final = str(r.url)
            if r.status_code != 200:
                raise httpx.HTTPStatusError(f"HTTP {r.status_code}",
                                            request=r.request, response=r)
            return r.content.decode("utf-8", "replace"), final
        raise httpx.HTTPStatusError("HTTP 429 (het luot thu)", request=None, response=r)

    async def load_sitemap(self, client):
        self.log("[SYS] Tai store-products-sitemap.xml...\n")
        try:
            html, _ = await self._get(client, STORE_SITEMAP)
        except Exception as e:
            self.log(f"[ERR] sitemap: {e}\n")
            return []
        urls = re.findall(r"<loc>\s*(https?://[^<\s]+/product-page/[^<\s]+?)\s*</loc>", html)
        seen, clean = set(), []
        for u in urls:
            u = u.split("?")[0].rstrip("/")
            if u not in seen:
                seen.add(u); clean.append(u)
        self.log(f"[SYS] Sitemap: {len(clean)} san pham.\n")
        return clean

    async def _process(self, client, url, entry):
        # Wix varnish sometimes serves a cache variant WITHOUT the warmupData
        # block -> re-fetch a couple times before giving up.
        rec = {"Input": entry, "URL": url, "Status": "No product JSON"}
        for attempt in range(3):
            # attempt>0: add a cache-buster -> forces a fresh Wix render when the
            # varnish cache is serving a broken warmupData-less variant.
            furl = url if attempt == 0 else \
                url + ("&" if "?" in url else "?") + f"nocache={attempt}"
            try:
                html, final = await self._get(client, furl)
            except httpx.TooManyRedirects:
                return {"Input": entry, "URL": url, "Status": "Redirect loop"}, False
            except httpx.HTTPStatusError as e:
                code = e.response.status_code if e.response is not None else "?"
                return {"Input": entry, "URL": url, "Status": f"HTTP {code}"}, False
            except asyncio.CancelledError:
                return {"Input": entry, "URL": url, "Status": "Stopped"}, False
            except Exception as e:
                return {"Input": entry, "URL": url, "Status": f"Error: {type(e).__name__}"}, False
            if "/product-page/" not in final:
                return {"Input": entry, "URL": final, "Status": "Redirect off store"}, False
            rec = parse_product(html, entry, url)
            if rec.get("Status") == "Success":
                return rec, True
            await asyncio.sleep(0.6 * (attempt + 1))
        return rec, False

    async def run(self, pipeline, items, conc, delay):
        self.results = []
        async with self._client(conc) as client:
            if pipeline == "all":
                work = await self.load_sitemap(client)
                entries = list(work)
            else:                                   # codes / excel-as-codes
                entries = list(items)
                work = [build_url(e) for e in items]
            await self._crawl(client, work, entries, conc, delay)
        return self.results

    async def _crawl(self, client, work, entries, conc, delay):
        seen, uw, ue = set(), [], []
        for u, e in zip(work, entries):
            if u not in seen:
                seen.add(u); uw.append(u); ue.append(e)
        work, entries = uw, ue

        total = len(work)
        if not total:
            self.log("[WARN] Khong co gi de crawl.\n")
            return
        if total > MAX_ITEMS:
            self.log(f"[WARN] {total} vuot {MAX_ITEMS}, cat bot.\n")
            work, entries, total = work[:MAX_ITEMS], entries[:MAX_ITEMS], MAX_ITEMS
        self.log(f"[SYS] Crawl {total} san pham | concurrency={conc}\n")

        sem = asyncio.Semaphore(conc)
        done = ok_n = consec_fail = 0
        lock = asyncio.Lock()

        async def one(url, entry):
            nonlocal done, ok_n, consec_fail
            while self.pause_evt.is_set() and not self.stop_evt.is_set():
                await asyncio.sleep(0.2)
            if self.stop_evt.is_set():
                return
            async with sem:
                if self.stop_evt.is_set():
                    return
                rec, ok = await self._process(client, url, entry)
            async with lock:
                self.results.append(rec)
                done += 1
                ok_n += 1 if ok else 0
                consec_fail = 0 if ok else consec_fail + 1
                tag = f"OK  {rec.get('Price','')}".strip() if ok else rec.get("Status", "MISS")
                self.log(f"[{tag}] {rec.get('Code') or slug_from_url(url)}\n")
                self.prog(done, total, ok_n)
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
    cols = ["Input", "Code", "SKU", "Price", "In Stock", "Category",
            "Description", "Image", "All Images", "URL", "Status"]
    extra = []
    for r in results:
        for k in r:
            if k not in cols and k not in extra:
                extra.append(k)
    df = pd.DataFrame(results).reindex(columns=cols + extra)
    df.to_excel(path, index=False)
    return df


# ----------------------------------------------------------------------------
# GUI application  (same skeleton as the ARM / Pony crawlers)
# ----------------------------------------------------------------------------
class EightCrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EIGHT Tool Crawler  |  eight-tool.co.jp")
        self.root.geometry("980x700")
        self.root.minsize(880, 600)

        self.results = []
        self.log_q = queue.Queue()
        self.stop_evt = threading.Event()
        self.pause_evt = threading.Event()
        self.file_rows = []
        self.is_running = False

        self.bg = "#f4f5f7"
        self.primary = "#1b4a8b"       # EIGHT blue
        self.accent = "#2f6bc0"
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
        ttk.Label(header, text="EIGHT TOOL CRAWLER", style="Title.TLabel").pack(side=LEFT)
        ttk.Label(header, text="  httpx + Wix warmupData JSON  •  sitemap + slug-direct",
                  font=("Segoe UI", 9, "italic"), foreground="#666").pack(side=LEFT, pady=5)

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=BOTH, expand=YES)

        left = ttk.LabelFrame(main, text=" Pipeline & Input ", padding=12, width=360)
        left.pack(side=LEFT, fill=Y, padx=(0, 10))
        left.pack_propagate(False)

        self.pipeline = StringVar(value="codes")
        ttk.Label(left, text="Chon pipeline:", font=("Segoe UI Semibold", 9)).pack(anchor=W)
        for val, txt in (("codes", "1. Theo MA / URL (dan ben duoi)"),
                         ("all",   "2. Crawl ALL (toan bo sitemap)"),
                         ("excel", "3. Theo cot trong Excel")):
            ttk.Radiobutton(left, text=txt, value=val, variable=self.pipeline,
                            command=self._on_pipeline).pack(anchor=W, pady=1)

        self.codes_frame = ttk.Frame(left)
        ttk.Label(self.codes_frame, text="Ma hoac URL (moi dong 1 muc):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.codes_text = Text(self.codes_frame, height=12, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.codes_text.pack(fill=BOTH, expand=YES)
        ttk.Label(self.codes_frame, text="vd: TLS-9N   004-6\n/product-page/tls-9n",
                  font=("Segoe UI", 8, "italic"), foreground="#888").pack(anchor=W)

        self.excel_frame = ttk.Frame(left)
        top = ttk.Frame(self.excel_frame); top.pack(fill=X, pady=(8, 4))
        ttk.Label(top, text="File Excel:", font=("Segoe UI Semibold", 9)).pack(side=LEFT)
        ttk.Button(top, text="Chon file...", width=12, command=self._browse_excel).pack(side=RIGHT)
        self.excel_cfg = ttk.Frame(self.excel_frame); self.excel_cfg.pack(fill=X)
        ttk.Label(self.excel_cfg, text="Chua chon file",
                  font=("Segoe UI", 9, "italic"), foreground="#888").pack(pady=8)

        self.all_frame = ttk.Frame(left)
        ttk.Label(self.all_frame,
                  text="Tai toan bo store-products-sitemap.xml\nroi boc JSON tung san pham.",
                  font=("Segoe UI", 9), foreground="#555").pack(anchor=W, pady=10)

        speed = ttk.Frame(left); speed.pack(fill=X, pady=(12, 8), side=BOTTOM)
        ttk.Label(speed, text="Concurrency:", font=("Segoe UI Semibold", 9)).grid(row=0, column=0, sticky=W)
        self.conc = Spinbox(speed, from_=1, to=32, width=5, justify=CENTER)
        self.conc.delete(0, END); self.conc.insert(0, "5")
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
        self._log("[SYS] San sang crawl eight-tool.co.jp.\n")

    def _on_pipeline(self):
        for f in (self.codes_frame, self.excel_frame, self.all_frame):
            f.pack_forget()
        p = self.pipeline.get()
        if p == "codes":
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
        pick = next((c for c in cols if re.search(r"ma|code|item|sku|slug|url|品番|型式", str(c), re.I)),
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
                done, total, ok = data
                self.pbar["maximum"] = total; self.pbar["value"] = done
                pct = int(done / total * 100) if total else 0
                self.prog_lbl.config(text=f"Crawling {pct}%")
                self.stats_lbl.config(text=f"{done} / {total}  |  {ok} OK")
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
            messagebox.showinfo("Xong", f"Crawl xong {n} san pham. Bam EXPORT EXCEL.")

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
        eng = EightEngine(log=lambda s: self.log_q.put(("log", s)),
                          prog=lambda d, t, m: self.log_q.put(("prog", (d, t, m))),
                          stop_evt=self.stop_evt, pause_evt=self.pause_evt)
        if pipeline == "codes":
            items = self._gather(self.codes_text)
        elif pipeline == "excel":
            items = self._gather_excel()
            pipeline = "codes"
        else:
            items = []

        if pipeline == "codes" and not items:
            self.log_q.put(("log", "[WARN] Chua nhap ma nao.\n"))
            self.log_q.put(("done", True)); return

        await eng.run(pipeline, items, conc, delay)
        self.results = eng.results
        self.log_q.put(("done", self.stop_evt.is_set()))

    def _export(self):
        if not self.results:
            messagebox.showwarning("Trong", "Chua co du lieu."); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")],
                                            initialfile="eighttool_data.xlsx")
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
    EightCrawlerApp(root)
    root.mainloop()
