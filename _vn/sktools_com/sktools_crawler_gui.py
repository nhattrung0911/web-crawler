"""
SK Tools Crawler  (sktools.com)
===============================
Modern async crawler: httpx (pooled HTTP) + the Shopify public JSON API.
Same engine family as the ARM / Anex / Tsunoda / Pony / EIGHT crawlers.

Site recon (sktools.com):
  * Platform = SHOPIFY (header "powered-by: Shopify", cdn.shopify.com).
  * API-first: the public storefront JSON needs NO auth:
        /products.json?limit=250&page=N    -> paginate the whole catalog
        /products/{handle}.json             -> one product by handle
    ~1878 products across 8 pages of 250.  This is the fastest possible path —
    one JSON pull per page, no HTML scraping.
  * Each product carries: title, handle, body_html (description), vendor,
    product_type, tags, options, images[], variants[] (id, sku, price,
    compare_at_price, available, option1..3, featured_image).
    The SKU lives in variant.sku (e.g. "SK03449"); most products have a single
    "Default Title" variant, some sets have several -> emit ONE ROW PER VARIANT.
  * No custom auth / no rate problems at this scale; still keep the safety net.

Four user-selectable pipelines:
  1) Model / SKU  - paste SKUs -> paginate catalog & keep matching variants
  2) Handle / URL - paste product handles or full /products/ URLs (direct JSON)
  3) Crawl ALL    - every product/variant in the catalog
  4) Excel column - pick file -> sheet -> SKU / handle / URL column
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

BASE = "https://sktools.com"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

PAGE_SIZE = 250

# ---- anti-loop / safety limits ----
MAX_PAGES = 200            # runaway guard on pagination
MAX_ITEMS = 20000          # runaway guard on work-list size
CIRCUIT_BREAK = 25


# ----------------------------------------------------------------------------
# Core helpers (pure, no GUI)
# ----------------------------------------------------------------------------
def norm_code(s: str) -> str:
    """Normalize a SKU for matching: drop spaces + leading '#', upper-case."""
    return re.sub(r"[\s#]+", "", str(s)).upper()


def handle_from_entry(entry: str) -> str:
    e = entry.strip()
    if e.startswith("http"):
        m = re.search(r"/products/([^/?#]+)", e)
        return m.group(1) if m else e
    return e.strip("/").split("/")[-1]


def _clean_html(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r"<[^>]+>", " ", s)
    s = (s.replace("&nbsp;", " ").replace("&amp;", "&")
           .replace("&lt;", "<").replace("&gt;", ">").replace("&#39;", "'"))
    return re.sub(r"\s+", " ", s).strip()


def _img(u: str) -> str:
    if not u:
        return ""
    return ("https:" + u) if u.startswith("//") else u


def product_to_rows(p: dict, entry: str = "") -> list:
    """One row per variant, with product-level fields attached."""
    handle = p.get("handle", "")
    url = f"{BASE}/products/{handle}"
    tags = p.get("tags", [])
    tags = ", ".join(tags) if isinstance(tags, list) else str(tags or "")
    desc = _clean_html(p.get("body_html", ""))
    main_img = _img((p.get("images") or [{}])[0].get("src", "")) if p.get("images") else ""
    opt_names = [o.get("name") for o in (p.get("options") or [])]

    rows = []
    for v in (p.get("variants") or [{}]):
        vimg = ""
        fi = v.get("featured_image")
        if isinstance(fi, dict):
            vimg = _img(fi.get("src", ""))
        opts = [v.get("option1"), v.get("option2"), v.get("option3")]
        opts = [o for o in opts if o and o != "Default Title"]
        rows.append({
            "Input": entry or (v.get("sku") or handle),
            "Title": p.get("title", ""),
            "SKU": v.get("sku", "") or "",
            "Variant": " / ".join(opts),
            "Price": v.get("price", ""),
            "Compare Price": v.get("compare_at_price", "") or "",
            "Available": "Yes" if v.get("available") else "No",
            "Vendor": p.get("vendor", ""),
            "Type": p.get("product_type", ""),
            "Tags": tags,
            "Options": ", ".join(n for n in opt_names if n and n != "Title"),
            "Description": desc,
            "Image": vimg or main_img,
            "Handle": handle,
            "URL": url,
            "Status": "Success",
        })
    return rows


# ----------------------------------------------------------------------------
# Async crawl engine  (usable head-less; GUI wraps it)
# ----------------------------------------------------------------------------
class SKEngine:
    def __init__(self, log=None, prog=None, stop_evt=None, pause_evt=None):
        self.log = log or (lambda s: None)
        self.prog = prog or (lambda d, t, m: None)
        self.stop_evt = stop_evt or threading.Event()
        self.pause_evt = pause_evt or threading.Event()
        self.results = []

    def _client(self, conc):
        limits = httpx.Limits(max_connections=conc, max_keepalive_connections=conc)
        transport = httpx.AsyncHTTPTransport(retries=2)
        return httpx.AsyncClient(headers={"User-Agent": UA,
                                          "Accept": "application/json"},
                                 timeout=30.0, limits=limits, transport=transport,
                                 follow_redirects=True, max_redirects=5)

    async def _get_json(self, client, url):
        for attempt in range(5):
            if self.stop_evt.is_set():
                raise asyncio.CancelledError()
            r = await client.get(url)
            if r.status_code == 429:
                ra = r.headers.get("retry-after")
                wait = float(ra) if (ra and ra.replace(".", "").isdigit()) else min(2 ** attempt, 20)
                self.log(f"[429] cho {wait:.0f}s roi thu lai\n")
                await asyncio.sleep(wait)
                continue
            if r.status_code != 200:
                raise httpx.HTTPStatusError(f"HTTP {r.status_code}",
                                            request=r.request, response=r)
            return r.json()
        raise httpx.HTTPStatusError("HTTP 429 (het luot)", request=None, response=r)

    async def fetch_all(self, client):
        """Paginate /products.json until an empty page."""
        self.log("[SYS] Tai catalog qua products.json...\n")
        products = []
        page = 1
        while page <= MAX_PAGES and not self.stop_evt.is_set():
            while self.pause_evt.is_set() and not self.stop_evt.is_set():
                await asyncio.sleep(0.2)
            url = f"{BASE}/products.json?limit={PAGE_SIZE}&page={page}"
            try:
                data = await self._get_json(client, url)
            except asyncio.CancelledError:
                break
            except Exception as e:
                self.log(f"[ERR] page {page}: {e}\n")
                break
            batch = data.get("products", [])
            if not batch:
                break
            products.extend(batch)
            self.log(f"[SYS] page {page}: +{len(batch)} (tong {len(products)})\n")
            self.prog(len(products), len(products) + PAGE_SIZE, len(products))
            if len(batch) < PAGE_SIZE:
                break
            page += 1
            if len(products) > MAX_ITEMS:
                self.log(f"[WARN] vuot {MAX_ITEMS}, dung.\n")
                break
        self.log(f"[SYS] Catalog: {len(products)} san pham.\n")
        return products

    async def fetch_handle(self, client, entry):
        handle = handle_from_entry(entry)
        url = f"{BASE}/products/{handle}.json"
        try:
            data = await self._get_json(client, url)
        except httpx.HTTPStatusError as e:
            code = e.response.status_code if e.response is not None else "?"
            return [{"Input": entry, "Status": f"HTTP {code}"}]
        except Exception as e:
            return [{"Input": entry, "Status": f"Error: {type(e).__name__}"}]
        prod = data.get("product")
        if not prod:
            return [{"Input": entry, "Status": "Not Found"}]
        return product_to_rows(prod, entry)

    async def run(self, pipeline, items, conc, delay):
        self.results = []
        async with self._client(conc) as client:
            if pipeline == "all":
                prods = await self.fetch_all(client)
                for p in prods:
                    self.results.extend(product_to_rows(p))
                self.log(f"[SYS] Tong {len(self.results)} variant.\n")

            elif pipeline == "model":
                match = {}
                for c in items:
                    match.setdefault(norm_code(c), c)
                self.log(f"[SYS] Tim {len(match)} SKU -> quet catalog.\n")
                prods = await self.fetch_all(client)
                found = set()
                for p in prods:
                    for row in product_to_rows(p):
                        nc = norm_code(row.get("SKU", ""))
                        if nc in match:
                            row["Input"] = match[nc]
                            self.results.append(row)
                            found.add(nc)
                for nc, orig in match.items():
                    if nc not in found:
                        self.results.append({"Input": orig, "SKU": orig, "Status": "Not Found"})
                        self.log(f"[MISS] {orig} khong thay.\n")
                self.log(f"[SYS] Khop {len(found)}/{len(match)} SKU.\n")

            elif pipeline == "codes":
                await self._direct(client, items, conc, delay)
        return self.results

    async def _direct(self, client, items, conc, delay):
        seen, work = set(), []
        for e in items:
            h = handle_from_entry(e)
            if h not in seen:
                seen.add(h); work.append(e)
        total = len(work)
        if not total:
            self.log("[WARN] Khong co gi.\n"); return
        self.log(f"[SYS] Direct {total} handle | concurrency={conc}\n")
        sem = asyncio.Semaphore(conc)
        done = ok_n = consec_fail = 0
        lock = asyncio.Lock()

        async def one(entry):
            nonlocal done, ok_n, consec_fail
            while self.pause_evt.is_set() and not self.stop_evt.is_set():
                await asyncio.sleep(0.2)
            if self.stop_evt.is_set():
                return
            async with sem:
                if self.stop_evt.is_set():
                    return
                rows = await self.fetch_handle(client, entry)
            ok = any(r.get("Status") == "Success" for r in rows)
            async with lock:
                self.results.extend(rows)
                done += 1
                n = sum(1 for r in rows if r.get("Status") == "Success")
                ok_n += n
                consec_fail = 0 if ok else consec_fail + 1
                tag = f"OK  {n} variant" if ok else (rows[0].get("Status", "MISS") if rows else "MISS")
                self.log(f"[{tag}] {handle_from_entry(entry)}\n")
                self.prog(done, total, ok_n)
                if consec_fail >= CIRCUIT_BREAK and not self.stop_evt.is_set():
                    self.stop_evt.set()
                    self.log(f"[STOP] {CIRCUIT_BREAK} loi lien tiep -> ngat.\n")
            if delay:
                await asyncio.sleep(delay)

        await asyncio.gather(*(one(e) for e in work))


# ----------------------------------------------------------------------------
# Excel export
# ----------------------------------------------------------------------------
def export_excel(results, path):
    cols = ["Input", "Title", "SKU", "Variant", "Price", "Compare Price",
            "Available", "Vendor", "Type", "Tags", "Options", "Description",
            "Image", "Handle", "URL", "Status"]
    extra = []
    for r in results:
        for k in r:
            if k not in cols and k not in extra:
                extra.append(k)
    df = pd.DataFrame(results).reindex(columns=cols + extra)
    df.to_excel(path, index=False)
    return df


# ----------------------------------------------------------------------------
# GUI application  (same skeleton as the ARM / EIGHT crawlers)
# ----------------------------------------------------------------------------
class SKCrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SK Tools Crawler  |  sktools.com")
        self.root.geometry("980x700")
        self.root.minsize(880, 600)

        self.results = []
        self.log_q = queue.Queue()
        self.stop_evt = threading.Event()
        self.pause_evt = threading.Event()
        self.file_rows = []
        self.is_running = False

        self.bg = "#f4f5f7"
        self.primary = "#c8102e"       # SK red
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
        ttk.Label(header, text="SK TOOLS CRAWLER", style="Title.TLabel").pack(side=LEFT)
        ttk.Label(header, text="  httpx + Shopify products.json  •  API-first",
                  font=("Segoe UI", 9, "italic"), foreground="#666").pack(side=LEFT, pady=5)

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=BOTH, expand=YES)

        left = ttk.LabelFrame(main, text=" Pipeline & Input ", padding=12, width=360)
        left.pack(side=LEFT, fill=Y, padx=(0, 10))
        left.pack_propagate(False)

        self.pipeline = StringVar(value="model")
        ttk.Label(left, text="Chon pipeline:", font=("Segoe UI Semibold", 9)).pack(anchor=W)
        for val, txt in (("model", "1. Theo MA / SKU"),
                         ("codes", "2. Theo handle / URL san pham"),
                         ("all",   "3. Crawl ALL (toan bo catalog)"),
                         ("excel", "4. Theo cot trong Excel")):
            ttk.Radiobutton(left, text=txt, value=val, variable=self.pipeline,
                            command=self._on_pipeline).pack(anchor=W, pady=1)

        self.model_frame = ttk.Frame(left)
        ttk.Label(self.model_frame, text="SKU (moi dong 1 ma):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.model_text = Text(self.model_frame, height=11, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.model_text.pack(fill=BOTH, expand=YES)
        ttk.Label(self.model_frame, text="vd: SK03449  SK90032\n(quet catalog roi loc dung SKU)",
                  font=("Segoe UI", 8, "italic"), foreground="#888").pack(anchor=W)

        self.codes_frame = ttk.Frame(left)
        ttk.Label(self.codes_frame, text="Handle hoac URL (moi dong 1 muc):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.codes_text = Text(self.codes_frame, height=11, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.codes_text.pack(fill=BOTH, expand=YES)
        ttk.Label(self.codes_frame, text="vd: sktools.com/products/<handle>",
                  font=("Segoe UI", 8, "italic"), foreground="#888").pack(anchor=W)

        self.excel_frame = ttk.Frame(left)
        top = ttk.Frame(self.excel_frame); top.pack(fill=X, pady=(8, 4))
        ttk.Label(top, text="File Excel:", font=("Segoe UI Semibold", 9)).pack(side=LEFT)
        ttk.Button(top, text="Chon file...", width=12, command=self._browse_excel).pack(side=RIGHT)
        self.excel_cfg = ttk.Frame(self.excel_frame); self.excel_cfg.pack(fill=X)
        ttk.Label(self.excel_cfg, text="Chua chon file",
                  font=("Segoe UI", 9, "italic"), foreground="#888").pack(pady=8)
        self.excel_as_model = BooleanVar(value=True)
        Checkbutton(self.excel_frame, text="Cot la SKU (tim theo ma)",
                    variable=self.excel_as_model, bg=self.bg, fg=self.text,
                    font=("Segoe UI", 9), activebackground=self.bg,
                    anchor=W).pack(fill=X, pady=(4, 0))

        self.all_frame = ttk.Frame(left)
        ttk.Label(self.all_frame,
                  text="Phan trang products.json (250/trang)\nroi boc moi variant.",
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
        self.stats_lbl = ttk.Label(info, text="0  |  0 variant", foreground="#555")
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
        self._log("[SYS] San sang crawl sktools.com (Shopify).\n")

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
        pick = next((c for c in cols if re.search(r"sku|ma|code|item|handle|url", str(c), re.I)),
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
                self.pbar["maximum"] = max(total, 1); self.pbar["value"] = done
                pct = int(done / total * 100) if total else 0
                self.prog_lbl.config(text=f"Crawling {pct}%")
                self.stats_lbl.config(text=f"{done}  |  {models} variant")
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
            messagebox.showinfo("Xong", f"Crawl xong {n} variant. Bam EXPORT EXCEL.")

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
        eng = SKEngine(log=lambda s: self.log_q.put(("log", s)),
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
                                            initialfile="sktools_data.xlsx")
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
    SKCrawlerApp(root)
    root.mainloop()
