"""
ARM Sangyo Product Crawler  (armsangyo.co.jp)
=============================================
Modern async crawler: httpx (pooled HTTP) + selectolax (Lexbor parser).

Site recon (armsangyo.co.jp):
  * Static HTML site on nginx. NO WordPress, NO REST API, NO sitemap, NO search.
  * Products live at  /tool/{slug}.html  (bolt-cutter, wrc20-wirecutter, ...).
  * Each page = one or more spec TABLES. First row = headers
    (Item No | Cutting Capacity | Output | Weight ...), following rows = models.
    "Item No" column == the model / SKU.
  * A model's specs are sometimes split across two adjacent tables
    ([Item No + capacity] then [Output + weight]) -> merged by position here.
  * Pages served without a charset header -> force UTF-8 decode.
  * No sitemap -> enumerate by BFS over internal /tool/*.html links.

Three user-selectable pipelines:
  1) Slug / URL    - paste tool-page slugs or full URLs (one per line)
  2) Crawl ALL     - BFS-discover every /tool/*.html page
  3) Excel column  - pick file -> sheet -> column of slugs/URLs
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

BASE = "https://www.armsangyo.co.jp"
HOST = "www.armsangyo.co.jp"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

ITEMNO_RE = re.compile(r"item\s*no|品番|型式|型番|model|catalog|cat\.?\s*no", re.I)
SKIP_IMG = ("logo", "totop", "cata-bana", "/navi/", "down.png", "con_", "jis")

# ---- anti-loop / safety limits ----
MAX_REDIRECTS = 5
CIRCUIT_BREAK = 25          # abort after N consecutive page failures (IP block)
MAX_ITEMS = 5000            # runaway guard on work-list size
MAX_BFS_PAGES = 2000        # runaway guard on crawl-all discovery


# ----------------------------------------------------------------------------
# Core helpers (pure, no GUI)
# ----------------------------------------------------------------------------
def norm_code(s: str) -> str:
    """Normalize a model code for matching: drop whitespace, upper-case."""
    return re.sub(r"\s+", "", str(s)).upper()


def build_page_url(entry: str) -> str:
    """Accept a slug, a *.html name, a /tool/ path, or a full URL -> page URL."""
    e = entry.strip()
    if e.startswith("http"):
        return e
    if e.startswith("/"):
        return BASE + e
    if not e.endswith(".html"):
        e = e + ".html"
    if not e.startswith("tool/"):
        e = "tool/" + e
    return f"{BASE}/{e}"


def _abs_img(src: str, page_url: str) -> str:
    if not src:
        return ""
    return urljoin(page_url, src)


def _img_key(src: str) -> str:
    """Alphanumeric key from an image filename for matching against model codes.

    ../img/tool/lclamp/ld-300.jpg -> 'ld300'  ;  s-bc16li.jpg -> 'bc16li'
    """
    base = src.rsplit("/", 1)[-1].rsplit(".", 1)[0]
    base = re.sub(r"^(s[-_]|m[-_]|thumb[-_]?)", "", base, flags=re.I)   # drop thumb markers
    return re.sub(r"[^0-9a-z]", "", base.lower())


def _collect_group_images(tree, page_url: str) -> list:
    """Group/product photos in reading order: [(abs_url, key), ...].

    Prefers the .tool_list_1 / .tool_list_all group blocks (the horizontal row
    of product photos); falls back to any /img/tool/ image.
    """
    nodes = tree.css("div.tool_list_1 img") or tree.css("div.tool_list_all img")
    if not nodes:
        nodes = tree.css("img")
    out, seen = [], set()
    for img in nodes:
        src = img.attributes.get("src") or img.attributes.get("data-src") or ""
        low = src.lower()
        if "/img/tool/" not in low or any(s in low for s in SKIP_IMG):
            continue
        url = _abs_img(src, page_url)
        if url in seen:
            continue
        seen.add(url)
        out.append((url, _img_key(src)))
    return out


def _pick_image(item_no: str, imgs: list) -> str:
    """Pick the group image whose filename best prefix-matches the model code."""
    if not imgs:
        return ""
    key = re.sub(r"[^0-9a-z]", "", item_no.lower())
    best, best_n = imgs[0][0], 1        # default to first group image
    for url, k in imgs:
        n = 0
        for a, b in zip(key, k):
            if a == b:
                n += 1
            else:
                break
        if n > best_n:
            best_n, best = n, url
    return best


def _table_rows(table):
    rows = []
    for tr in table.css("tr"):
        cells = [c.text(separator=" ", strip=True) for c in tr.css("th,td")]
        cells = [re.sub(r"\s+", " ", c).strip() for c in cells]
        if any(cells):
            rows.append(cells)
    return rows


def parse_page(html: str, entry: str, url: str) -> list:
    """Parse an ARM tool page into a list of model records (one per table row)."""
    tree = LexborHTMLParser(html)
    h1 = tree.css_first("h1")
    page_name = h1.text(strip=True) if h1 else ""
    group_imgs = _collect_group_images(tree, url)
    main_img = group_imgs[0][0] if group_imgs else ""

    records = []
    last_models = []          # models from the most recent Item-No table
    for table in tree.css("table"):
        rows = _table_rows(table)
        if len(rows) < 2:
            continue
        header = rows[0]
        # pad/truncate data rows to header width
        parsed = []
        for r in rows[1:]:
            r = (r + [""] * len(header))[:len(header)]
            parsed.append(dict(zip(header, r)))
        if not parsed:
            continue

        has_item = any(ITEMNO_RE.search(h) for h in header)
        if has_item:
            item_key = next(h for h in header if ITEMNO_RE.search(h))
            last_models = []
            for d in parsed:
                item = d.get(item_key, "").strip()
                if not item:
                    continue
                rec = {"Input": entry, "Page": page_name, "Item No": item,
                       "Image": _pick_image(item, group_imgs) or main_img,
                       "URL": url, "Status": "Success"}
                rec.update(d)
                records.append(rec)
                last_models.append(rec)
        elif last_models:
            # supplementary spec table -> merge into the preceding models by position
            for i, d in enumerate(parsed):
                if len(parsed) == 1:
                    targets = last_models
                elif i < len(last_models):
                    targets = [last_models[i]]
                else:
                    targets = []
                for tgt in targets:
                    for k, v in d.items():
                        tgt.setdefault(k, v)

    if not records:
        return [{"Input": entry, "Page": page_name, "Item No": "",
                 "Image": main_img, "URL": url, "Status": "No models"}]
    return records


def extract_tool_links(html: str, page_url: str) -> list:
    """Return absolute /tool/*.html URLs found on a page (for BFS crawl-all)."""
    tree = LexborHTMLParser(html)
    out = []
    for a in tree.css("a"):
        href = a.attributes.get("href") or ""
        if not href or href.startswith("#") or href.startswith("mailto:"):
            continue
        full = urljoin(page_url, href)
        p = urlparse(full)
        if p.netloc == HOST and p.path.startswith("/tool/") and p.path.endswith(".html"):
            out.append(f"{BASE}{p.path}")
    return out


# ----------------------------------------------------------------------------
# GUI application
# ----------------------------------------------------------------------------
class ArmCrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ARM Sangyo Crawler  |  armsangyo.co.jp")
        self.root.geometry("980x690")
        self.root.minsize(880, 600)

        self.results = []
        self.log_q = queue.Queue()
        self.stop_evt = threading.Event()
        self.pause_evt = threading.Event()
        self.file_rows = []
        self.is_running = False

        self.bg = "#f4f5f7"
        self.primary = "#0b5c9c"       # ARM blue
        self.accent = "#1878c0"
        self.text = "#222222"

        self._styles()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._drain_logs)

    def _on_close(self):
        if self.is_running:
            if not messagebox.askyesno("Thoát", "Đang crawl. Thoát và tắt luôn?"):
                return
        self.stop_evt.set()
        try:
            self.root.destroy()
        finally:
            os._exit(0)

    # ---- styling -----------------------------------------------------------
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

    # ---- layout ------------------------------------------------------------
    def _build_ui(self):
        self.root.configure(bg=self.bg)
        header = ttk.Frame(self.root, padding=(15, 10))
        header.pack(fill=X)
        ttk.Label(header, text="ARM SANGYO CRAWLER", style="Title.TLabel").pack(side=LEFT)
        ttk.Label(header, text="  httpx + selectolax  •  static-site BFS",
                  font=("Segoe UI", 9, "italic"), foreground="#666").pack(side=LEFT, pady=5)

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=BOTH, expand=YES)

        left = ttk.LabelFrame(main, text=" Pipeline & Input ", padding=12, width=350)
        left.pack(side=LEFT, fill=Y, padx=(0, 10))
        left.pack_propagate(False)

        self.pipeline = StringVar(value="model")
        ttk.Label(left, text="Chọn pipeline:", font=("Segoe UI Semibold", 9)).pack(anchor=W)
        for val, txt in (("model", "1. Theo MÃ MODEL (Item No)"),
                         ("codes", "2. Theo slug / URL trang tool"),
                         ("all",   "3. Crawl ALL (BFS toàn site)"),
                         ("excel", "4. Theo cột trong Excel")):
            ttk.Radiobutton(left, text=txt, value=val, variable=self.pipeline,
                            command=self._on_pipeline).pack(anchor=W, pady=1)

        # --- model-code search (crawl-all + filter by Item No) ---
        self.model_frame = ttk.Frame(left)
        ttk.Label(self.model_frame, text="Mã model (mỗi dòng 1 mã):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.model_text = Text(self.model_frame, height=10, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.model_text.pack(fill=BOTH, expand=YES)
        ttk.Label(self.model_frame, text="vd: LD-500  LH-100  GP-200\n(quét toàn site rồi lọc đúng mã)",
                  font=("Segoe UI", 8, "italic"), foreground="#888").pack(anchor=W)

        self.codes_frame = ttk.Frame(left)
        ttk.Label(self.codes_frame, text="Slug hoặc URL (mỗi dòng 1 mục):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.codes_text = Text(self.codes_frame, height=10, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.codes_text.pack(fill=BOTH, expand=YES)
        ttk.Label(self.codes_frame, text="vd: bolt-cutter  /  pipe-wrench.html",
                  font=("Segoe UI", 8, "italic"), foreground="#888").pack(anchor=W)

        self.excel_frame = ttk.Frame(left)
        top = ttk.Frame(self.excel_frame); top.pack(fill=X, pady=(8, 4))
        ttk.Label(top, text="File Excel:", font=("Segoe UI Semibold", 9)).pack(side=LEFT)
        ttk.Button(top, text="Chọn file...", width=12, command=self._browse_excel).pack(side=RIGHT)
        self.excel_cfg = ttk.Frame(self.excel_frame); self.excel_cfg.pack(fill=X)
        ttk.Label(self.excel_cfg, text="Chưa chọn file",
                  font=("Segoe UI", 9, "italic"), foreground="#888").pack(pady=8)
        self.excel_as_model = BooleanVar(value=True)
        Checkbutton(self.excel_frame, text="Cột là MÃ MODEL (tìm theo mã)",
                    variable=self.excel_as_model, bg=self.bg, fg=self.text,
                    font=("Segoe UI", 9), activebackground=self.bg,
                    anchor=W).pack(fill=X, pady=(4, 0))

        self.all_frame = ttk.Frame(left)
        ttk.Label(self.all_frame,
                  text="BFS toàn bộ trang /tool/*.html\nrồi bóc mọi model trong bảng.",
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
        self._log("[SYS] Sẵn sàng crawl armsangyo.co.jp.\n")

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

    # ---- excel config ------------------------------------------------------
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
            self._log(f"[ERR] Không đọc được file: {e}\n")
            return
        ttk.Label(self.excel_cfg, text=os.path.basename(path),
                  font=("Segoe UI Semibold", 8), foreground=self.primary).pack(anchor=W)
        row = ttk.Frame(self.excel_cfg); row.pack(fill=X, pady=4)
        ttk.Label(row, text="Sheet:", width=6).pack(side=LEFT)
        sheet_cb = ttk.Combobox(row, state="readonly", width=16, values=sheets); sheet_cb.pack(side=LEFT, padx=4)
        row2 = ttk.Frame(self.excel_cfg); row2.pack(fill=X, pady=4)
        ttk.Label(row2, text="Cột:", width=6).pack(side=LEFT)
        col_cb = ttk.Combobox(row2, state="readonly", width=16); col_cb.pack(side=LEFT, padx=4)
        rowdata = {"path": path, "sheet": sheet_cb, "col": col_cb}
        self.file_rows = [rowdata]
        sheet_cb.bind("<<ComboboxSelected>>", lambda e: self._load_cols(rowdata))
        if sheets:
            sheet_cb.set(sheets[0]); self._load_cols(rowdata)
        self._log(f"[SYS] Đã nạp file: {os.path.basename(path)}\n")

    def _load_cols(self, rowdata):
        try:
            df = pd.read_excel(rowdata["path"], sheet_name=rowdata["sheet"].get(), nrows=5)
        except Exception as e:
            self._log(f"[ERR] Đọc sheet lỗi: {e}\n"); return
        cols = list(df.columns)
        rowdata["col"]["values"] = cols
        pick = next((c for c in cols if re.search(r"slug|url|mã|code|item|品番", str(c), re.I)),
                    cols[0] if cols else "")
        if pick:
            rowdata["col"].set(pick)

    # ---- console + queue ---------------------------------------------------
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
        self.prog_lbl.config(text="Đã dừng" if aborted else "Hoàn tất")
        if not aborted:
            messagebox.showinfo("Xong", f"Crawl xong {len(self.results)} model. Bấm EXPORT EXCEL.")

    def _toggle_pause(self):
        if self.pause_evt.is_set():
            self.pause_evt.clear(); self.pause_btn.config(text="Pause")
            self.log_q.put(("log", "[SYS] Tiếp tục.\n"))
        else:
            self.pause_evt.set(); self.pause_btn.config(text="Resume")
            self.log_q.put(("log", "[SYS] Tạm dừng.\n"))

    def _stop(self):
        if messagebox.askyesno("Dừng", "Dừng crawl? Kết quả hiện có sẽ được giữ lại."):
            self.stop_evt.set(); self.pause_evt.clear()

    # ---- start -------------------------------------------------------------
    def _start(self):
        try:
            conc = int(self.conc.get()); delay = float(self.delay.get())
        except ValueError:
            messagebox.showerror("Lỗi", "Concurrency phải là số nguyên, Delay là số thực."); return
        self.results = []
        self.stop_evt.clear(); self.pause_evt.clear()
        self.is_running = True
        self.start_btn.config(state=DISABLED)
        self.pause_btn.config(state=NORMAL); self.stop_btn.config(state=NORMAL)
        self.export_btn.config(state=DISABLED)
        self.console.config(state=NORMAL); self.console.delete("1.0", END); self.console.config(state=DISABLED)
        threading.Thread(target=self._run_async,
                         args=(self.pipeline.get(), conc, delay), daemon=True).start()

    def _gather_codes(self):
        raw = self.codes_text.get("1.0", "end-1c")
        return [x.strip() for x in re.split(r"[\r\n]+", raw) if x.strip()]

    def _gather_model_codes(self):
        raw = self.model_text.get("1.0", "end-1c")
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
                self.log_q.put(("log", f"[WARN] Không thấy cột '{col}'.\n"))
        return out

    # ---- async engine ------------------------------------------------------
    def _run_async(self, pipeline, conc, delay):
        try:
            asyncio.run(self._crawl(pipeline, conc, delay))
        except Exception as e:
            self.log_q.put(("log", f"[FATAL] {e}\n"))
            self.log_q.put(("done", True))

    async def _get_text(self, client, url):
        """Fetch and force UTF-8 decode (site sends no charset header)."""
        r = await client.get(url)
        final = str(r.url)
        if r.status_code != 200:
            raise httpx.HTTPStatusError(f"HTTP {r.status_code}", request=r.request, response=r)
        return r.content.decode("utf-8", "replace"), final

    async def _crawl(self, pipeline, conc, delay):
        limits = httpx.Limits(max_connections=conc, max_keepalive_connections=conc)
        transport = httpx.AsyncHTTPTransport(retries=2)
        async with httpx.AsyncClient(headers={"User-Agent": UA}, timeout=20.0,
                                     limits=limits, transport=transport,
                                     follow_redirects=True, max_redirects=MAX_REDIRECTS) as client:

            match_set = None       # {normalized_code: original} -> model-search mode
            if pipeline == "model":
                codes = self._gather_model_codes()
                if not codes:
                    self.log_q.put(("log", "[WARN] Chưa nhập mã model nào.\n"))
                    self.log_q.put(("done", True)); return
                match_set = {}
                for c in codes:
                    match_set.setdefault(norm_code(c), c)
                self.log_q.put(("log",
                    f"[SYS] Tìm {len(match_set)} mã model -> quét toàn site...\n"))
                work = await self._bfs_pages(client, conc)
            elif pipeline == "all":
                self.log_q.put(("log", "[SYS] BFS phát hiện trang sản phẩm...\n"))
                work = await self._bfs_pages(client, conc)
            elif pipeline == "excel":
                vals = self._gather_excel()
                if self.excel_as_model.get():
                    if not vals:
                        self.log_q.put(("log", "[WARN] Cột không có mã nào.\n"))
                        self.log_q.put(("done", True)); return
                    match_set = {}
                    for c in vals:
                        match_set.setdefault(norm_code(c), c)
                    self.log_q.put(("log",
                        f"[SYS] Excel: tìm {len(match_set)} mã model -> quét toàn site...\n"))
                    work = await self._bfs_pages(client, conc)
                else:
                    work = [build_page_url(e) for e in vals]
            else:
                work = [build_page_url(e) for e in self._gather_codes()]

            # dedup keep order
            seen, uniq = set(), []
            for u in work:
                if u not in seen:
                    seen.add(u); uniq.append(u)
            work = uniq

            total = len(work)
            if not total:
                self.log_q.put(("log", "[WARN] Không có trang nào để crawl.\n"))
                self.log_q.put(("done", True)); return
            if total > MAX_ITEMS:
                self.log_q.put(("log", f"[WARN] {total} trang vượt {MAX_ITEMS}, cắt bớt.\n"))
                work = work[:MAX_ITEMS]; total = MAX_ITEMS

            self.log_q.put(("log", f"[SYS] Crawl {total} trang | concurrency={conc}\n"))
            sem = asyncio.Semaphore(conc)
            done = models = consec_fail = 0
            found = set()          # normalized codes matched (model mode)
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
                    try:
                        html, final = await self._get_text(client, url)
                        if "/tool/" not in final:   # anti-loop: bounced off product area
                            recs = [{"Input": url, "URL": final, "Status": "Redirect off /tool/"}]
                            ok = False
                        else:
                            recs = parse_page(html, url, final)
                            ok = any(r.get("Status") == "Success" for r in recs)
                    except httpx.TooManyRedirects:
                        recs = [{"Input": url, "URL": url, "Status": "Redirect loop"}]; ok = False
                    except Exception as e:
                        recs = [{"Input": url, "URL": url, "Status": f"Error: {type(e).__name__}"}]; ok = False

                # model-search: keep only rows whose Item No matches a requested code
                if match_set is not None:
                    kept = []
                    for r in recs:
                        nc = norm_code(r.get("Item No", ""))
                        if nc in match_set:
                            r = {**r, "Input": match_set[nc]}
                            kept.append(r)
                    recs = kept

                async with lock:
                    self.results.extend(recs)
                    done += 1
                    n = sum(1 for r in recs if r.get("Status") == "Success")
                    models += n
                    if match_set is not None:
                        for r in recs:
                            found.add(norm_code(r.get("Item No", "")))
                        consec_fail = 0
                        if n:
                            self.log_q.put(("log",
                                f"[HIT {n}] {url.split('/tool/')[-1]}\n"))
                    else:
                        consec_fail = 0 if ok else consec_fail + 1
                        tag = f"OK  {n} model" if ok else recs[0].get("Status", "MISS")
                        self.log_q.put(("log", f"[{tag}] {url.split('/tool/')[-1]}\n"))
                    self.log_q.put(("prog", (done, total, models)))
                    if consec_fail >= CIRCUIT_BREAK and not self.stop_evt.is_set():
                        self.stop_evt.set()
                        self.log_q.put(("log",
                            f"[STOP] {CIRCUIT_BREAK} lỗi liên tiếp -> ngắt an toàn.\n"))
                if delay:
                    await asyncio.sleep(delay)

            await asyncio.gather(*(one(u) for u in work))

            # model-search: list codes never found
            if match_set is not None:
                for nc, orig in match_set.items():
                    if nc not in found:
                        self.results.append({"Input": orig, "Item No": orig,
                                             "Status": "Not Found"})
                        self.log_q.put(("log", f"[MISS] {orig} không thấy trên site.\n"))
                self.log_q.put(("log",
                    f"[SYS] Khớp {len(found & set(match_set))}/{len(match_set)} mã.\n"))
            self.log_q.put(("done", self.stop_evt.is_set()))

    async def _bfs_pages(self, client, conc):
        """Breadth-first discovery of all /tool/*.html pages from the homepage."""
        discovered = set()
        frontier = {BASE + "/", BASE + "/tool/"}
        visited = set()
        sem = asyncio.Semaphore(conc)

        async def grab(url):
            async with sem:
                try:
                    html, final = await self._get_text(client, url)
                    return extract_tool_links(html, final)
                except Exception:
                    return []

        while frontier and len(discovered) < MAX_BFS_PAGES and not self.stop_evt.is_set():
            batch = [u for u in frontier if u not in visited]
            visited.update(batch)
            frontier = set()
            results = await asyncio.gather(*(grab(u) for u in batch))
            for links in results:
                for l in links:
                    if l not in discovered:
                        discovered.add(l)
                        if l not in visited:
                            frontier.add(l)
        pages = sorted(discovered)
        self.log_q.put(("log", f"[SYS] Tìm thấy {len(pages)} trang sản phẩm.\n"))
        return pages

    # ---- export ------------------------------------------------------------
    def _export(self):
        if not self.results:
            messagebox.showwarning("Trống", "Chưa có dữ liệu."); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")],
                                            initialfile="armtools_data.xlsx")
        if not path:
            return
        # union of all keys, preferred order first
        lead = ["Input", "Page", "Item No", "Status"]
        tail = ["Image", "URL"]
        spec_cols = []
        for r in self.results:
            for k in r:
                if k not in lead and k not in tail and k not in spec_cols:
                    spec_cols.append(k)
        ordered = lead + spec_cols + tail
        df = pd.DataFrame(self.results).reindex(columns=ordered)
        try:
            df.to_excel(path, index=False)
            self._log(f"[SYS] Đã xuất -> {path}\n")
            messagebox.showinfo("Xong", f"Đã lưu:\n{path}")
        except Exception as e:
            messagebox.showerror("Lỗi xuất", str(e))


if __name__ == "__main__":
    root = Tk()
    ArmCrawlerApp(root)
    root.mainloop()
