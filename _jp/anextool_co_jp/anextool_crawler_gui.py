"""
AnexTool High-Speed Product Crawler  (anextool.co.jp)
=====================================================
Modern rewrite: httpx (async pooled HTTP) + selectolax (Lexbor parser).
Replaces the old requests + BeautifulSoup path (slow, no parsing logic).

Target site = WordPress, product custom-post-type `item`.
Key facts discovered by probing the live site:
  * Product page          : https://www.anextool.co.jp/item/{slug}/
  * URL slug == normalized maker part-no  (ACHX-2010 -> achx-2010,
                                            No.7765-6x100 -> 7765-6-100)
    -> O(1) direct lookup, the built-in "?s=" search is broken/irrelevant.
  * Specs are <dl><dt>label</dt><dd>value</dd></dl> pairs (メーカー品番, JAN,
    定価, 製造国, 対応サイズ ...).
  * Images live under /wp-content/uploads/{CODE}_N-768x768.jpg
    -> strip the -WxH suffix for full-resolution.
  * All products enumerable via item-sitemap*.xml  (crawl-all pipeline).

Three user-selectable pipelines:
  1) Codes / URLs   - paste maker codes or full /item/ URLs (one per line)
  2) Crawl ALL      - pull every product from the sitemaps
  3) Excel column   - pick file -> sheet -> code column
"""

import os
import re
import sys
import queue
import asyncio
import threading
from tkinter import *
from tkinter import ttk, filedialog, messagebox

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser

BASE = "https://www.anextool.co.jp"
SITEMAPS = [f"{BASE}/item-sitemap.xml", f"{BASE}/item-sitemap2.xml"]
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# dl labels that are page furniture, not product specs -> dropped from output
SKIP_LABELS = {"製品画像", "製品データ", "本社・工場", "福岡営業所", "お問い合わせ"}

# ---- anti-loop / safety limits ----
MAX_REDIRECTS = 5           # httpx redirect cap (WP 404 -> homepage bounce)
CIRCUIT_BREAK = 25          # abort after this many CONSECUTIVE failures (IP block)
MAX_ITEMS = 50000           # hard cap on work-list size (runaway guard)


# ----------------------------------------------------------------------------
# Core helpers (pure, testable, no GUI)
# ----------------------------------------------------------------------------
def code_to_slug(code: str) -> str:
    """Normalize a maker part-number into the WordPress item slug.

    ACHX-2010      -> achx-2010
    No.7765-6x100  -> 7765-6-100
    No.7900+2x200  -> 7900-2-200
    """
    s = code.strip()
    s = re.sub(r"^No\.?\s*", "", s, flags=re.I)      # drop leading "No."
    s = re.sub(r"(?<=\d)\s*[xX×＊*]\s*(?=\d)", "-", s)  # 6x100 -> 6-100
    s = re.sub(r"(?<=\d)\s*[+＋]\s*(?=\d)", "-", s)     # 7900+2 -> 7900-2
    s = s.replace(".", "-").replace("/", "-")
    s = re.sub(r"\s+", "-", s).lower()
    s = re.sub(r"-+", "-", s).strip("-")
    return s


def build_item_url(entry: str) -> str:
    """Accept a raw code, a slug, or a full URL -> canonical item URL."""
    entry = entry.strip()
    if entry.startswith("http"):
        return entry.rstrip("/") + "/"
    if entry.startswith("/item/"):
        return BASE + entry.rstrip("/") + "/"
    return f"{BASE}/item/{code_to_slug(entry)}/"


def _full_res(img_url: str) -> str:
    if not img_url:
        return ""
    if img_url.startswith("//"):
        img_url = "https:" + img_url
    # AKL-305_2-768x768.jpg -> AKL-305_2.jpg
    return re.sub(r"-\d+x\d+(?=\.\w+$)", "", img_url)


def parse_product(html: str, code: str, url: str) -> dict:
    """Extract product data from an item page using selectolax (Lexbor)."""
    tree = LexborHTMLParser(html)

    # Name
    h1 = tree.css_first("h1")
    name = h1.text(strip=True) if h1 else ""

    # Spec pairs: <dl><dt>..</dt><dd>..</dd></dl>
    specs = {}
    part_no = ""
    for dl in tree.css("dl"):
        dt = dl.css_first("dt")
        dd = dl.css_first("dd")
        if not dt or not dd:
            continue
        label = dt.text(strip=True)
        if not label or label in SKIP_LABELS:
            continue
        # keep multi-line dd content (bullet lists) readable
        value = re.sub(r"[ \t]+", " ", dd.text(separator="\n", strip=True)).strip()
        if not value:
            continue
        specs.setdefault(label, value)
        if "品番" in label and not part_no:
            part_no = value

    # Main image (first product photo under /uploads/), full resolution
    img_url = ""
    for img in tree.css("img"):
        src = img.attributes.get("src") or img.attributes.get("data-src") or ""
        if "/uploads/" in src:
            img_url = _full_res(src)
            break

    # Category / breadcrumb if present
    cats = []
    for sel in (".breadcrumb a", ".breadcrumbs a", ".pankuzu a", "nav[aria-label] a"):
        nodes = tree.css(sel)
        if nodes:
            for a in nodes:
                t = a.text(strip=True).strip("/ ")
                if t and t.lower() not in ("home", "top", "ホーム", "トップ"):
                    cats.append(t)
            break

    return {
        "Input": code,
        "Part No": part_no,
        "Name": name,
        "Image": img_url,
        "URL": url,
        "categories": cats,
        "specs": specs,
        "Status": "Success",
    }


# ----------------------------------------------------------------------------
# GUI application
# ----------------------------------------------------------------------------
class AnexCrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AnexTool High-Speed Crawler  |  anextool.co.jp")
        self.root.geometry("960x680")
        self.root.minsize(860, 600)

        # runtime state
        self.results = []
        self.log_q = queue.Queue()
        self.stop_evt = threading.Event()
        self.pause_evt = threading.Event()
        self.file_rows = []            # excel config rows
        self.is_running = False

        # theme
        self.bg = "#f5f6f8"
        self.primary = "#c8102e"       # ANEX red
        self.accent = "#e63c50"
        self.text = "#222222"

        self._styles()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._drain_logs)

    def _on_close(self):
        # confirm if a crawl is live, then hard-exit so the terminal closes too
        if self.is_running:
            if not messagebox.askyesno("Thoát", "Đang crawl. Thoát và tắt luôn?"):
                return
        self.stop_evt.set()
        try:
            self.root.destroy()
        finally:
            os._exit(0)   # force-kill: no lingering threads / console window

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
        ttk.Label(header, text="ANEXTOOL CRAWLER", style="Title.TLabel").pack(side=LEFT)
        ttk.Label(header, text="  httpx + selectolax  •  slug-direct lookup",
                  font=("Segoe UI", 9, "italic"), foreground="#666").pack(side=LEFT, pady=5)

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=BOTH, expand=YES)

        left = ttk.LabelFrame(main, text=" Pipeline & Input ", padding=12, width=340)
        left.pack(side=LEFT, fill=Y, padx=(0, 10))
        left.pack_propagate(False)

        # --- pipeline selector ---
        self.pipeline = StringVar(value="codes")
        ttk.Label(left, text="Chọn pipeline:", font=("Segoe UI Semibold", 9)).pack(anchor=W)
        for val, txt in (("codes", "1. Theo mã / URL (dán bên dưới)"),
                         ("all",   "2. Crawl ALL (toàn bộ sitemap)"),
                         ("excel", "3. Theo cột mã trong Excel")):
            ttk.Radiobutton(left, text=txt, value=val, variable=self.pipeline,
                            command=self._on_pipeline).pack(anchor=W, pady=1)

        # --- codes/URL input ---
        self.codes_frame = ttk.Frame(left)
        ttk.Label(self.codes_frame, text="Mã hoặc URL (mỗi dòng 1 mục):",
                  font=("Segoe UI Semibold", 9)).pack(anchor=W, pady=(8, 2))
        self.codes_text = Text(self.codes_frame, height=10, wrap=WORD,
                               font=("Consolas", 9), relief=SOLID, borderwidth=1)
        self.codes_text.pack(fill=BOTH, expand=YES)

        # --- excel input ---
        self.excel_frame = ttk.Frame(left)
        top = ttk.Frame(self.excel_frame)
        top.pack(fill=X, pady=(8, 4))
        ttk.Label(top, text="File Excel:", font=("Segoe UI Semibold", 9)).pack(side=LEFT)
        ttk.Button(top, text="Chọn file...", width=12, command=self._browse_excel).pack(side=RIGHT)
        self.excel_cfg = ttk.Frame(self.excel_frame)
        self.excel_cfg.pack(fill=X)
        self.excel_hint = ttk.Label(self.excel_cfg, text="Chưa chọn file",
                                    font=("Segoe UI", 9, "italic"), foreground="#888")
        self.excel_hint.pack(pady=8)

        # --- all-info ---
        self.all_frame = ttk.Frame(left)
        ttk.Label(self.all_frame,
                  text="Sẽ tải toàn bộ mã sản phẩm\ntừ item-sitemap*.xml rồi crawl.",
                  font=("Segoe UI", 9), foreground="#555").pack(anchor=W, pady=10)

        # --- speed config ---
        speed = ttk.Frame(left)
        speed.pack(fill=X, pady=(12, 8), side=BOTTOM)
        ttk.Label(speed, text="Concurrency:", font=("Segoe UI Semibold", 9)).grid(row=0, column=0, sticky=W)
        self.conc = Spinbox(speed, from_=1, to=32, width=5, justify=CENTER)
        self.conc.delete(0, END); self.conc.insert(0, "10")
        self.conc.grid(row=0, column=1, sticky=E, padx=6, pady=3)
        ttk.Label(speed, text="Delay (s):", font=("Segoe UI Semibold", 9)).grid(row=1, column=0, sticky=W)
        self.delay = Spinbox(speed, from_=0.0, to=3.0, increment=0.05, width=5, justify=CENTER)
        self.delay.delete(0, END); self.delay.insert(0, "0.05")
        self.delay.grid(row=1, column=1, sticky=E, padx=6, pady=3)

        # --- control buttons ---
        ctrl = ttk.Frame(left)
        ctrl.pack(fill=X, side=BOTTOM)
        self.start_btn = ttk.Button(ctrl, text="START CRAWL", style="Primary.TButton", command=self._start)
        self.start_btn.pack(fill=X, pady=(0, 6))
        sub = ttk.Frame(ctrl); sub.pack(fill=X)
        self.pause_btn = ttk.Button(sub, text="Pause", state=DISABLED, command=self._toggle_pause)
        self.pause_btn.pack(side=LEFT, fill=X, expand=YES, padx=(0, 4))
        self.stop_btn = ttk.Button(sub, text="Stop", state=DISABLED, command=self._stop)
        self.stop_btn.pack(side=RIGHT, fill=X, expand=YES, padx=(4, 0))
        self.export_btn = ttk.Button(ctrl, text="EXPORT EXCEL", state=DISABLED, command=self._export)
        self.export_btn.pack(fill=X, pady=(6, 0))

        # --- right: console ---
        right = ttk.LabelFrame(main, text=" Console & Progress ", padding=12)
        right.pack(side=RIGHT, fill=BOTH, expand=YES)
        info = ttk.Frame(right); info.pack(fill=X)
        self.prog_lbl = ttk.Label(info, text="Idle", font=("Segoe UI Semibold", 10))
        self.prog_lbl.pack(side=LEFT)
        self.stats_lbl = ttk.Label(info, text="0 / 0  (0 lỗi)", foreground="#555")
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
        self._log("[SYS] Sẵn sàng crawl anextool.co.jp.\n")

    # ---- pipeline visibility ----------------------------------------------
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

    # ---- excel file config -------------------------------------------------
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
        sheet_cb = ttk.Combobox(row, state="readonly", width=14, values=sheets)
        sheet_cb.pack(side=LEFT, padx=4)
        row2 = ttk.Frame(self.excel_cfg); row2.pack(fill=X, pady=4)
        ttk.Label(row2, text="Cột mã:", width=6).pack(side=LEFT)
        col_cb = ttk.Combobox(row2, state="readonly", width=14)
        col_cb.pack(side=LEFT, padx=4)

        rowdata = {"path": path, "sheet": sheet_cb, "col": col_cb}
        self.file_rows = [rowdata]
        sheet_cb.bind("<<ComboboxSelected>>", lambda e: self._load_cols(rowdata))
        if sheets:
            sheet_cb.set(sheets[0])
            self._load_cols(rowdata)
        self._log(f"[SYS] Đã nạp file: {os.path.basename(path)}\n")

    def _load_cols(self, rowdata):
        try:
            df = pd.read_excel(rowdata["path"], sheet_name=rowdata["sheet"].get(), nrows=5)
        except Exception as e:
            self._log(f"[ERR] Đọc sheet lỗi: {e}\n")
            return
        cols = list(df.columns)
        rowdata["col"]["values"] = cols
        pick = next((c for c in cols if re.search(r"mã|code|sku|品番", str(c), re.I)), cols[0] if cols else "")
        if pick:
            rowdata["col"].set(pick)

    # ---- console + queue ---------------------------------------------------
    def _log(self, text):
        self.console.config(state=NORMAL)
        self.console.insert(END, text)
        self.console.see(END)
        self.console.config(state=DISABLED)

    def _drain_logs(self):
        while not self.log_q.empty():
            kind, data = self.log_q.get()
            if kind == "log":
                self._log(data)
            elif kind == "prog":
                done, total, failed = data
                self.pbar["maximum"] = total
                self.pbar["value"] = done
                pct = int(done / total * 100) if total else 0
                self.prog_lbl.config(text=f"Crawling {pct}%")
                self.stats_lbl.config(text=f"{done} / {total}  ({failed} lỗi)")
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
            messagebox.showinfo("Xong", f"Crawl xong {len(self.results)} sản phẩm. Bấm EXPORT EXCEL.")

    # ---- controls ----------------------------------------------------------
    def _toggle_pause(self):
        if self.pause_evt.is_set():
            self.pause_evt.clear()
            self.pause_btn.config(text="Pause")
            self.log_q.put(("log", "[SYS] Tiếp tục.\n"))
        else:
            self.pause_evt.set()
            self.pause_btn.config(text="Resume")
            self.log_q.put(("log", "[SYS] Tạm dừng.\n"))

    def _stop(self):
        if messagebox.askyesno("Dừng", "Dừng crawl? Kết quả hiện có sẽ được giữ lại."):
            self.stop_evt.set()
            self.pause_evt.clear()

    # ---- start -------------------------------------------------------------
    def _start(self):
        try:
            conc = int(self.conc.get())
            delay = float(self.delay.get())
        except ValueError:
            messagebox.showerror("Lỗi", "Concurrency phải là số nguyên, Delay là số thực.")
            return

        self.results = []
        self.stop_evt.clear()
        self.pause_evt.clear()
        self.is_running = True
        self.start_btn.config(state=DISABLED)
        self.pause_btn.config(state=NORMAL)
        self.stop_btn.config(state=NORMAL)
        self.export_btn.config(state=DISABLED)
        self.console.config(state=NORMAL); self.console.delete("1.0", END); self.console.config(state=DISABLED)

        threading.Thread(target=self._run_async,
                         args=(self.pipeline.get(), conc, delay), daemon=True).start()

    def _gather_codes(self):
        raw = self.codes_text.get("1.0", "end-1c")
        return [x.strip() for x in re.split(r"[\r\n]+", raw) if x.strip()]

    def _gather_excel(self):
        out = []
        for r in self.file_rows:
            try:
                df = pd.read_excel(r["path"], sheet_name=r["sheet"].get())
            except Exception as e:
                self.log_q.put(("log", f"[ERR] {e}\n"))
                continue
            col = r["col"].get()
            if col in df.columns:
                out += [str(v).strip() for v in df[col].dropna() if str(v).strip()]
            else:
                self.log_q.put(("log", f"[WARN] Không thấy cột '{col}'.\n"))
        return out

    # ---- async engine (runs in worker thread) ------------------------------
    def _run_async(self, pipeline, conc, delay):
        try:
            asyncio.run(self._crawl(pipeline, conc, delay))
        except Exception as e:
            self.log_q.put(("log", f"[FATAL] {e}\n"))
            self.log_q.put(("done", True))

    async def _crawl(self, pipeline, conc, delay):
        limits = httpx.Limits(max_connections=conc, max_keepalive_connections=conc)
        transport = httpx.AsyncHTTPTransport(retries=2)
        async with httpx.AsyncClient(headers={"User-Agent": UA}, timeout=20.0,
                                     limits=limits, transport=transport,
                                     follow_redirects=True, max_redirects=MAX_REDIRECTS) as client:

            # --- resolve the work list per pipeline ---
            if pipeline == "all":
                self.log_q.put(("log", "[SYS] Đang tải danh sách sản phẩm từ sitemap...\n"))
                entries = await self._load_sitemaps(client)
            elif pipeline == "excel":
                entries = self._gather_excel()
            else:
                entries = self._gather_codes()

            # dedup, keep order
            seen, work = set(), []
            for e in entries:
                if e not in seen:
                    seen.add(e); work.append(e)

            total = len(work)
            if not total:
                self.log_q.put(("log", "[WARN] Không có mã/URL nào để crawl.\n"))
                self.log_q.put(("done", True))
                return
            if total > MAX_ITEMS:
                self.log_q.put(("log", f"[WARN] {total} mục vượt giới hạn {MAX_ITEMS}, cắt bớt.\n"))
                work = work[:MAX_ITEMS]
                total = MAX_ITEMS

            self.log_q.put(("log", f"[SYS] Bắt đầu crawl {total} mục | concurrency={conc}\n"))
            sem = asyncio.Semaphore(conc)
            done = failed = consec_fail = 0
            lock = asyncio.Lock()

            async def one(entry):
                nonlocal done, failed, consec_fail
                # pause / stop gate
                while self.pause_evt.is_set() and not self.stop_evt.is_set():
                    await asyncio.sleep(0.2)
                if self.stop_evt.is_set():
                    return
                async with sem:
                    if self.stop_evt.is_set():   # re-check after acquiring slot
                        return
                    url = build_item_url(entry)
                    rec = await self._fetch_one(client, entry, url)
                async with lock:
                    self.results.append(rec)
                    done += 1
                    ok = rec["Status"] == "Success"
                    if ok:
                        consec_fail = 0
                    else:
                        failed += 1
                        consec_fail += 1
                    tag = "OK " if ok else "MISS"
                    self.log_q.put(("log", f"[{tag}] {entry} -> {rec.get('Name','') or rec['Status']}\n"))
                    self.log_q.put(("prog", (done, total, failed)))
                    # circuit breaker: too many consecutive failures -> likely blocked
                    if consec_fail >= CIRCUIT_BREAK and not self.stop_evt.is_set():
                        self.stop_evt.set()
                        self.log_q.put(("log",
                            f"[STOP] {CIRCUIT_BREAK} lỗi liên tiếp -> ngắt an toàn (có thể bị chặn IP).\n"))
                if delay:
                    await asyncio.sleep(delay)

            await asyncio.gather(*(one(e) for e in work))
            self.log_q.put(("done", self.stop_evt.is_set()))

    async def _fetch_one(self, client, entry, url):
        blank = {"Input": entry, "URL": url, "Part No": "", "Name": "",
                 "Image": "", "categories": [], "specs": {}}
        try:
            r = await client.get(url)
            final = str(r.url)
            # anti-loop: WP 404 bounces to homepage/search -> not a real product page
            if "/item/" not in final:
                return {**blank, "Status": "Not Found (redirect)"}
            if r.status_code == 200:
                rec = parse_product(r.text, entry, final)
                if not rec["Name"] and not rec["Part No"]:
                    rec["Status"] = "Empty"
                return rec
            return {**blank, "Status": f"HTTP {r.status_code}"}
        except httpx.TooManyRedirects:
            return {**blank, "Status": "Redirect loop"}
        except Exception as e:
            return {**blank, "Status": f"Error: {type(e).__name__}"}

    async def _load_sitemaps(self, client):
        urls = []
        for sm in SITEMAPS:
            try:
                r = await client.get(sm)
                if r.status_code == 200:
                    urls += re.findall(r"https://www\.anextool\.co\.jp/item/[^<\]\s]+", r.text)
            except Exception as e:
                self.log_q.put(("log", f"[WARN] sitemap {sm}: {e}\n"))
        # unique, strip trailing junk
        clean = []
        seen = set()
        for u in urls:
            u = u.rstrip("/") + "/"
            if u not in seen:
                seen.add(u); clean.append(u)
        self.log_q.put(("log", f"[SYS] Sitemap: {len(clean)} sản phẩm.\n"))
        return clean

    # ---- export ------------------------------------------------------------
    def _export(self):
        if not self.results:
            messagebox.showwarning("Trống", "Chưa có dữ liệu.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")],
                                            initialfile="anextool_data.xlsx")
        if not path:
            return

        # collect all spec labels (dynamic columns), preserve first-seen order
        spec_cols = []
        for r in self.results:
            for k in r.get("specs", {}):
                if k not in spec_cols:
                    spec_cols.append(k)
        max_cat = max((len(r.get("categories", [])) for r in self.results), default=0)

        rows = []
        for r in self.results:
            row = {
                "Input": r.get("Input", ""),
                "Part No": r.get("Part No", ""),
                "Name": r.get("Name", ""),
                "Status": r.get("Status", ""),
                "Image": r.get("Image", ""),
                "URL": r.get("URL", ""),
            }
            cats = r.get("categories", [])
            for i in range(max_cat):
                row[f"Category {i+1}"] = cats[i] if i < len(cats) else ""
            for k in spec_cols:
                row[k] = r.get("specs", {}).get(k, "")
            rows.append(row)

        base = ["Input", "Part No", "Name", "Status"]
        cat_cols = [f"Category {i+1}" for i in range(max_cat)]
        tail = ["Image", "URL"]
        ordered = base + cat_cols + spec_cols + tail
        df = pd.DataFrame(rows).reindex(columns=ordered)
        try:
            df.to_excel(path, index=False)
            self._log(f"[SYS] Đã xuất -> {path}\n")
            messagebox.showinfo("Xong", f"Đã lưu:\n{path}")
        except Exception as e:
            messagebox.showerror("Lỗi xuất", str(e))


if __name__ == "__main__":
    root = Tk()
    AnexCrawlerApp(root)
    root.mainloop()
