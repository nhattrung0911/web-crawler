"""WADFOW Philippines crawler -- API-first, no Selenium.

Live recon (2026-08-10): wadfow.com/ph is a React/UMI SPA.  The public
gateway ``https://gatewayapi-sg.rdmcenter.com/website-common-api/`` exposes
``product`` (pageNum/pageSize/keyword), ``product/detail``
(productExtNo), and ``website-common/allCategory``.  Product detail contains
the parameter HTML, stock, price and gallery, so Playwright is not needed for
the crawler itself.  The API requires the WADFOW site-context headers below.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import pathlib
import queue
import re
import sys
import threading
import time
from typing import Any

import httpx
import pandas as pd

HERE = pathlib.Path(__file__).parent
API_BASE = "https://gatewayapi-sg.rdmcenter.com/website-common-api/"
SITE = "https://www.wadfow.com/ph"
HEADERS = {
    "domain": "www.wadfow.com/ph", "lang": "en_US", "timezoneoffset": "-480",
    "siteorigin": SITE, "referer": "https://www.wadfow.com/",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
}
TIMEOUT = 30.0
MAX_REDIRECTS = 5
CIRCUIT_BREAK = 25
MAX_ITEMS = 5000
MAX_BFS_PAGES = 1500
MAX_PAGES = 600


def clean(value: Any) -> str:
    """Make an API scalar safe for a spreadsheet cell."""
    text = re.sub(r"<br\s*/?>", "\n", str(value or ""), flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    return "\n".join(" ".join(part.split()) for part in text.splitlines() if part.strip())


def norm_code(value: Any) -> str:
    """Normalise product codes only; never uppercase a case-sensitive URL."""
    value = str(value or "").strip()
    return value if value.lower().startswith(("http://", "https://")) else re.sub(r"\s+", "", value).upper()


def extract_code(value: Any) -> str:
    value = str(value or "").strip()
    if value.lower().startswith(("http://", "https://")):
        return value.rstrip("/").split("?")[0].rsplit("/", 1)[-1]
    return value


def is_exact_code(query: Any, item: dict[str, Any]) -> bool:
    want = norm_code(query)
    return want in {norm_code(item.get("productExtNo")), norm_code(item.get("productNo"))}


def parse_parameters(value: Any) -> dict[str, str]:
    specs: dict[str, str] = {}
    for line in re.split(r"<br\s*/?>|\r?\n", str(value or ""), flags=re.I):
        line = clean(line)
        if ":" in line:
            key, val = line.split(":", 1)
            if clean(key):
                specs.setdefault(clean(key), clean(val))
        elif line:
            specs.setdefault(line, "")
    return specs


def product_url(data: dict[str, Any]) -> str:
    explicit = data.get("url") or data.get("productUrl") or data.get("linkUrl")
    if isinstance(explicit, str) and explicit.startswith("http"):
        return explicit
    code = str(data.get("productExtNo") or data.get("productNo") or "")
    slug = str(data.get("pathReplacement") or data.get("productName") or "product")
    slug = re.sub(r"[^a-z0-9]+", "-", slug.lower()).strip("-")
    return f"{SITE}/product/{slug}/{code}" if code else ""


def parse_detail(data: dict[str, Any]) -> dict[str, Any]:
    code = str(data.get("productExtNo") or data.get("productNo") or "")
    pics = data.get("productPicList") or data.get("productPictures") or []
    if isinstance(pics, str):
        pics = [pics]
    return {
        "Input": "", "Page": clean(data.get("productName") or data.get("name")),
        "Item No": code, "Status": "OK",
        "_cats": [clean(x) for x in (data.get("categoryName"), data.get("categoryPath")) if clean(x)],
        "Gia": data.get("price"), "Stock": data.get("stock"), "Weight": data.get("weight"),
        "_spec": parse_parameters(data.get("parameter")),
        "Image": " | ".join(str(x) for x in pics if x), "URL": product_url(data),
    }


def result_error(source: str, status: str) -> dict[str, Any]:
    return {"Input": source, "Page": "", "Item No": extract_code(source), "Status": status,
            "_cats": [], "Gia": None, "Stock": None, "Weight": None,
            "_spec": {}, "Image": "", "URL": source if source.startswith("http") else ""}


def export_excel(results: list[dict[str, Any]], path: pathlib.Path) -> pathlib.Path:
    """Export flat records, keeping every specification in one ``Thong so`` column."""
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

    max_cat = max((len(r.get("_cats") or []) for r in results), default=0)
    rows: list[dict[str, Any]] = []
    for result in results:
        row = {key: result.get(key, "") for key in ("Input", "Page", "Item No", "Status")}
        cats = result.get("_cats") or []
        for index in range(max_cat):
            row[f"Category {index + 1}"] = cats[index] if index < len(cats) else ""
        row.update({"Gia (PHP)": result.get("Gia"), "Stock": result.get("Stock"),
                    "Weight": result.get("Weight")})
        specs = result.get("_spec") or {}
        row["Thong so"] = "\n".join(f"{key}: {value}" if value else key for key, value in specs.items())
        row["Image"], row["URL"] = result.get("Image", ""), result.get("URL", "")
        rows.append(row)
    frame = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        frame.to_excel(writer, index=False)
        sheet = writer.sheets["Sheet1"]
        if "Thong so" in frame:
            col = get_column_letter(frame.columns.get_loc("Thong so") + 1)
            sheet.column_dimensions[col].width = 60
            for cell in sheet[col][1:]:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    return path


class Engine:
    """Async WADFOW API engine.  It is independent of the Tk GUI and resumable."""

    def __init__(self, conc: int = 6, delay: float = 0.08, log=print,
                 stop_evt: threading.Event | None = None,
                 pause_evt: threading.Event | None = None) -> None:
        self.conc, self.delay, self.log = max(1, min(int(conc), 20)), max(0.0, delay), log
        self.stop_evt, self.pause_evt = stop_evt or threading.Event(), pause_evt or threading.Event()
        self.results: list[dict[str, Any]] = []
        self.done = self.total = self.fails = 0
        self.key = ""
        self._ckpt_loaded = False
        self._write_lock = asyncio.Lock()

    def checkpoint(self) -> pathlib.Path:
        safe = re.sub(r"[^A-Za-z0-9]+", "_", self.key)[:80]
        return HERE / f"_raw_{safe}.jsonl"

    def _record_key(self, result: dict[str, Any]) -> tuple[str, str]:
        return str(result.get("Input", "")), str(result.get("Item No", ""))

    def _load_checkpoint(self) -> set[tuple[str, str]]:
        if self._ckpt_loaded:
            return {self._record_key(r) for r in self.results if r.get("Status") != "Failed"}
        self._ckpt_loaded = True
        seen: set[tuple[str, str]] = set()
        path = self.checkpoint()
        if not path.exists():
            return seen
        failed = 0
        for line in path.read_text("utf-8", errors="ignore").splitlines():
            try:
                result = json.loads(line)
            except json.JSONDecodeError:
                continue
            if result.get("Status") == "Failed":
                failed += 1
                continue
            key = self._record_key(result)
            if key not in seen:
                self.results.append(result)
                seen.add(key)
        if seen or failed:
            self.log(f"[SYS] Resume: {len(seen)} record da xong; {failed} Failed se thu lai")
        return seen

    async def _append(self, result: dict[str, Any]) -> None:
        async with self._write_lock:
            self.results.append(result)
            with self.checkpoint().open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(result, ensure_ascii=False) + "\n")
                handle.flush()
            self.done += 1

    def cleanup(self) -> None:
        path = self.checkpoint()
        failed = sum(r.get("Status") == "Failed" for r in self.results)
        if path.exists() and not self.stop_evt.is_set() and not failed:
            path.unlink()
        elif failed:
            self.log(f"[SYS] Con {failed} Failed -> giu checkpoint de chay lai")

    async def _gate(self) -> bool:
        while self.pause_evt.is_set() and not self.stop_evt.is_set():
            await asyncio.sleep(0.2)
        return not self.stop_evt.is_set()

    async def request(self, client: httpx.AsyncClient, endpoint: str,
                      params: dict[str, Any] | None = None) -> tuple[Any | None, str]:
        """Return ``(data, status)``; Failed means transport/WAF uncertainty, not a miss."""
        url = API_BASE + endpoint
        for attempt in range(5):
            if not await self._gate():
                return None, "Failed"
            try:
                response = await client.get(url, params=params)
                if response.status_code in (403, 429, 503, 520, 521, 522):
                    wait = min(2 ** attempt * 2, 30)
                    self.log(f"[WARN] {response.status_code} -> doi {wait}s | {endpoint}")
                    await asyncio.sleep(wait)
                    continue
                response.raise_for_status()
                payload = response.json()
                if str(payload.get("code")) == "0":
                    self.fails = 0
                    return payload.get("data"), "OK"
                return None, "Not Found"
            except (httpx.HTTPError, ValueError) as exc:
                if attempt < 4:
                    await asyncio.sleep(min(2 ** attempt, 20))
                    continue
                self.fails += 1
                self.log(f"[ERR] {type(exc).__name__} | {endpoint}")
        if self.fails >= CIRCUIT_BREAK:
            self.log(f"[STOP] {CIRCUIT_BREAK} loi lien tiep -- dung de bao ve IP")
            self.stop_evt.set()
        return None, "Failed"

    async def _detail(self, client: httpx.AsyncClient, code: str, source: str) -> dict[str, Any]:
        data, status = await self.request(client, "product/detail",
                                          {"productExtNo": code, "url": "product/:name"})
        if not data:
            return result_error(source, status)
        result = parse_detail(data)
        result["Input"] = source
        return result

    async def _process_items(self, client: httpx.AsyncClient, items: list[dict[str, Any]],
                             source_for, seen: set[tuple[str, str]], limit: int) -> None:
        sem = asyncio.Semaphore(self.conc)

        async def one(item: dict[str, Any]) -> None:
            if not await self._gate():
                return
            code = str(item.get("productExtNo") or item.get("productNo") or "")
            source = source_for(item)
            if not code or (source, code) in seen or (limit and len(self.results) >= limit):
                return
            async with sem:
                if not await self._gate():
                    return
                result = await self._detail(client, code, source)
                await self._append(result)
                seen.add(self._record_key(result))
                if self.delay:
                    await asyncio.sleep(self.delay)

        await asyncio.gather(*(one(item) for item in items))

    async def _search_one(self, client: httpx.AsyncClient, query: str,
                          mode: str, seen: set[tuple[str, str]]) -> None:
        code = extract_code(query)
        data, state = await self.request(client, "product", {"keyword": code, "pageNum": 1, "pageSize": 100})
        if state == "Failed":
            await self._append(result_error(query, "Failed"))
            return
        items = (data or {}).get("list") or []
        if mode in ("code", "link", "excel"):
            items = [item for item in items if is_exact_code(code, item)]
        if not items:
            # Product URLs already expose the manufacturer code; direct detail avoids blind hits.
            result = await self._detail(client, code, query)
            if self._record_key(result) not in seen:
                await self._append(result)
            return
        await self._process_items(client, items, lambda _: query, seen, 0)

    async def run(self, mode: str, target: str = "", excel: str = "", sheet: str = "",
                  col: str = "", limit: int = 0) -> list[dict[str, Any]]:
        self.key = f"{mode}_{target or excel}_{col}"
        seen = self._load_checkpoint()
        timeout = httpx.Timeout(TIMEOUT, connect=15.0)
        async with httpx.AsyncClient(headers=HEADERS, timeout=timeout, follow_redirects=True,
                                     max_redirects=MAX_REDIRECTS) as client:
            if mode == "all":
                previous: set[str] = set()
                for page in range(1, MAX_PAGES + 1):
                    if not await self._gate() or (limit and len(self.results) >= limit):
                        break
                    data, state = await self.request(client, "product", {"pageNum": page, "pageSize": 100})
                    if state == "Failed" or not data:
                        break
                    items = data.get("list") or []
                    ids = {str(x.get("productExtNo") or x.get("productNo") or "") for x in items}
                    if not items or not ids.difference(previous):
                        self.log("[SYS] Pagination dung (trang rong/trung)")
                        break
                    previous.update(ids)
                    self.total = min(MAX_ITEMS, int(data.get("total") or len(previous)))
                    room = max(0, MAX_ITEMS - len(self.results))
                    if limit:
                        room = min(room, max(0, limit - len(self.results)))
                    await self._process_items(client, items[:room], lambda _: "", seen, limit)
                    self.log(f"[OK] ALL page {page}: {len(previous)} URL / {len(self.results)} record")
                    if len(previous) >= MAX_ITEMS:
                        self.log(f"[WARN] Cham MAX_ITEMS={MAX_ITEMS}")
                        break
            else:
                if mode == "excel":
                    frame = pd.read_excel(excel, sheet_name=sheet or 0, dtype=str)
                    if col not in frame.columns:
                        raise ValueError(f"Khong thay cot '{col}'. Cot: {list(frame.columns)}")
                    inputs = [str(v).strip() for v in frame[col].dropna() if str(v).strip()]
                elif mode == "link":
                    inputs = [line.strip() for line in target.splitlines() if line.strip()]
                else:
                    inputs = [target.strip()] if target.strip() else []
                self.total = min(len(inputs), MAX_ITEMS)
                for query in inputs[:MAX_ITEMS]:
                    if not await self._gate():
                        break
                    await self._search_one(client, query, mode, seen)
        return self.results


def run_gui() -> None:
    import tkinter as tk
    from tkinter import messagebox, ttk

    root_dir = next(p for p in (HERE, *HERE.parents) if (p / "gui_common.py").exists())
    sys.path.insert(0, str(root_dir))
    from gui_common import ask_export_path, attach_excel_picker

    root = tk.Tk()
    root.title("WADFOW Crawler -- wadfow.com/ph")
    root.geometry("1080x680")
    ttk.Style().theme_use("clam")
    logs: queue.Queue[str] = queue.Queue()
    stop_evt, pause_evt = threading.Event(), threading.Event()
    state: dict[str, Any] = {"running": False, "results": [], "engine": None}

    left = ttk.Frame(root, padding=10); left.pack(side="left", fill="y")
    right = ttk.Frame(root, padding=(0, 10, 10, 10)); right.pack(side="right", fill="both", expand=True)
    mode = tk.StringVar(value="code")
    ttk.Label(left, text="Pipeline", font=("Segoe UI", 10, "bold")).pack(anchor="w")
    for value, label in [("code", "1. Theo MA"), ("keyword", "2. Theo TU KHOA"),
                         ("link", "3. Theo LINK (moi dong 1 link)"),
                         ("excel", "4. Theo cot EXCEL"), ("all", "5. Crawl ALL API")]:
        ttk.Radiobutton(left, text=label, value=value, variable=mode).pack(anchor="w")
    ttk.Label(left, text="\nMa / Tu khoa:").pack(anchor="w")
    entry = ttk.Entry(left, width=36); entry.pack(anchor="w")
    ttk.Label(left, text="Link (moi dong 1 link):").pack(anchor="w")
    link_box = tk.Text(left, width=36, height=5); link_box.pack(anchor="w")
    xls, sheet, col, outdir = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar(value="images")
    attach_excel_picker(left, xls, sheet, col, outdir)
    conc, delay, limit = tk.IntVar(value=6), tk.DoubleVar(value=0.08), tk.IntVar(value=0)
    for label, var in [("Concurrency:", conc), ("Delay (s):", delay), ("Limit (0=all):", limit)]:
        row = ttk.Frame(left); row.pack(anchor="w", pady=1)
        ttk.Label(row, text=label, width=13).pack(side="left")
        ttk.Entry(row, textvariable=var, width=8).pack(side="left")
    console = tk.Text(right, bg="#12161c", fg="#d6e2ee", insertbackground="#fff", wrap="none")
    console.pack(fill="both", expand=True)
    progress = ttk.Progressbar(right, mode="determinate"); progress.pack(fill="x", pady=4)
    stats = ttk.Label(right, text="0/0 | 0 record"); stats.pack(anchor="w")

    def drain() -> None:
        while not logs.empty():
            console.insert("end", logs.get() + "\n"); console.see("end")
        engine: Engine | None = state["engine"]
        if engine:
            progress["maximum"] = max(1, engine.total)
            progress["value"] = min(engine.done, max(1, engine.total))
            stats.config(text=f"{engine.done}/{engine.total} | {len(engine.results)} record")
        root.after(100, drain)

    def worker() -> None:
        current = mode.get()
        target = link_box.get("1.0", "end").strip() if current == "link" else entry.get().strip()
        engine = Engine(conc.get(), delay.get(), logs.put, stop_evt, pause_evt)
        state["engine"] = engine
        try:
            asyncio.run(engine.run(current, target, xls.get(), sheet.get(), col.get(), limit.get()))
            engine.cleanup()
        except BaseException as exc:
            logs.put(f"[FATAL] {type(exc).__name__}: {exc}")
        finally:
            state["results"], state["running"] = engine.results, False
            logs.put(f"[SYS] XONG -- {len(engine.results)} record. Bam EXPORT de xuat Excel.")

    def start() -> None:
        if state["running"]:
            messagebox.showinfo("Dang chay", "Dang crawl. Bam Stop truoc."); return
        current = mode.get()
        if current == "excel" and not (xls.get() and col.get()):
            messagebox.showwarning("Thieu du lieu", "Chon Excel va cot truoc."); return
        if current in ("code", "keyword") and not entry.get().strip():
            messagebox.showwarning("Thieu du lieu", "Nhap ma / tu khoa truoc."); return
        if current == "link" and not link_box.get("1.0", "end").strip():
            messagebox.showwarning("Thieu du lieu", "Dan link vao o (moi dong 1 link)."); return
        stop_evt.clear(); pause_evt.clear(); state["running"] = True; console.delete("1.0", "end")
        threading.Thread(target=worker, daemon=True).start()

    def export() -> None:
        if not state["results"]:
            messagebox.showinfo("Export", "Chua co du lieu"); return
        path = ask_export_path(xls.get(), HERE, "wadfow.xlsx")
        if path:
            export_excel(state["results"], pathlib.Path(path))
            messagebox.showinfo("Export", f"Da luu:\n{path}")

    controls = ttk.Frame(left); controls.pack(anchor="w", pady=8)
    ttk.Button(controls, text="START", command=start).pack(side="left")
    ttk.Button(controls, text="Pause", command=lambda: pause_evt.clear() if pause_evt.is_set() else pause_evt.set()).pack(side="left")
    ttk.Button(controls, text="Stop", command=stop_evt.set).pack(side="left")
    ttk.Button(controls, text="EXPORT", command=export).pack(side="left")

    def on_close() -> None:
        if state["running"] and not messagebox.askyesno("Thoat", "Dang crawl. Thoat va tat luon?"):
            return
        stop_evt.set()
        try:
            root.destroy()
        finally:
            os._exit(0)

    root.protocol("WM_DELETE_WINDOW", on_close)
    drain(); root.mainloop()


def main() -> None:
    parser = argparse.ArgumentParser(description="WADFOW Philippines API crawler")
    parser.add_argument("--mode", choices=["code", "keyword", "link", "excel", "all"])
    parser.add_argument("--code"); parser.add_argument("--keyword"); parser.add_argument("--link")
    parser.add_argument("--excel"); parser.add_argument("--sheet", default=""); parser.add_argument("--col", default="")
    parser.add_argument("--out", default="wadfow.xlsx"); parser.add_argument("--conc", type=int, default=6)
    parser.add_argument("--delay", type=float, default=0.08); parser.add_argument("--limit", type=int, default=0)
    args = parser.parse_args()
    if not args.mode:
        run_gui(); return
    target = args.code or args.keyword or args.link or ""
    engine = Engine(args.conc, args.delay)
    started = time.time()
    try:
        asyncio.run(engine.run(args.mode, target, args.excel or "", args.sheet, args.col, args.limit))
    finally:
        if engine.results:
            path = export_excel(engine.results, pathlib.Path(args.out))
            print(f"[SYS] {len(engine.results)} record -> {path} ({time.time() - started:.1f}s)")
        engine.cleanup()


if __name__ == "__main__":
    main()
