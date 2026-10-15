"""
thietbidungcukythuat.com - WooCommerce Store API crawler.

RECON (2026-08-11)
------------------
- Nen tang: WordPress + WooCommerce.
- Store API: /wp-json/wc/store/v1/products?sku=<code> cho exact code, ?search=<kw> cho keyword.
- Total catalog: 5,502 san pham, phan trang per_page=100.
- Link mode: dung WP REST /wp-json/wp/v2/product?slug=<slug>&_embed=1 de resolve URL that.
- Search web ?s=... tra 410 va robots chan -> khong dung HTML search.
- Spec va mo ta nam trong description HTML cua Store API.
"""
from __future__ import annotations

import argparse
import asyncio
import html
import json
import os
import pathlib
import queue
import re
import sys
import threading
import time
from typing import Any
from urllib.parse import quote

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser

HERE = pathlib.Path(__file__).parent
BASE = "https://thietbidungcukythuat.com"
STORE_API = BASE + "/wp-json/wc/store/v1/products"
WP_API = BASE + "/wp-json/wp/v2/product"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
HEADERS = {"User-Agent": UA, "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8"}

CIRCUIT_BREAK = 25
MAX_ITEMS = 10000
MAX_PAGES = 100
MAX_REDIRECTS = 5
TIMEOUT = 45.0


def clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", html.unescape(str(value or "")).replace("\xa0", " ")).strip()


def norm_code(value: Any) -> str:
    v = re.sub(r"\s+", "", str(value or ""))
    return v if v.lower().startswith("http") else v.upper()


def slug_from_url(url: str) -> str:
    return str(url or "").strip().rstrip("/").rsplit("/", 1)[-1]


def parse_description(desc_html: str) -> tuple[dict[str, str], str]:
    tree = LexborHTMLParser(desc_html or "")
    spec: dict[str, str] = {}
    prose: list[str] = []
    for li in tree.css("li"):
        text = clean_text(li.text(separator=" ", strip=True))
        if not text:
            continue
        if ":" in text:
            key, val = text.split(":", 1)
            key = clean_text(key).rstrip(":")
            val = clean_text(val)
            if key and val and key not in spec:
                spec[key] = val
                continue
        prose.append(text)
    for p in tree.css("p"):
        text = clean_text(p.text(separator=" ", strip=True))
        if text and ":" not in text:
            prose.append(text)
    prose = [line for line in dict.fromkeys(prose) if line.lower() not in {"thong so ky thuat:"}]
    return spec, "\n".join(prose)


def parse_store_product(product: dict, input_value: str = "") -> dict:
    spec, desc = parse_description(product.get("description") or "")
    images = [clean_text(img.get("src")) for img in (product.get("images") or []) if clean_text(img.get("src"))]
    cats = [clean_text(cat.get("name")) for cat in (product.get("categories") or []) if clean_text(cat.get("name"))]
    return {
        "Input": input_value,
        "Item No": clean_text(product.get("sku")),
        "Page": clean_text(product.get("name")),
        "Status": "OK",
        "Gia (VND)": int(product.get("prices", {}).get("price") or 0) or "",
        "Ton kho": "Co" if product.get("is_in_stock") else "Het",
        "_cats": cats,
        "_spec": spec,
        "Mo ta": desc,
        "Image": " | ".join(dict.fromkeys(images)),
        "URL": clean_text(product.get("permalink")),
    }


def exact_code(code: str, row: dict) -> bool:
    c = norm_code(code)
    hay = " ".join(norm_code(row.get(k)) for k in ("Item No", "Page", "URL"))
    return bool(re.search(rf"(?<![A-Z0-9]){re.escape(c)}(?![A-Z0-9])", hay))


def not_found(input_value: str, status: str = "Not Found") -> dict:
    return {
        "Input": input_value,
        "Item No": input_value,
        "Page": "",
        "Status": status,
        "Gia (VND)": "",
        "Ton kho": "",
        "_cats": [],
        "_spec": {},
        "Mo ta": "",
        "Image": "",
        "URL": "",
    }


def export_excel(rows: list[dict], path: pathlib.Path) -> pathlib.Path:
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

    fixed_head = ["Input", "Page", "Item No", "Status", "Gia (VND)", "Ton kho"]
    max_cat = max((len(row.get("_cats") or []) for row in rows), default=0)
    out = []
    for row in rows:
        item = {k: row.get(k, "") for k in fixed_head}
        cats = row.get("_cats") or []
        for i in range(max_cat):
            item[f"Category {i + 1}"] = cats[i] if i < len(cats) else ""
        item["Thong so"] = "\n".join(f"{k}: {v}" for k, v in (row.get("_spec") or {}).items())
        item["Mo ta"] = row.get("Mo ta", "")
        item["Image"] = row.get("Image", "")
        item["URL"] = row.get("URL", "")
        out.append(item)
    df = pd.DataFrame(out)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
        ws = writer.sheets["Sheet1"]
        align = Alignment(wrap_text=True, vertical="top")
        for name, width in [("Thong so", 60), ("Mo ta", 80)]:
            if name not in df.columns:
                continue
            col = get_column_letter(df.columns.get_loc(name) + 1)
            ws.column_dimensions[col].width = width
            for cell in ws[col][1:]:
                cell.alignment = align
    return path


def read_excel_col(path: str, sheet: str, col: str) -> list[str]:
    if not path:
        raise ValueError("Chua chon file Excel")
    df = pd.read_excel(path, sheet_name=sheet or 0, dtype=str)
    df.columns = df.columns.map(str)
    if col not in df.columns:
        raise ValueError(f"Khong thay cot '{col}'")
    return [norm_code(v) for v in df[col].dropna().tolist()]


class Engine:
    def __init__(
        self,
        conc: int = 8,
        delay: float = 0.05,
        log=print,
        stop_evt: threading.Event | None = None,
        pause_evt: threading.Event | None = None,
    ) -> None:
        self.conc = max(1, conc)
        self.delay = max(0.0, delay)
        self.log = log
        self.stop_evt = stop_evt or threading.Event()
        self.pause_evt = pause_evt or threading.Event()
        self.results: list[dict] = []
        self.done = 0
        self.total = 0
        self.fails = 0
        self.key = ""
        self._ckpt_loaded = False

    def ckpt(self, key: str) -> pathlib.Path:
        safe = re.sub(r"[^A-Za-z0-9]+", "_", key)[:50]
        return HERE / f"_raw_{safe}.jsonl"

    def _load_ckpt(self, key: str) -> set[str]:
        if self._ckpt_loaded:
            return {row.get("URL", "") for row in self.results if row.get("URL")}
        self._ckpt_loaded = True
        seen: set[str] = set()
        p = self.ckpt(key)
        if not p.exists():
            return seen
        for line in p.read_text("utf-8", errors="ignore").splitlines():
            try:
                row = json.loads(line)
            except Exception:
                continue
            if row.get("Status") != "OK":
                continue
            self.results.append(row)
            if row.get("URL"):
                seen.add(row["URL"])
        return seen

    def _save(self, key: str, rows: list[dict]) -> None:
        with self.ckpt(key).open("a", encoding="utf-8") as fh:
            for row in rows:
                fh.write(json.dumps(row, ensure_ascii=False) + "\n")
            fh.flush()

    def cleanup(self) -> None:
        p = self.ckpt(self.key) if self.key else None
        if p and p.exists() and not self.stop_evt.is_set():
            p.unlink()

    async def _gate(self) -> bool:
        while self.pause_evt.is_set() and not self.stop_evt.is_set():
            await asyncio.sleep(0.2)
        return not self.stop_evt.is_set()

    async def fetch_json(self, client: httpx.AsyncClient, url: str, params: dict[str, Any] | None = None) -> Any:
        for i in range(4):
            if not await self._gate():
                return None
            try:
                resp = await client.get(url, params=params)
                if resp.status_code in (403, 429, 503):
                    wait = min(5 * (2 ** i), 60)
                    self.log(f"[WARN] {resp.status_code} -> ngu {wait}s | {resp.url}")
                    await asyncio.sleep(wait)
                    continue
                if resp.status_code == 404:
                    return None
                resp.raise_for_status()
                self.fails = 0
                return resp.json()
            except Exception as exc:
                if i < 3:
                    await asyncio.sleep(min(2 ** i, 20))
                    continue
                self.fails += 1
                self.log(f"[ERR] {type(exc).__name__} | {url}")
                if self.fails >= CIRCUIT_BREAK:
                    self.log(f"[STOP] {CIRCUIT_BREAK} loi lien tiep - dung.")
                    self.stop_evt.set()
                return None
        return None

    async def resolve_link(self, client: httpx.AsyncClient, url: str) -> dict | None:
        slug = slug_from_url(url)
        data = await self.fetch_json(client, WP_API, {"slug": slug, "per_page": 1})
        if not data:
            return None
        item = data[0]
        store = await self.fetch_json(client, STORE_API, {"slug": item.get("slug"), "per_page": 1})
        if store:
            return store[0]
        return None

    async def _run_code(self, client: httpx.AsyncClient, codes: list[str]) -> None:
        done_inputs = {norm_code(row.get("Input")) for row in self.results if row.get("Input")}
        todo = [code for code in codes if norm_code(code) not in done_inputs][:MAX_ITEMS]
        self.total += len(todo)
        sem = asyncio.Semaphore(self.conc)

        async def one(code: str) -> None:
            async with sem:
                if not await self._gate():
                    return  # Bi Stop truoc khi kiem tra -> bo qua, lan sau se retry
                data = await self.fetch_json(client, STORE_API, {"sku": code, "per_page": 100})
                rows = [parse_store_product(p, code) for p in (data or []) if exact_code(code, parse_store_product(p, code))]
                if rows:
                    self.results.extend(rows)
                    self._save(self.key, rows)
                else:
                    # Stop giua chung -> code nay chua duoc kiem tra that, khong duoc bao Not Found gia
                    status = "Unverified - catalog scan interrupted" if self.stop_evt.is_set() else "Not Found"
                    row = not_found(code, status)
                    self.results.append(row)
                    self._save(self.key, [row])
                self.done += 1
                self.log(f"[{'OK ' if rows else 'MISS'}] {self.done}/{self.total} | {code}")

        await asyncio.gather(*(one(code) for code in todo))

    async def _run_keyword(self, client: httpx.AsyncClient, keyword: str, limit: int) -> None:
        data = await self.fetch_json(client, STORE_API, {"search": keyword, "per_page": min(limit or 100, 100)})
        rows = [parse_store_product(p, keyword) for p in (data or [])]
        if limit:
            rows = rows[:limit]
        self.results.extend(rows)
        if rows:
            self._save(self.key, rows)
        self.total = self.done = len(rows)

    async def _run_links(self, client: httpx.AsyncClient, links: list[str]) -> None:
        self.total += len(links)
        for link in links:
            if not await self._gate():
                break  # Bi Stop -> cac link con lai chua kiem tra, khong bao Not Found gia, de lan sau retry
            product = await self.resolve_link(client, link)
            if product:
                rows = [parse_store_product(product, link)]
            else:
                # Stop giua chung -> link nay chua duoc kiem tra that, khong duoc bao Not Found gia
                status = "Unverified - catalog scan interrupted" if self.stop_evt.is_set() else "Not Found"
                rows = [not_found(link, status)]
            self.results.extend(rows)
            self._save(self.key, rows)
            self.done += 1
            self.log(f"[{'OK ' if product else 'MISS'}] {self.done}/{self.total} | {link[-60:]}")

    async def _run_all(self, client: httpx.AsyncClient, limit: int, seen_urls: set[str]) -> None:
        page = 1
        seen_ids: set[int] = set()
        while page <= MAX_PAGES:
            data = await self.fetch_json(client, STORE_API, {"page": page, "per_page": 100})
            items = data or []
            if not items:
                break
            new = [p for p in items if p.get("id") not in seen_ids]
            if not new:
                break
            for item in new:
                if item.get("id") is not None:
                    seen_ids.add(item["id"])
            # Bo qua san pham da co trong checkpoint (resume) de khong nhan ban
            fresh = [p for p in new if clean_text(p.get("permalink")) not in seen_urls]
            rows = [parse_store_product(p, "") for p in fresh]
            self.results.extend(rows)
            self._save(self.key, rows)
            self.done += len(rows)
            self.total += len(rows)
            if limit and len(self.results) >= limit:
                return
            page += 1
            if self.delay:
                await asyncio.sleep(self.delay)

    async def run(
        self,
        mode: str,
        target: str = "",
        excel: str = "",
        sheet: str = "",
        col: str = "",
        limit: int = 0,
    ) -> list[dict]:
        self.key = f"{mode}_{target or excel}_{col}"
        seen_urls = self._load_ckpt(self.key)
        async with httpx.AsyncClient(
            headers=HEADERS,
            timeout=TIMEOUT,
            follow_redirects=True,
            max_redirects=MAX_REDIRECTS,
            limits=httpx.Limits(max_connections=self.conc + 4),
        ) as client:
            if mode == "all":
                await self._run_all(client, limit, seen_urls)
            elif mode in {"code", "excel"}:
                raw = [norm_code(target)] if mode == "code" else read_excel_col(excel, sheet, col)
                links = [v for v in raw if v.lower().startswith("http")]
                codes = [v for v in raw if v not in links]
                if links:
                    await self._run_links(client, links)
                if codes:
                    await self._run_code(client, codes)
            elif mode == "keyword":
                await self._run_keyword(client, target, limit)
            elif mode == "link":
                links = [line.strip() for line in target.splitlines() if line.strip()]
                await self._run_links(client, links)
            else:
                raise ValueError(f"Mode khong hop le: {mode}")
        return self.results


def run_gui() -> None:
    import tkinter as tk
    from tkinter import messagebox, ttk

    root_dir = next(p for p in (HERE, *HERE.parents) if (p / "gui_common.py").exists())
    sys.path.insert(0, str(root_dir))
    from gui_common import ask_export_path, attach_excel_picker

    root = tk.Tk()
    root.title("Thiet Bi Dung Cu Ky Thuat Crawler")
    root.geometry("1080x700")
    ttk.Style().theme_use("clam")

    q: queue.Queue[str] = queue.Queue()
    stop_evt = threading.Event()
    pause_evt = threading.Event()
    state: dict[str, Any] = {"running": False, "rows": [], "eng": None}

    left = ttk.Frame(root, padding=10)
    left.pack(side="left", fill="y")
    right = ttk.Frame(root, padding=(0, 10, 10, 10))
    right.pack(side="right", fill="both", expand=True)

    mode = tk.StringVar(value="code")
    ttk.Label(left, text="Pipeline", font=("Segoe UI", 10, "bold")).pack(anchor="w")
    for val, label in [
        ("code", "1. Theo MA"),
        ("keyword", "2. Theo TU KHOA"),
        ("link", "3. Theo LINK"),
        ("excel", "4. Theo EXCEL"),
        ("all", "5. Crawl ALL"),
    ]:
        ttk.Radiobutton(left, text=label, value=val, variable=mode).pack(anchor="w")

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
    for label, var in [("Concurrency:", conc), ("Delay (s):", delay), ("Limit (0=all):", limit)]:
        fr = ttk.Frame(left)
        fr.pack(anchor="w", pady=1)
        ttk.Label(fr, text=label, width=13).pack(side="left")
        ttk.Entry(fr, textvariable=var, width=8).pack(side="left")

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
            pb["maximum"] = eng.total
            pb["value"] = eng.done
            stat.config(text=f"{eng.done}/{eng.total} | {len(eng.results)} dong")
        root.after(100, drain)

    def worker() -> None:
        target = box.get("1.0", "end").strip() if mode.get() == "link" else ent.get().strip()
        eng = Engine(conc.get(), delay.get(), q.put, stop_evt, pause_evt)
        state["eng"] = eng
        try:
            state["rows"] = asyncio.run(eng.run(mode.get(), target, xls.get(), sheet.get(), col.get(), limit.get()))
            eng.cleanup()
        except BaseException as exc:
            q.put(f"[FATAL] {type(exc).__name__}: {exc}")
        finally:
            state["running"] = False
            q.put(f"[SYS] XONG - {len(state['rows'])} dong. Bam EXPORT de xuat Excel.")

    def start() -> None:
        if state["running"]:
            return
        if mode.get() in {"code", "keyword"} and not ent.get().strip():
            messagebox.showwarning("Thieu du lieu", "Nhap ma hoac tu khoa truoc.")
            return
        if mode.get() == "link" and not box.get("1.0", "end").strip():
            messagebox.showwarning("Thieu du lieu", "Dan link vao o moi dong 1 link.")
            return
        if mode.get() == "excel" and not (xls.get() and col.get()):
            messagebox.showwarning("Thieu du lieu", "Chon file Excel va cot.")
            return
        stop_evt.clear()
        pause_evt.clear()
        state["running"] = True
        console.delete("1.0", "end")
        threading.Thread(target=worker, daemon=True).start()

    def export() -> None:
        if not state["rows"]:
            messagebox.showinfo("Export", "Chua co du lieu")
            return
        path = ask_export_path(xls.get(), HERE, "thietbidungcukythuat.xlsx")
        if path:
            export_excel(state["rows"], pathlib.Path(path))
            messagebox.showinfo("Export", f"Da luu:\n{path}")

    fr_btn = ttk.Frame(left)
    fr_btn.pack(anchor="w", pady=8)
    ttk.Button(fr_btn, text="START", command=start).pack(side="left")
    ttk.Button(fr_btn, text="Pause", command=lambda: (pause_evt.clear() if pause_evt.is_set() else pause_evt.set())).pack(side="left")
    ttk.Button(fr_btn, text="Stop", command=stop_evt.set).pack(side="left")
    ttk.Button(fr_btn, text="EXPORT", command=export).pack(side="left")

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


def main() -> None:
    ap = argparse.ArgumentParser(description="Crawler thietbidungcukythuat.com")
    ap.add_argument("--mode", choices=["code", "keyword", "link", "excel", "all"])
    ap.add_argument("--code")
    ap.add_argument("--keyword")
    ap.add_argument("--link")
    ap.add_argument("--excel")
    ap.add_argument("--sheet", default="")
    ap.add_argument("--col", default="")
    ap.add_argument("--out", default="thietbidungcukythuat.xlsx")
    ap.add_argument("--conc", type=int, default=8)
    ap.add_argument("--delay", type=float, default=0.05)
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    if not args.mode:
        run_gui()
        return

    target = args.code or args.keyword or args.link or ""
    engine = Engine(args.conc, args.delay)
    t0 = time.time()
    try:
        rows = asyncio.run(engine.run(args.mode, target, args.excel or "", args.sheet, args.col, args.limit))
    finally:
        if engine.results:
            export_excel(engine.results, pathlib.Path(args.out))
        engine.cleanup()
    print(f"[SYS] {len(rows)} dong -> {args.out} ({time.time() - t0:.1f}s)")


if __name__ == "__main__":
    main()
