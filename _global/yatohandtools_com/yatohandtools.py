"""
yatohandtools.com — YATO Tools (nha may TQ, catalog tieng Anh). httpx + selectolax, KHONG browser.

RECON (2026-07-16)
------------------
API      : KHONG co JSON API. Server-side render => httpx + selectolax la du.
robots   : Disallow /adminsoft/ /wap/ /install/ /public/ /html/ /api/ /htaccess/.
           Trang san pham + sitemap KHONG bi cam.
sitemap  : /sitemap.xml = sitemapindex -> categories.xml | **plist.xml** (1.486 SP) | news.xml
detail   : /<Category>/<slug>       vd /Automotive-tools/yato-yt-05800-cylinder-hone-tool-set-...
           - h1                                 -> ten SP (VIET HOA het)
           - table[0]  -> THONG SO, dang cap doi: [Nhan:, Gia tri, Nhan:, Gia tri] moi <tr>
             (Grade / Warranty / Certification / Place of Origin / Brand Name / Model Number /
              Dimensions(cm) / Weight / Product name / Application / MOQ / OEM / GROUP ...)
             `Model Number` = MA SAN PHAM (vd YT-05800).
           - table[1] -> dong tom tat (ma, brand, ten, group) - bo qua, trung table[0]
           - anh: **.swiper-slide img** (anh SP that, CDN cdn-pro.analyspeed.com/617/upload/product/)
             CANH BAO: trang con co anh /upload/product/20210522/pic_*.jpg = BANNER mo ta chung
             (Product Category), KHONG phai anh SP -> phai loai.
           - category: lay tu path URL (/Automotive-tools/...)
search   : /search?keyword=<ma> tra HTTP 200 nhung **0 san pham** (ca ma that lan ma bay).
           => Search HONG, khong dung duoc. Tra ma qua CODE-INDEX tu slug plist.xml
             (ma nam trong slug: yato-yt-0143-..., yato-yt-05800-...).
gia      : KHONG co gia (site B2B nha may, chi MOQ + hoi hang).

Run
---
  yatohandtools.bat                                  -> GUI
  python yatohandtools.py --mode code --code YT-05800
  python yatohandtools.py --mode excel --excel parts.xlsx --col Part
  python yatohandtools.py --mode link --link https://www.yatohandtools.com/Spanners/yato-yt-0143-...
  python yatohandtools.py --mode all [--limit 200]   -> quet sach ~1.5k SP
  python yatohandtools.py --mode images --excel parts.xlsx --col Part --outdir out
  python yatohandtools.py --build-index              -> dung code-index (~2s)
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
from urllib.parse import urljoin, urlparse

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser

HERE = pathlib.Path(__file__).parent
BASE = "https://www.yatohandtools.com"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")

CODE_INDEX = HERE / "_code_index.json"    # ma -> URL (search cua site hong, xem docstring)
INDEX_TTL_DAYS = 30

CIRCUIT_BREAK = 25        # N loi LIEN TIEP -> nghi bi chan IP -> dung han
MAX_ITEMS = 5000          # site co ~1.5k SP
MAX_REDIRECTS = 5
TIMEOUT = 30.0
GC_EVERY = 200

PROD_RE = re.compile(r"https://www\.yatohandtools\.com/[A-Za-z][\w-]*/[a-z0-9][a-z0-9-]+")
BANNER_RE = re.compile(r"/upload/product/\d{8}/pic_")   # banner mo ta chung, KHONG phai anh SP


# ================================================================ CORE (khong dung GUI)
def norm_code(s: Any) -> str:
    """Chuan hoa ma de so khop. URL thi GIU NGUYEN hoa/thuong (path phan biet chu)."""
    v = re.sub(r"\s+", "", str(s or ""))
    return v if v.lower().startswith("http") else v.upper()


def codes_from_slug(url: str) -> list[str]:
    """Ma trong slug: 'yato-yt-0143-flare-nut...' -> YT-0143; 'auto-repair-...-yato-12681' -> YT-12681.

    CHI la doan mo: 701/1446 trang KHONG co ma trong slug (vd 'yato-combination-spanner-9mm'
    -> ma that YT-0009 chi biet khi doc `Model Number` trong bang) => build_index PHAI crawl trang.
    """
    slug = urlparse(url).path.rsplit("/", 1)[-1]
    out = [f"{m.group(1).upper()}-{m.group(2).upper()}"
           for m in re.finditer(r"(?:^|-)(yt|ytm|yg)-?(\d[0-9a-z]*)", slug, re.I)]
    out += [f"YT-{m.group(1).upper()}"                 # 'yato-12681' (khong co prefix yt-)
            for m in re.finditer(r"(?:^|-)yato-(\d[0-9a-z]*)", slug, re.I)]
    return list(dict.fromkeys(out))


def parse_links(html: str) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for u in PROD_RE.findall(html):
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def parse_sitemap(xml: str) -> list[str]:
    return re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", xml)


def parse_spec(t: LexborHTMLParser) -> dict[str, str]:
    """table[0]: moi <tr> la CAP DOI [Nhan:, Gia tri, Nhan:, Gia tri] -> tach tung cap."""
    spec: dict[str, str] = {}
    tables = t.css("table")
    if not tables:
        return spec
    for tr in tables[0].css("tr"):
        cells = [re.sub(r"\s+", " ", c.text(strip=True)) for c in tr.css("td, th")]
        for i in range(0, len(cells) - 1, 2):
            k, v = cells[i].rstrip(":").strip(), cells[i + 1].strip()
            if k and v:
                spec.setdefault(k, v)
    return spec


def parse_page(html: str, url: str) -> dict | None:
    t = LexborHTMLParser(html)
    h1 = t.css_first("h1")
    if h1 is None:
        return None
    name = re.sub(r"\s+", " ", h1.text(strip=True))

    spec = parse_spec(t)
    code = spec.get("Model Number", "") or (codes_from_slug(url) or [""])[0]

    images: list[str] = []
    for im in t.css(".swiper-slide img"):      # anh SP that; banner pic_* bi loai o duoi
        u = im.attributes.get("src") or im.attributes.get("data-src") or ""
        if u and "/upload/product/" in u and not BANNER_RE.search(u):
            u = urljoin(BASE, u)
            if u not in images:
                images.append(u)

    parts = [p for p in urlparse(url).path.split("/") if p]
    cats = [parts[0].replace("-", " ")] if len(parts) > 1 else []

    return {"Input": "", "Ma": code, "Ten": name,
            "Thuong hieu": spec.get("Brand Name", ""), "Xuat xu": spec.get("Place of Origin", ""),
            "Status": "OK", "_cats": cats, "_spec": spec,
            "Image": " | ".join(images[:5]), "URL": url}


def export_excel(results: list[dict], path: pathlib.Path) -> pathlib.Path:
    """Spec gop 1 COT 'Thong so' (moi dong 'Ten: Gia tri'). Xem AGENTS.md 4.6."""
    from openpyxl.utils import get_column_letter

    fixed = ("Input", "Ma", "Ten", "Thuong hieu", "Xuat xu", "Status")
    max_cat = max((len(r.get("_cats") or []) for r in results), default=0)
    rows = []
    for r in results:
        row = {k: r.get(k, "") for k in fixed}
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
    def __init__(self, conc: int = 8, delay: float = 0.05, log=print,
                 stop_evt: threading.Event | None = None,
                 pause_evt: threading.Event | None = None) -> None:
        self.conc, self.delay, self.log = max(1, conc), max(0.0, delay), log
        self.stop_evt = stop_evt or threading.Event()
        self.pause_evt = pause_evt or threading.Event()
        self.fails = 0
        self.results: list[dict] = []
        self.seen_urls: set[str] = set()
        self.done = 0
        self.total = 0
        self.key = ""
        self._ckpt_loaded = False
        self._index: dict[str, str] | None = None
        self._index_built_at: float | None = None   # thoi diem build index (tuoi cua DU LIEU, khong phai mtime)
        self._index_stale = False                    # index vua load co qua han TTL khong

    # ---- resume: JSONL sidecar
    def ckpt(self, key: str) -> pathlib.Path:
        return HERE / f"_raw_{re.sub(r'[^A-Za-z0-9]+', '_', key)[:40]}.jsonl"

    def _load_ckpt(self, key: str) -> set[str]:
        p, seen = self.ckpt(key), set()
        if self._ckpt_loaded:
            return set(self.seen_urls)
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
            self.seen_urls.add(r.get("URL", ""))
        if seen:
            self.log(f"[SYS] Resume: {len(self.results)} record da co -> bo qua")
        return seen

    def add(self, recs: list[dict], key: str) -> None:
        fresh = [r for r in recs if r.get("URL") and r["URL"] not in self.seen_urls]
        for r in fresh:
            self.seen_urls.add(r["URL"])
        if fresh:
            self.results += fresh
            with self.ckpt(key).open("a", encoding="utf-8") as f:
                for r in fresh:
                    f.write(json.dumps(r, ensure_ascii=False) + "\n")
                f.flush()

    def cleanup(self) -> None:
        p = self.ckpt(self.key) if self.key else None
        if p and p.exists() and not self.stop_evt.is_set():
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
                if r.status_code in (403, 429, 503):
                    wait = min(2 ** i * 5, 60)
                    self.log(f"[WARN] {r.status_code} -> ngu {wait}s | {url[-55:]}")
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
                self.log(f"[ERR] {type(e).__name__} | {url[-55:]}")
                if self.fails >= CIRCUIT_BREAK:
                    self.log(f"[STOP] {CIRCUIT_BREAK} loi LIEN TIEP - nghi bi chan IP. DUNG.")
                    self.stop_evt.set()
                return None
        return None

    # ---- CODE-INDEX: search cua site tra 0 KQ -> ma -> URL lay tu SLUG trong plist.xml
    def _load_index(self) -> dict[str, str]:
        if self._index is not None:
            return self._index
        self._index = {}
        self._index_stale = True
        if CODE_INDEX.exists():
            try:
                raw = json.loads(CODE_INDEX.read_text("utf-8"))
                built_at = raw.get("built_at") if isinstance(raw, dict) else None
                if built_at is None:
                    pass    # dinh dang cu (bare mapping ma->URL, khong co built_at) -> coi la stale
                else:
                    self._index_built_at = built_at
                    if (time.time() - built_at) / 86400 < INDEX_TTL_DAYS:
                        self._index = raw.get("codes", {})
                        self._index_stale = False
                        self.log(f"[SYS] Code-index: {len(self._index)} ma da biet")
            except Exception:
                pass
        return self._index

    def save_index(self, full_rebuild: bool = False) -> None:
        if not self._index:
            return
        if full_rebuild or self._index_built_at is None:
            self._index_built_at = time.time()    # du lieu moi -> reset tuoi
        CODE_INDEX.write_text(
            json.dumps({"built_at": self._index_built_at, "codes": self._index}, ensure_ascii=False), "utf-8")

    async def product_urls(self, cli: httpx.AsyncClient) -> list[str]:
        """sitemap.xml -> plist.xml -> ~1.486 URL san pham."""
        idx_xml = await self.fetch(cli, f"{BASE}/sitemap.xml")
        if not idx_xml:
            return []
        sms = [u for u in parse_sitemap(idx_xml) if "plist" in u]
        urls: list[str] = []
        for sm in sms:
            xml = await self.fetch(cli, sm)
            if xml:
                urls += [u for u in parse_sitemap(xml) if PROD_RE.match(u)]
        return list(dict.fromkeys(urls))

    async def build_index(self, cli: httpx.AsyncClient) -> dict[str, str]:
        """CRAWL that tung trang de lay `Model Number` trong bang.

        KHONG the chi doc slug: do duoc 701/1446 trang (48%) khong co ma trong slug
        (vd 'yato-combination-spanners--sets-combination-spanner-9mm' -> ma that YT-0009).
        Trang re (~0.15s) nen quet ca site ~200s, lam 1 lan, dung lai 30 ngay.
        """
        urls = await self.product_urls(cli)
        idx = self._load_index()
        self.log(f"[SYS] Dung code-index: quet {len(urls)} trang (~{len(urls) * 0.15:.0f}s)...")
        sem = asyncio.Semaphore(self.conc)
        done = 0

        async def one(u: str) -> None:
            nonlocal done
            async with sem:
                if not await self._gate():
                    return
                html = await self.fetch(cli, u)
                rec = parse_page(html, u) if html else None
                if rec and rec["Ma"]:
                    idx.setdefault(norm_code(rec["Ma"]), u)
                else:                                  # fallback: doan tu slug
                    for c in codes_from_slug(u):
                        idx.setdefault(c, u)
                done += 1
                if done % 200 == 0:
                    self.log(f"[SYS] index: {done}/{len(urls)} trang | {len(idx)} ma")

        await asyncio.gather(*(one(u) for u in urls))
        self._index_stale = False    # vua rebuild xong trong lan chay nay -> khong rebuild lai nua
        self.save_index(full_rebuild=True)
        self.log(f"[SYS] Code-index: {len(urls)} URL -> {len(idx)} ma (tra ma tuc thi tu gio)")
        return idx

    async def _resolve_codes(self, cli: httpx.AsyncClient, codes: list[str]) -> tuple[list[str], dict]:
        """Ma -> URL qua code-index (search cua site hong, tra 0 KQ)."""
        idx = self._load_index()
        # Dung index khi CHUA co HOAC da qua han TTL (~200s). Da co ma van thieu -> ma do that su
        # khong ton tai tren site, dung lai quet ca site them lan nua (chi 1 lan / lan chay).
        if (not idx or self._index_stale) and codes:
            idx = await self.build_index(cli)
        urls: list[str] = []
        inputs: dict[str, str] = {}
        miss: list[str] = []
        for c in codes:
            u = idx.get(c)
            if u:
                urls.append(u)
                inputs[u] = c
            else:
                miss.append(c)
        if miss:
            self.log(f"[MISS] {len(miss)} ma khong co trong sitemap: {', '.join(miss[:5])}"
                     f"{'...' if len(miss) > 5 else ''}")
        return urls, inputs

    async def _crawl_urls(self, cli: httpx.AsyncClient, urls: list[str], key: str,
                          inputs: dict[str, str] | None = None) -> None:
        seen = self._load_ckpt(key)
        urls = [u for u in dict.fromkeys(urls) if u not in seen][:MAX_ITEMS]
        self.total, self.done = len(urls) + len(seen), len(seen)
        sem = asyncio.Semaphore(self.conc)

        async def one(u: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                html = await self.fetch(cli, u)
                await asyncio.sleep(self.delay)
                rec = parse_page(html, u) if html else None
                if rec:
                    rec["Input"] = (inputs or {}).get(u, "") or rec["Ma"]
                    self.add([rec], key)
                self.done += 1
                self.log(f"[{'OK ' if rec else 'MISS'}] {self.done}/{self.total} | {u[-52:]}")
                if self.done % GC_EVERY == 0:
                    gc.collect()

        try:
            await asyncio.gather(*(one(u) for u in urls))
        except (KeyboardInterrupt, Exception) as e:
            self.log(f"[FATAL] {type(e).__name__}: {e}")

    async def run(self, mode: str, target: str = "", excel: str = "", sheet: str = "",
                  col: str = "", limit: int = 0, outdir: str = "out") -> list[dict]:
        key = self.key = f"{mode}_{target or excel}_{col}"
        headers = {"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"}
        async with httpx.AsyncClient(headers=headers, timeout=TIMEOUT, follow_redirects=True,
                                     max_redirects=MAX_REDIRECTS,
                                     limits=httpx.Limits(max_connections=self.conc + 4)) as cli:
            if mode == "all":
                urls = (await self.product_urls(cli))[:limit or MAX_ITEMS]
                self.log(f"[SYS] Crawl ALL: {len(urls)} trang san pham")
                await self._crawl_urls(cli, urls, key)

            elif mode in ("code", "excel"):
                raw = [norm_code(target)] if mode == "code" else read_excel_col(excel, sheet, col)
                raw = [c for c in dict.fromkeys(raw) if c][:MAX_ITEMS]
                links = [c for c in raw if c.lower().startswith("http")]
                codes = [c for c in raw if c not in links]
                if links:
                    await self._crawl_urls(cli, links, key, inputs={u: u for u in links})
                if codes:
                    urls, inputs = await self._resolve_codes(cli, codes)
                    await self._crawl_urls(cli, urls, key, inputs)
                have = {norm_code(v) for r in self.results
                        for v in (r.get("Input"), r.get("Ma")) if v}
                for c in codes:                       # ma miss van co dong Not Found
                    if c not in have:
                        self.results.append({"Input": c, "Ma": c, "Ten": "", "Thuong hieu": "",
                                             "Xuat xu": "", "Status": "Not Found", "_cats": [],
                                             "_spec": {}, "Image": "", "URL": ""})

            elif mode == "keyword":
                self.log("[ERR] Search cua yatohandtools.com HONG (tra 200 nhung 0 SP). "
                         "Dung mode 'code'/'excel' (tra qua sitemap) hoac 'all'.")

            elif mode == "link":
                await self._crawl_urls(
                    cli, [u.strip() for u in target.splitlines() if u.strip().startswith("http")], key)

            elif mode == "images":
                await self._download_images(cli, excel, sheet, col, outdir)

            else:
                self.log(f"[ERR] mode khong hop le: {mode}")
        self.save_index()
        return self.results

    async def _download_images(self, cli: httpx.AsyncClient, excel: str, sheet: str,
                               col: str, outdir: str) -> None:
        root = pathlib.Path(outdir) if os.path.isabs(outdir) else HERE / outdir
        root.mkdir(parents=True, exist_ok=True)
        keys = read_excel_col(excel, sheet, col)[:MAX_ITEMS]
        self.total = len(keys)
        _, inputs = await self._resolve_codes(
            cli, [k for k in keys if not k.lower().startswith("http")])
        by_code = {v: u for u, v in inputs.items()}
        sem = asyncio.Semaphore(self.conc)

        async def one(k: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                url = k if k.lower().startswith("http") else by_code.get(k)
                rec = None
                if url:
                    html = await self.fetch(cli, url)
                    rec = parse_page(html, url) if html else None
                self.done += 1
                if not rec or not rec["Image"]:
                    self.log(f"[MISS] {self.done}/{self.total} khong ra anh: {k}")
                    return
                name = re.sub(r'[\\/*?:"<>|]', "_", rec["Ma"] or k)[:60]
                folder = root / name
                folder.mkdir(parents=True, exist_ok=True)
                ok = 0
                for i, iu in enumerate(rec["Image"].split(" | "), 1):
                    try:
                        rr = await cli.get(iu, headers={"Referer": rec["URL"]})
                        rr.raise_for_status()
                        ext = pathlib.Path(urlparse(iu).path).suffix or ".jpg"
                        (folder / f"{name}_{i}{ext}").write_bytes(rr.content)
                        ok += 1
                    except Exception as e:
                        self.log(f"[ERR] tai anh {iu[-45:]}: {type(e).__name__}")
                self.log(f"[OK ] {self.done}/{self.total} {name} -> {ok} anh")

        await asyncio.gather(*(one(k) for k in keys))


def read_excel_col(path: str, sheet: str, col: str) -> list[str]:
    """Loi -> ValueError (KHONG SystemExit: worker GUI bat `except Exception` se khong bat duoc)."""
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

    root_dir = next(p for p in (HERE, *HERE.parents) if (p / "gui_common.py").exists())
    sys.path.insert(0, str(root_dir))
    from gui_common import ask_export_path, attach_excel_picker

    root = tk.Tk()
    root.title("YATO Hand Tools Crawler — yatohandtools.com")
    root.geometry("1080x680")
    ttk.Style().theme_use("clam")

    q: queue.Queue[str] = queue.Queue()
    stop_evt, pause_evt = threading.Event(), threading.Event()
    state: dict[str, Any] = {"running": False, "results": [], "eng": None}

    left = ttk.Frame(root, padding=10)
    left.pack(side="left", fill="y")
    right = ttk.Frame(root, padding=(0, 10, 10, 10))
    right.pack(side="right", fill="both", expand=True)

    mode = tk.StringVar(value="excel")
    ttk.Label(left, text="Pipeline", font=("Segoe UI", 10, "bold")).pack(anchor="w")
    for val, txt in [("code", "1. Theo MA"), ("link", "2. Theo LINK (moi dong 1 link)"),
                     ("excel", "3. Theo cot EXCEL"), ("all", "4. Crawl ALL (~1.5k SP)"),
                     ("images", "5. Tai ANH")]:
        ttk.Radiobutton(left, text=txt, value=val, variable=mode).pack(anchor="w")
    ttk.Label(left, text="(site khong co gia; search cua site hong)",
              foreground="#777").pack(anchor="w")

    ttk.Label(left, text="\nMa:").pack(anchor="w")
    ent = ttk.Entry(left, width=36)
    ent.pack(anchor="w")
    ttk.Label(left, text="Link (moi dong 1 link):").pack(anchor="w")
    box = tk.Text(left, width=36, height=5)
    box.pack(anchor="w")

    xls, sheet, col = tk.StringVar(), tk.StringVar(), tk.StringVar()
    outdir = tk.StringVar(value="out")
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
        except BaseException as e:     # ke ca SystemExit -> khong de thread chet cam
            q.put(f"[FATAL] {type(e).__name__}: {e}")
        finally:                       # LUON nha co, khong khoa START vinh vien
            state["results"] = eng.results
            state["running"] = False
            q.put(f"[SYS] XONG — {len(eng.results)} record. Bam EXPORT de xuat Excel.")

    def start() -> None:
        if state["running"]:
            messagebox.showinfo("Dang chay", "Dang crawl. Bam Stop truoc.")
            return
        m = mode.get()
        if m in ("excel", "images") and not (xls.get() and col.get()):
            messagebox.showwarning("Thieu du lieu", "Chon file Excel va cot truoc.")
            return
        if m == "code" and not ent.get().strip():
            messagebox.showwarning("Thieu du lieu", "Nhap ma truoc.")
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
        p = ask_export_path(xls.get(), HERE, "yatohandtools.xlsx")
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
    ap = argparse.ArgumentParser(description="yatohandtools.com crawler")
    ap.add_argument("--mode", choices=["code", "link", "excel", "all", "images"])
    ap.add_argument("--code")
    ap.add_argument("--link")
    ap.add_argument("--excel")
    ap.add_argument("--sheet", default="")
    ap.add_argument("--col", default="Part")
    ap.add_argument("--outdir", default="out")
    ap.add_argument("--out", default="yatohandtools.xlsx")
    ap.add_argument("--conc", type=int, default=8)
    ap.add_argument("--delay", type=float, default=0.05)
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--build-index", action="store_true",
                    help="dung code-index tu sitemap (~2s) -> tra ma tuc thi")
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

    eng = Engine(a.conc, a.delay)
    t0 = time.time()
    try:
        asyncio.run(eng.run(a.mode, a.code or a.link or "", a.excel or "",
                            a.sheet, a.col, a.limit, a.outdir))
    finally:
        if eng.results and a.mode != "images":
            p = export_excel(eng.results, pathlib.Path(a.out))
            print(f"[SYS] {len(eng.results)} record -> {p} ({time.time() - t0:.1f}s)")
        eng.cleanup()


if __name__ == "__main__":
    main()
