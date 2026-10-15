"""
yato.vn — Yato Tools Viet Nam (nha phan phoi doc quyen). httpx + selectolax, KHONG browser.

RECON (2026-07-16)
------------------
CMS      : PHP tu viet (sitemap.php lo path /home/yatovn/domains/yato.vn/). Server-side render.
API      : KHONG co JSON API => httpx + selectolax la du.
robots   : chi Disallow /dang-nhap.html /dang-ky.html /daotao.html. Sitemap /sitemap.xml duoc phep.
sitemap  : /sitemap.xml -> ~3.857 URL (1 file duy nhat; dau file co warning PHP Deprecated ->
           PHAI cat bo truoc khi parse XML, hoac dung regex <loc> nhu o day).
detail   : /<slug>-p<ID>            vd /bo-15-mui-phay-go-...-yato-yt-68007-p34170
           - h1                                  -> ten SP
           - .price                              -> gia VND (vd "840.000VND"); co SP khong co gia
           - div.c-box_description               -> BOOTSTRAP TABS:
               ul.nav-tabs   li.active = "Gioi thieu san pham" (CHI LA NHAN TAB, KHONG co spec!)
               div.tab-content
                 div.tab-pane#home   (class active) -> bai viet marketing (~6.4k ky tu, 54 <li>)
                 div.tab-pane#menu1                 -> **THONG SO KY THUAT that** (13 <li>
                                                       dang "Nhan:Gia tri") <- LAY CAI NAY
           - anh: /images/products/<Y>/<m>/<d>/large/<slug>_<ts>.jpg (co ban small/resized/original)
           - breadcrumb chi 1 cap (Trang chu > ten SP) -> khong co category huu ich
search   : /tim-kiem/<tu khoa>-page<N>  (PATH-based, KHONG phai query param).
           **KHOP RAT KEM**: 'YT-68007' -> site tra "Tim thay 0 san pham" du SP CO THAT.
           => KHONG dung search de tra ma. Dung CODE-INDEX tu sitemap (ma nam trong slug).
gia      : CO gia VND cong khai.

Chien luoc tra ma (giong toptul, xem AGENTS.md 5.1)
  moi lan tra code/excel deu tai sitemap live. _code_index.json (TTL 30 ngay)
  chi la hint cho URL khong co ma trong slug; hint luon phai fetch PDP live va
  khop lai ma truoc khi dung. Neu hint sai, quet fallback detail tu sitemap.

Run
---
  yato.bat                                     -> GUI
  python yato.py --mode code --code YT-68007
  python yato.py --mode excel --excel parts.xlsx --col Part
  python yato.py --mode keyword --keyword "co le" --limit 50
  python yato.py --mode link --link https://yato.vn/<slug>-p34170
  python yato.py --mode all [--limit 500]      -> quet sach ~3.3k SP
  python yato.py --mode images --excel parts.xlsx --col Part --outdir out
  python yato.py --build-index                 -> slug + detail fallback index
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
from urllib.parse import quote, urljoin, urlparse

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser

HERE = pathlib.Path(__file__).parent
BASE = "https://yato.vn"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")

CODE_INDEX = HERE / "_code_index.json"    # ma -> URL (search yato.vn hong, xem docstring)
INDEX_TTL_DAYS = 30

CIRCUIT_BREAK = 25        # N loi LIEN TIEP -> nghi bi chan IP -> dung han
MAX_ITEMS = 8000          # site co ~3.8k SP
MAX_PAGES = 300
MAX_REDIRECTS = 5
TIMEOUT = 30.0
GC_EVERY = 300

PROD_RE = re.compile(r"https://yato\.vn/[a-z0-9][a-z0-9-]*-p\d+")
# ma Yato/Sthor trong slug: yt-68007, yt-0767, 58705 (sthor) ...
CODE_IN_SLUG_RE = re.compile(r"-(yt-[0-9a-z]+|yato-[0-9a-z-]+|sthor-[0-9a-z-]+)-p\d+", re.I)
CODE_RE = re.compile(r"^(YT|YG|STHOR)?-?[0-9][A-Z0-9./-]*$", re.I)
PRODUCT_CODE_RE = re.compile(
    r"(?<![A-Z0-9])(?P<prefix>YT|YG|YTM|STHOR)\s*-?\s*"
    r"(?P<body>[0-9][A-Z0-9]*)", re.I
)


def extract_product_codes(text: Any) -> list[str]:
    """Extract every Yato-family model token and normalize its separator."""
    out: list[str] = []
    for m in PRODUCT_CODE_RE.finditer(str(text or "")):
        code = f"{m.group('prefix').upper()}-{m.group('body').upper()}"
        if code not in out:
            out.append(code)
    return out


def merge_code_index(index: dict[str, str], url: str, codes: list[str]) -> None:
    """Add all codes for one product without replacing an existing mapping."""
    for code in codes:
        index.setdefault(norm_code(code), url)


def merge_live_index(cached: dict[str, str], live: dict[str, str]) -> dict[str, str]:
    """Refresh live slug mappings while retaining non-slug cache hints."""
    merged = dict(cached)
    merged.update(live)
    return merged


# ================================================================ CORE (khong dung GUI)
def norm_code(s: Any) -> str:
    """Chuan hoa ma de so khop. URL thi GIU NGUYEN hoa/thuong (path phan biet chu)."""
    v = re.sub(r"\s+", "", str(s or ""))
    return v if v.lower().startswith("http") else v.upper()


def search_url(key: str, page: int = 1) -> str:
    """PATH-based: /tim-kiem/<tu khoa>-page<N>. (Search khop kem - chi dung cho mode keyword.)"""
    return f"{BASE}/tim-kiem/{quote(key)}-page{page}"


def codes_from_slug(url: str) -> list[str]:
    """Rut ma san pham tu slug URL. 1 URL co the chua nhieu ma (vd '...yato-yt-0767-...')."""
    slug = urlparse(url).path.rsplit("/", 1)[-1]
    slug = re.sub(r"-p\d+$", "", slug)
    out: list[str] = []
    for m in re.finditer(r"(?:^|-)(yt|yg|ytm)-?(\d[0-9a-z]*)", slug, re.I):
        out.append(f"{m.group(1).upper()}-{m.group(2).upper()}")
    for m in re.finditer(r"(?:^|-)sthor-(\d[0-9a-z]*)", slug, re.I):
        out.append(m.group(1).upper())
    return list(dict.fromkeys(out))


def build_live_slug_index(urls: list[str]) -> tuple[dict[str, str], list[str]]:
    """Build a fresh slug index and return URLs needing live detail inspection."""
    index: dict[str, str] = {}
    fallback: list[str] = []
    for url in dict.fromkeys(urls):
        codes = codes_from_slug(url)
        if codes:
            merge_code_index(index, url, codes)
        else:
            fallback.append(url)
    return index, fallback


def record_matches_code(record: dict | None, code: str) -> bool:
    """Whether a live parsed record explicitly contains the requested code."""
    if not record:
        return False
    wanted = norm_code(code)
    return wanted in {norm_code(c) for c in record.get("_codes", [])}


def parse_links(html: str) -> list[str]:
    """Link san pham trong trang search/category — giu thu tu, bo trung."""
    seen: set[str] = set()
    out: list[str] = []
    for u in PROD_RE.findall(html):
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def parse_sitemap(xml: str) -> list[str]:
    """sitemap.xml co warning PHP o dau file -> dung regex <loc>, khong parse XML nghiem ngat."""
    return [u for u in re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", xml) if PROD_RE.match(u)]


def _kv(line: str) -> tuple[str, str]:
    """'Thuong hieu:Yato - Ba Lan' -> ('Thuong hieu', 'Yato - Ba Lan')."""
    k, sep, v = line.partition(":")
    return (k.strip(), v.strip()) if sep else ("", line.strip())


def parse_spec(t: LexborHTMLParser) -> dict[str, str]:
    """THONG SO nam o tab-pane #menu1 ('Thong so ky thuat').

    CANH BAO: `div.c-box_description > ul > li.active` KHONG phai spec — do la NHAN tab
    ("Gioi thieu san pham") cua ul.nav-tabs. Con #home la bai viet marketing (54 <li>:
    chinh sach bao hanh, so dien thoai chi nhanh...). Chi #menu1 moi la thong so that.
    """
    spec: dict[str, str] = {}
    pane = t.css_first("#menu1")
    if pane is None:                       # fallback: trang khong co tab spec
        pane = t.css_first("#home")
    if pane is None:
        return spec
    for li in pane.css("li"):
        line = re.sub(r"\s+", " ", li.text(strip=True))
        # Older records occasionally put two labelled fields in one <li>:
        # "Ma san pham: YT-07087 Tieu chuan: CE".
        label_start = r"[A-ZÀ-ỸĐ][A-Za-zÀ-ỹĐđ]*(?:\s+[A-Za-zÀ-ỹĐđ]+){0,8}"
        for part in re.split(rf"\s+(?={label_start}:)", line):
            k, v = _kv(part)
            if k and v and len(k) < 60:
                spec.setdefault(k, v)
    return spec


def parse_page(html: str, url: str) -> dict | None:
    t = LexborHTMLParser(html)
    h1 = t.css_first("h1")
    if h1 is None:
        return None
    name = re.sub(r"\s+", " ", h1.text(strip=True))

    p = t.css_first(".price")
    price = re.sub(r"\s+", " ", p.text(strip=True)) if p else ""

    spec = parse_spec(t)
    code_values = [v for k, v in spec.items() if "mã" in k.lower() or "ma" == k.lower()]
    all_codes: list[str] = []
    for value in code_values:
        all_codes.extend(extract_product_codes(value))
        # Vorel/Sthor pages can expose a bare numeric model in this field.
        bare = re.fullmatch(r"[0-9][A-Z0-9./-]*", value.strip(), re.I)
        if bare:
            all_codes.append(value.strip().upper())
    all_codes.extend(extract_product_codes(name))
    all_codes.extend(extract_product_codes(urlparse(url).path))
    all_codes.extend(codes_from_slug(url))
    all_codes = list(dict.fromkeys(norm_code(c) for c in all_codes if c))
    code = all_codes[0] if all_codes else ""

    images: list[str] = []
    for a in t.css("a.MagicZoomPlus, a.Selector"):
        u = a.attributes.get("href") or ""
        if "/images/products/" in u:
            u = urljoin(BASE, u)
            if u not in images:
                images.append(u)
    if not images:                          # mot so trang chi co anh inline
        for im in t.css("img"):
            u = im.attributes.get("src") or im.attributes.get("data-src") or ""
            if "/images/products/" in u and "/cat/" not in u and "img_prd" not in (im.attributes.get("class") or ""):
                u = urljoin(BASE, u)
                if u not in images:
                    images.append(u)

    return {"Input": "", "Ma": code, "Ten": name,
            "Thuong hieu": spec.get("Thương hiệu", ""), "Gia": price, "Status": "OK",
            "_codes": all_codes, "_cats": [], "_spec": spec,
            "Image": " | ".join(images[:10]), "URL": url}


def export_excel(results: list[dict], path: pathlib.Path) -> pathlib.Path:
    """Spec gop 1 COT 'Thong so' (moi dong 'Ten: Gia tri'). Xem AGENTS.md 4.6."""
    from openpyxl.utils import get_column_letter

    fixed = ("Input", "Ma", "Ten", "Thuong hieu", "Gia", "Status")
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
        self.seen_records: set[tuple[str, str]] = set()
        self.done = 0
        self.total = 0
        self.key = ""
        self._ckpt_loaded = False
        self._index: dict[str, str] | None = None

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
            url = r.get("URL", "")
            seen.add(url)
            self.seen_urls.add(url)
            self.seen_records.add((url, norm_code(r.get("Input") or r.get("Ma"))))
        if seen:
            self.log(f"[SYS] Resume: {len(self.results)} record da co -> bo qua")
        return seen

    def add(self, recs: list[dict], key: str) -> None:
        fresh = []
        for r in recs:
            identity = (r.get("URL", ""), norm_code(r.get("Input") or r.get("Ma")))
            if r.get("URL") and identity not in self.seen_records:
                self.seen_records.add(identity)
                fresh.append(r)
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

    # ---- CODE-INDEX: search yato.vn hong (YT-68007 -> 0 KQ) nen ma -> URL lay tu SLUG sitemap
    def _load_index(self) -> dict[str, str]:
        if self._index is not None:
            return self._index
        self._index = {}
        if CODE_INDEX.exists() and (time.time() - CODE_INDEX.stat().st_mtime) / 86400 < INDEX_TTL_DAYS:
            try:
                self._index = json.loads(CODE_INDEX.read_text("utf-8"))
                self.log(f"[SYS] Code-index: {len(self._index)} ma da biet")
            except Exception:
                self._index = {}
        return self._index

    def save_index(self) -> None:
        if self._index:
            CODE_INDEX.write_text(json.dumps(self._index, ensure_ascii=False), "utf-8")

    async def build_index(self, cli: httpx.AsyncClient) -> dict[str, str]:
        """Sitemap -> index {ma -> URL}; detail fallback cho URL khong co ma trong slug."""
        live_idx, detail_urls, _ = await self._load_live_sitemap(cli)
        if not live_idx and not detail_urls:
            self.log("[ERR] Khong tai duoc sitemap")
            return self._load_index()
        idx = dict(live_idx)
        if detail_urls:
            sem = asyncio.Semaphore(self.conc)

            async def inspect(u: str) -> tuple[str, list[str]]:
                async with sem:
                    html = await self.fetch(cli, u)
                    rec = parse_page(html, u) if html else None
                    return u, (rec or {}).get("_codes", [])

            for u, codes in await asyncio.gather(*(inspect(u) for u in detail_urls)):
                merge_code_index(idx, u, codes)
        self._index = idx
        self.save_index()
        self.log(f"[SYS] Code-index: live sitemap + {len(detail_urls)} detail fallback -> {len(idx)} ma")
        return idx

    async def _load_live_sitemap(self, cli: httpx.AsyncClient) -> tuple[dict[str, str], list[str], set[str]]:
        """Fetch sitemap on every lookup; persisted index is never authoritative."""
        xml = await self.fetch(cli, f"{BASE}/sitemap.xml")
        if not xml:
            return {}, [], set()
        urls = list(dict.fromkeys(parse_sitemap(xml)))
        index, fallback = build_live_slug_index(urls)
        return index, fallback, set(urls)

    async def _resolve_codes(self, cli: httpx.AsyncClient, codes: list[str]) -> tuple[list[str], dict[str, list[str]], dict[str, dict]]:
        """Resolve against a fresh sitemap; cache is only a live-validated hint."""
        live_idx, fallback_urls, live_urls = await self._load_live_sitemap(cli)
        cached_idx = self._load_index()
        preloaded: dict[str, dict] = {}
        unresolved = []
        for code in codes:
            if code in live_idx:
                continue
            hint = cached_idx.get(code)
            if hint and hint in live_urls:
                rec = preloaded.get(hint)
                if rec is None:
                    html = await self.fetch(cli, hint)
                    rec = parse_page(html, hint) if html else None
                if record_matches_code(rec, code):
                    live_idx[code] = hint
                    preloaded[hint] = rec
                    continue
                self.log(f"[WARN] Cache hint khong khop {code} -> bo qua: {hint[-50:]}")
            unresolved.append(code)

        if unresolved and fallback_urls:
            sem = asyncio.Semaphore(min(self.conc, 2))

            async def inspect(u: str) -> tuple[str, dict | None]:
                async with sem:
                    html = await self.fetch(cli, u)
                    return u, (parse_page(html, u) if html else None)

            for u, rec in await asyncio.gather(*(inspect(u) for u in fallback_urls)):
                if not rec:
                    continue
                preloaded[u] = rec
                for found in rec.get("_codes", []):
                    live_idx.setdefault(norm_code(found), u)
        self._index = merge_live_index(cached_idx, live_idx)
        self.save_index()
        urls: list[str] = []
        inputs: dict[str, list[str]] = {}
        miss: list[str] = []
        for c in codes:
            u = live_idx.get(c)
            if u:
                urls.append(u)
                inputs.setdefault(u, []).append(c)
            else:
                miss.append(c)
        if miss:
            self.log(f"[MISS] {len(miss)} ma khong co trong sitemap: {', '.join(miss[:5])}"
                     f"{'...' if len(miss) > 5 else ''}")
        return urls, inputs, preloaded

    async def _crawl_urls(self, cli: httpx.AsyncClient, urls: list[str], key: str,
                           inputs: dict[str, list[str]] | None = None,
                           preloaded: dict[str, dict] | None = None) -> None:
        seen = self._load_ckpt(key)
        urls = [u for u in dict.fromkeys(urls) if u not in seen][:MAX_ITEMS]
        self.total, self.done = len(urls) + len(seen), len(seen)
        sem = asyncio.Semaphore(self.conc)

        async def one(u: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                rec = (preloaded or {}).get(u)
                if rec is None:
                    html = await self.fetch(cli, u)
                    await asyncio.sleep(self.delay)
                    rec = parse_page(html, u) if html else None
                if rec:
                    requested = (inputs or {}).get(u, [])
                    if isinstance(requested, str):
                        requested = [requested]
                    if requested:
                        expanded = []
                        record_codes = {norm_code(c) for c in rec.get("_codes", [])}
                        for item in requested:
                            copy = dict(rec)
                            copy["Input"] = item
                            copy["Status"] = (
                                "OK" if norm_code(item) in record_codes
                                else "Matched shared product"
                            )
                            expanded.append(copy)
                        self.add(expanded, key)
                    else:
                        rec["Input"] = rec["Ma"]
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
        headers = {"User-Agent": UA, "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8"}
        async with httpx.AsyncClient(headers=headers, timeout=TIMEOUT, follow_redirects=True,
                                     max_redirects=MAX_REDIRECTS,
                                     limits=httpx.Limits(max_connections=self.conc + 4)) as cli:
            if mode == "all":
                xml = await self.fetch(cli, f"{BASE}/sitemap.xml")
                urls = list(dict.fromkeys(parse_sitemap(xml or "")))[:limit or MAX_ITEMS]
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
                    urls, inputs, preloaded = await self._resolve_codes(cli, codes)
                    await self._crawl_urls(cli, urls, key, inputs, preloaded)
                have = {norm_code(v) for r in self.results
                        for v in (r.get("Input"), r.get("Ma")) if v}
                for c in codes:                       # ma miss van co dong Not Found
                    if c not in have:
                        self.results.append({"Input": c, "Ma": c, "Ten": "", "Thuong hieu": "",
                                             "Gia": "", "Status": "Not Found", "_cats": [],
                                             "_spec": {}, "Image": "", "URL": ""})

            elif mode == "keyword":
                urls, page = [], 1
                while page <= MAX_PAGES and await self._gate():
                    html = await self.fetch(cli, search_url(target, page))
                    new = [u for u in (parse_links(html) if html else []) if u not in urls]
                    if not new:                       # stuck-page = het ket qua
                        break
                    urls += new
                    self.log(f"[SYS] '{target}' trang {page} -> {len(urls)} URL")
                    if limit and len(urls) >= limit:
                        break
                    page += 1
                    await asyncio.sleep(self.delay)
                if not urls:
                    self.log("[MISS] 0 ket qua. Search yato.vn khop RAT KEM (vd 'YT-68007' -> 0 KQ) "
                             "-> dung mode 'code'/'excel' (tra qua sitemap) hoac doi tu khoa.")
                await self._crawl_urls(cli, urls[:limit or MAX_ITEMS], key)

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
        idx_urls, inputs, _ = await self._resolve_codes(
            cli, [k for k in keys if not k.lower().startswith("http")])
        by_code = {code: u for u, codes_for_url in inputs.items() for code in codes_for_url}
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
    root.title("YATO Crawler — yato.vn")
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
    for val, txt in [("code", "1. Theo MA"), ("keyword", "2. Theo TU KHOA"),
                     ("link", "3. Theo LINK (moi dong 1 link)"), ("excel", "4. Theo cot EXCEL"),
                      ("all", "5. Crawl ALL (~3.3k SP)"), ("images", "6. Tai ANH")]:
        ttk.Radiobutton(left, text=txt, value=val, variable=mode).pack(anchor="w")

    ttk.Label(left, text="\nMa / Tu khoa:").pack(anchor="w")
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
        if m in ("code", "keyword") and not ent.get().strip():
            messagebox.showwarning("Thieu du lieu", "Nhap ma / tu khoa truoc.")
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
        p = ask_export_path(xls.get(), HERE, "yato.xlsx")
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
    ap = argparse.ArgumentParser(description="yato.vn crawler")
    ap.add_argument("--mode", choices=["code", "keyword", "link", "excel", "all", "images"])
    ap.add_argument("--code")
    ap.add_argument("--keyword")
    ap.add_argument("--link")
    ap.add_argument("--excel")
    ap.add_argument("--sheet", default="")
    ap.add_argument("--col", default="Part")
    ap.add_argument("--outdir", default="out")
    ap.add_argument("--out", default="yato.xlsx")
    ap.add_argument("--conc", type=int, default=8)
    ap.add_argument("--delay", type=float, default=0.05)
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--build-index", action="store_true",
                    help="build day du cache hint tu sitemap + detail fallback (khong la nguon ket qua)")
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
        asyncio.run(eng.run(a.mode, a.code or a.keyword or a.link or "", a.excel or "",
                            a.sheet, a.col, a.limit, a.outdir))
    finally:
        if eng.results and a.mode != "images":
            p = export_excel(eng.results, pathlib.Path(a.out))
            print(f"[SYS] {len(eng.results)} record -> {p} ({time.time() - t0:.1f}s)")
        eng.cleanup()


if __name__ == "__main__":
    main()
