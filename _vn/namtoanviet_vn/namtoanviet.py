"""
namtoanviet.vn (Nam Toàn Việt) — httpx + selectolax/regex, NO browser, NO Selenium.

RECON (2026-08-10)
-------------------
Nen tang : PHP tu viet (X-Powered-By: PHP/7.4.33, LiteSpeed), KHONG phai Sapo/Haravan/
           Shopify. jQuery, khong SPA.
API      : KHONG co JSON API cong khai. HTML render server-side (curl thuong doc duoc).
sitemap  : /sitemap.xml (sitemapindex) -> sitemap-products.xml THAT, MOI (lastmod = hom
           nay, tu sinh) -> 4.084 URL san pham. Dung lam MASTER LIST (khong can BFS
           category). sitemap-categories.xml co 210 category (khong can, breadcrumb
           trong JSON-LD da du category-path cho tung SP).
detail   : GET <slug san pham> (slug = ten khong dau, KHONG phai ma). Data nam o:
    - <script type="application/ld+json"> block "Product": sku (ma he thong "sp-NNN"),
      mpn (ma hang, vd "DCD805N-B1"), brand, name, image[], offers.AggregateOffer
      (lowPrice/highPrice/offerCount).
    - block "BreadcrumbList": duong dan category day du (bo muc 1 "San pham" va muc
      cuoi = chinh SP do).
    - the <table> 2 cot key:value = thong so ky thuat (khong co rowspan/colspan).
    - BAY: offerCount trong JSON-LD (2-4) KHONG khop voi so nut bien the that (0-2)
      tren trang — offer thua co the la gia tra gop/gia so sanh, KHONG phai bien
      the that. Nguon THAT cho bien the la <div class="btn-option-variant"
      data-name=... data-gia-number=...> (co the 0 cai = SP khong bien the).
      SP khong bien the -> gia lay tu <span class="price-new">/<span class="price-old">
      trong #price-area (JSON-LD offers khong dang tin cho truong hop nay).
    - 404 that (slug sai) -> HTTP 404 sach, khong co the ld+json Product.
search   : /tim-kiem?q=<tu khoa> — DA THU voi tu khoa CHINH XAC lay tu ten SP that
    (vd "dewalt", "khoan") -> LUON tra "Khong tim thay ket qua", ke ca ma/tu dung.
    KHONG dung duoc (co the site dang loi, hoac can token/POST khac ma chua ro).
    robots.txt cung tu Disallow /tim-kiem — co the chinh site cung biet no hong.
    => KHONG dung search. Code/keyword mode tu dung nguon rieng, xem duoi.
ma-index : Khong co search that -> tra ma theo 2 tang:
    1. DIRECT (0 request phu): chuan hoa ma (bo dau, bo ky tu khong phai chu/so,
       thuong hoa) roi so SUBSTRING voi slug (cung chuan hoa) trong danh sach sitemap
       da tai san — da do that: slug "may-khoan-dong-luc-dewalt-dcd805n" CHUA
       "dcd805n" -> khop ngay, khoi can build index.
    2. FALLBACK: ma khong nam trong slug nao (vd trung ma he thong "sp-323" hoac
       ma hang khac cach dat ten slug) -> CODE_INDEX cache (sku+mpn -> url, xay bang
       cach crawl HET sitemap 1 lan -> cache ra dia TTL 3 ngay (sitemap tu sinh MOI
       NGAY -> catalog doi nhanh hon toptul, TTL ngan hon 30 ngay cua toptul).
       DA DO THAT (2026-08-10, conc=8): 4.082 trang het **45 phut** (~0.66s/trang —
       KHOP toc do 1 trang don le, nhung TONG cham hon uoc tinh ban dau (~5 phut) rat
       nhieu — site nay KHONG song song tot nhu toptul/insumart, co the do server
       PHP/LiteSpeed xu ly tuan tu hoac rate-limit ngam khong tra ma loi ro rang.
       KHONG DUOC doan lai so nay, dung so DA DO. Nguoi dung nen build 1 lan roi
       de cache lam viec (TTL 3 ngay), KHONG build lai tuy tien.
    Chua ma nao trong index/slug -> "Not Found" that (KHONG con noi nao khac de tim).
gia      : cong khai VND, khong can dang nhap.
charset  : utf-8.

Run
---
  namtoanviet.bat                                          -> GUI
  python namtoanviet.py --mode all [--limit 500]           -> quet sach sitemap (~4084 SP)
  python namtoanviet.py --mode code --code DCD805N
  python namtoanviet.py --mode keyword --keyword "may khoan dewalt"
  python namtoanviet.py --mode link --link https://www.namtoanviet.vn/may-khoan-dong-luc-dewalt-dcd805n
  python namtoanviet.py --mode excel --excel parts.xlsx --sheet Sheet1 --col Part

  Keyword / ma khong co trong slug -> lan dau se XAY CODE_INDEX (crawl het sitemap,
  vai phut), cac lan sau tra cuu tuc thi (cache TTL 3 ngay).
"""
from __future__ import annotations

import argparse
import asyncio
import gc
import html as htmlmod
import json
import os
import pathlib
import queue
import re
import threading
import time
import unicodedata
from typing import Any
from urllib.parse import urljoin

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser

HERE = pathlib.Path(__file__).parent
BASE = "https://www.namtoanviet.vn"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
HEADERS = {"User-Agent": UA, "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8"}

CODE_INDEX = HERE / "_code_index.json"   # ma (sku+mpn) -> {url,name}; sitemap tu sinh
INDEX_PROGRESS = HERE / "_code_index_progress.json"  # URL index da xong + trang thai complete
INDEX_TTL_DAYS = 3                       # moi ngay -> TTL ngan hon toptul (30 ngay)

CIRCUIT_BREAK = 25       # N loi LIEN TIEP -> nghi bi chan IP -> dung han
MAX_ITEMS = 20000        # tran work-list (catalog that ~4.1k, du du)
MAX_REDIRECTS = 5
TIMEOUT = 30.0
GC_EVERY = 500

LD_RE = re.compile(r'<script type="application/ld\+json">(.*?)</script>', re.S)
LOC_RE = re.compile(r"<loc>([^<]+)</loc>")
VARIANT_RE = re.compile(r'<div class="btn-option-variant[^"]*"([^>]*)>')
ATTR_RE = re.compile(r'([\w-]+)="([^"]*)"')
PRICE_NEW_RE = re.compile(r'<span class="price-new">([^<]*)</span>')
PRICE_OLD_RE = re.compile(r'<span class="price-old">([^<]*)</span>')
TABLE_RE = re.compile(r"<table[^>]*>(.*?)</table>", re.S | re.I)
TR_RE = re.compile(r"<tr[^>]*>(.*?)</tr>", re.S | re.I)
TD_RE = re.compile(r"<t[dh][^>]*>(.*?)</t[dh]>", re.S | re.I)
TAG_RE = re.compile(r"<[^>]+>")
SLUG_CLEAN_RE = re.compile(r"[^a-z0-9]+")
MIN_SLUG_CODE_LEN = 6
XLSX_ILLEGAL_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


# ================================================================ CORE (khong dung GUI)
def norm_code(s: Any) -> str:
    """Chuan hoa ma de so khop. URL thi GIU NGUYEN hoa/thuong (path phan biet chu)."""
    v = re.sub(r"\s+", "", str(s or ""))
    return v if v.lower().startswith("http") else v.upper()


def slugify_match(s: str) -> str:
    """Chuan hoa de so KHOP SUBSTRING voi slug: bo het ky tu khong phai a-z0-9,
    thuong hoa. Dung cho ca ma nhap lan slug URL (khong xu ly dau tieng Viet vi
    slug goc da khong dau san)."""
    return SLUG_CLEAN_RE.sub("", str(s or "").lower())


def select_slug_candidate(code: str, entries: list[tuple[str, str]]) -> tuple[str | None, str]:
    """Return one safe tier-1 slug match, never the first arbitrary substring hit.

    Entries are (url, raw_slug) tuples. Exact slug-token / contiguous-token
    matches are preferred. Raw substring is only safe when it identifies one URL.
    """
    key = slugify_match(code)
    if len(key) < MIN_SLUG_CODE_LEN:
        return None, "short"

    token_hits: list[str] = []
    for url, raw_slug in entries:
        tokens = [slugify_match(token) for token in raw_slug.split("-") if slugify_match(token)]
        if any(key == "".join(tokens[start:end])
               for start in range(len(tokens)) for end in range(start + 1, len(tokens) + 1)):
            token_hits.append(url)
    token_hits = list(dict.fromkeys(token_hits))
    if len(token_hits) == 1:
        return token_hits[0], "token"
    if len(token_hits) > 1:
        return None, "ambiguous-token"

    raw_hits = [url for url, raw_slug in entries if key in slugify_match(raw_slug)]
    raw_hits = list(dict.fromkeys(raw_hits))
    if len(raw_hits) == 1:
        return raw_hits[0], "substring"
    return None, "ambiguous-substring" if raw_hits else "miss"


def strip_html(s: str) -> str:
    return htmlmod.unescape(TAG_RE.sub("", s or "")).replace("\xa0", " ").strip()


def parse_price(s: str) -> int | None:
    digits = re.sub(r"[^\d]", "", s or "")
    return int(digits) if digits else None


def alias_from_url(u: str) -> str:
    a = re.sub(r"^https?://[^/]+", "", str(u or "").strip()).split("?")[0].split("#")[0]
    return a.strip("/")


def _tree(html_or_tree: str | LexborHTMLParser) -> LexborHTMLParser:
    return html_or_tree if isinstance(html_or_tree, LexborHTMLParser) else LexborHTMLParser(html_or_tree)


def _node_text(node: Any) -> str:
    return re.sub(r"\s+([,.;:!?])", r"\1", " ".join((node.text(separator=" ", strip=True) or "").split()))


def parse_spec_table(html_or_tree: str | LexborHTMLParser) -> dict[str, str]:
    """Bang spec 2 cot key:value (khong rowspan/colspan — da kiem tra thuc te)."""
    out: dict[str, str] = {}
    for table in _tree(html_or_tree).css("table"):
        for tr in table.css("tr"):
            tds = tr.css("th, td")
            if len(tds) != 2:
                continue
            k = _node_text(tds[0]).rstrip(":").strip()
            v = _node_text(tds[1])
            if k and v and k not in out:
                out[k] = v
    return out


def parse_ld_json(html_or_tree: str | LexborHTMLParser) -> tuple[dict | None, list[str]]:
    """<script ld+json> Product + BreadcrumbList -> (product, category_path).
    category_path = bo muc dau "San pham" (chung chung) va muc cuoi (chinh SP do)."""
    product = None
    cats: list[str] = []
    for node in _tree(html_or_tree).css('script[type="application/ld+json"]'):
        try:
            d = json.loads(node.text() or "")
        except Exception:
            continue
        t = d.get("@type")
        if t == "Product":
            product = d
        elif t == "BreadcrumbList":
            items = d.get("itemListElement") or []
            cats = [it.get("name", "") for it in items[1:-1] if it.get("name")]
    return product, cats


def parse_variants(html_or_tree: str | LexborHTMLParser) -> list[dict]:
    """<div class="btn-option-variant" data-name=... data-gia-number=...> — nguon
    THAT cho bien the (KHAC offerCount trong JSON-LD, xem BAY trong docstring dau file)."""
    out = []
    for node in _tree(html_or_tree).css("div.btn-option-variant"):
        attrs = node.attributes
        out.append({
            "name": htmlmod.unescape(attrs.get("data-name", "")),
            "price": int(attrs["data-gia-number"]) if attrs.get("data-gia-number", "").isdigit() else None,
            "photo": attrs.get("data-photo", ""),
        })
    return out


def parse_product_info(tree: LexborHTMLParser) -> dict[str, str]:
    """Cac truong hien thi trong PDP, khong nam day du trong JSON-LD."""
    aliases = {"thương hiệu": "Thuong hieu", "model": "Model",
               "bảo hành": "Bao hanh", "lượt xem": "Luot xem"}
    fields: dict[str, str] = {}
    for item in tree.css(".product-info-list .product-info-item"):
        label, value = item.css_first(".info-label"), item.css_first(".info-value")
        normalized = unicodedata.normalize("NFKD", _node_text(label).rstrip(":")) if label else ""
        normalized = normalized.encode("ascii", "ignore").decode().lower()
        key = {"thuong hieu": "Thuong hieu", "model": "Model", "bao hanh": "Bao hanh",
               "luot xem": "Luot xem"}.get(normalized)
        if key and value:
            fields[key] = _node_text(value)
    return fields


def parse_description(tree: LexborHTMLParser) -> str:
    """Giu text nguoi dung doc duoc cua khung Chi tiet san pham, theo doan."""
    node = tree.css_first("#descToggle") or tree.css_first(".product-description")
    if not node:
        return ""
    parts = [_node_text(part) for selector in ("h1", "h2", "h3", "h4", "h5", "h6", "p", "li")
             for part in node.css(selector)]
    return "\n".join(dict.fromkeys(p for p in parts if p)) or _node_text(node)


def parse_page(html: str, url: str) -> list[dict]:
    """1 trang detail -> list record (1 record / bien the, hoac 1 record neu SP
    khong bien the). Tra [] neu KHONG phai trang san pham (404 / slug sai)."""
    tree = LexborHTMLParser(html)
    product, cats = parse_ld_json(tree)
    if not product:
        return []
    sku = str(product.get("sku") or "").strip()
    mpn = str(product.get("mpn") or "").strip()
    name = product.get("name") or ""
    info = parse_product_info(tree)
    brand = info.get("Thuong hieu") or ((product.get("brand") or {}).get("name") if isinstance(product.get("brand"), dict) else "")
    imgs = [urljoin(BASE, i) for i in (product.get("image") or []) if i]
    spec = parse_spec_table(tree)
    description = parse_description(tree)

    def _rec(variant: str, price: int | None, price_old: int | None, photo: str) -> dict:
        # data-photo la TEN FILE TRAN (khong co duong dan) — anh that nam trong
        # upload/product/<ten file>, da kiem tra thuc te qua the <img src=...>.
        img_list = ([urljoin(BASE, "upload/product/" + photo)] if photo else []) + imgs
        return {
            "Item No": mpn or sku, "SKU": sku, "Page": name, "Status": "OK",
            "_cats": ([brand] if brand else []) + cats, "_spec": spec,
            "Thuong hieu": brand, "Model": info.get("Model") or mpn,
            "Bao hanh": info.get("Bao hanh", ""), "Luot xem": info.get("Luot xem", ""),
            "Mo ta": description,
            "Bien the": variant, "Gia (VND)": price, "Gia goc (VND)": price_old,
            "Image": " | ".join(dict.fromkeys(img_list)),
            "URL": url,
        }

    variants = parse_variants(tree)
    if not variants:
        m1, m2 = PRICE_NEW_RE.search(html), PRICE_OLD_RE.search(html)
        price = parse_price(m1.group(1)) if m1 else None
        price_old = parse_price(m2.group(1)) if m2 else None
        offers = (product.get("offers") or {})
        if price is None:
            price = parse_price(str(offers.get("lowPrice") or ""))
        return [_rec("", price, price_old, "")]
    return [_rec(v["name"], v["price"], None, v["photo"]) for v in variants]


def not_found(code: str, status: str = "Not Found") -> dict:
    return {"Item No": code, "SKU": "", "Page": "", "Status": status,
            "_cats": [], "_spec": {}, "Bien the": "", "Gia (VND)": None,
            "Gia goc (VND)": None, "Thuong hieu": "", "Model": "", "Bao hanh": "",
            "Luot xem": "", "Mo ta": "", "Image": "", "URL": ""}


def failure_record(url: str, input_value: str = "") -> dict:
    """Keep a transport/parser failure visible and resumable; never silently lose a PDP."""
    record = not_found(alias_from_url(url), status="Failed")
    record["Input"] = input_value
    record["URL"] = url
    return record


def empty_page_record(url: str, input_value: str = "", fetched: bool = False) -> dict:
    """Distinguish a fetched non-PDP URL from an unresolved transport failure."""
    record = not_found(alias_from_url(url), status="Not Found" if fetched else "Failed")
    record["Input"] = input_value
    record["URL"] = url
    return record


def excel_safe(value: Any) -> Any:
    """Remove XML 1.0 control characters which openpyxl cannot write to XLSX."""
    return XLSX_ILLEGAL_RE.sub("", value) if isinstance(value, str) else value


def export_excel(results: list[dict], path: pathlib.Path) -> pathlib.Path:
    """Spec gop 1 COT 'Thong so' — khong trai union key thanh hang tram cot rong."""
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

    max_cat = max((len(r.get("_cats") or []) for r in results), default=0)
    rows = []
    for r in results:
        row = {"Input": r.get("Input", ""), "Item No": r.get("Item No", ""),
               "SKU": r.get("SKU", ""), "Page": r.get("Page", ""), "Status": r.get("Status", "")}
        row["Thuong hieu"] = r.get("Thuong hieu", "")
        row["Model"] = r.get("Model", "")
        row["Bao hanh"] = r.get("Bao hanh", "")
        row["Luot xem"] = r.get("Luot xem", "")
        cats = r.get("_cats") or []
        for i in range(max_cat):
            row[f"Category {i + 1}"] = cats[i] if i < len(cats) else ""
        row["Bien the"] = r.get("Bien the", "")
        row["Gia (VND)"] = r.get("Gia (VND)")
        row["Gia goc (VND)"] = r.get("Gia goc (VND)")
        row["Thong so"] = "\n".join(f"{k}: {v}" for k, v in (r.get("_spec") or {}).items())
        row["Mo ta"] = r.get("Mo ta", "")
        row["Image"] = r.get("Image", "")
        row["URL"] = r.get("URL", "")
        rows.append({key: excel_safe(value) for key, value in row.items()})
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, index=False)
        ws = xw.sheets["Sheet1"]
        col = get_column_letter(df.columns.get_loc("Thong so") + 1)
        desc_col = get_column_letter(df.columns.get_loc("Mo ta") + 1)
        ws.column_dimensions[col].width = 60
        ws.column_dimensions[desc_col].width = 80
        align = Alignment(wrap_text=True, vertical="top")   # 1 object dung chung, khong tao moi/o
        for c in [*ws[col][1:], *ws[desc_col][1:]]:
            c.alignment = align
    return path


class Engine:
    """Crawler core — chay headless/CLI duoc, khong phu thuoc GUI."""

    def __init__(self, conc: int = 8, delay: float = 0.05, log=print,
                 stop_evt: threading.Event | None = None,
                 pause_evt: threading.Event | None = None) -> None:
        self.conc, self.delay, self.log = max(1, conc), max(0.0, delay), log
        self.stop_evt = stop_evt or threading.Event()
        self.pause_evt = pause_evt or threading.Event()
        self.fails = 0
        self.results: list[dict] = []
        self.done = 0
        self.total = 0
        self.key = ""
        self._ckpt_loaded = False
        self._index: dict[str, dict] | None = None
        # BAY (dinh that khi build): index_records() duoc goi trong MOI phase crawl
        # (link/code/all) de "lam giau" index tien ich cho lan sau -> self._index tro
        # thanh dict KHONG RONG chi sau 1-2 SP crawl le te. Neu dung "idx rong hay
        # khong" de quyet dinh co can build_index() (quet HET sitemap) hay khong thi
        # 1 link crawl truoc do se lam "idx co du lieu" gia, khien build_index() KHONG
        # BAO GIO chay, ma/tu khoa that (chua tung crawl qua) bi bao "Not Found" oan.
        # -> dung co RIENG bao index da quet DU CHUA, khong suy tu do dict rong/khong.
        self._index_complete = False
        self._index_done_urls: set[str] = set()
        self._index_progress_loaded = False
        self._unverified_codes: set[str] = set()
        self._sitemap_urls: list[str] | None = None

    # ---- resume: JSONL sidecar (append tung item; chay lai = doc lai, bo qua da xong)
    def ckpt(self, key: str) -> pathlib.Path:
        return HERE / f"_raw_{re.sub(r'[^A-Za-z0-9]+', '_', key)[:40]}.jsonl"

    def _load_ckpt(self, key: str) -> set[str]:
        p = self.ckpt(key)
        seen: set[str] = set()
        stale = 0
        recheck = 0
        if self._ckpt_loaded:
            return {r.get("URL", "") for r in self.results}
        self._ckpt_loaded = True
        if not p.exists():
            return seen
        for line in p.read_text("utf-8", errors="ignore").splitlines():
            try:
                r = json.loads(line)
            except Exception:
                continue
            if r.get("Status") != "OK":
                recheck += 1
                continue
            if r.get("Status") == "OK" and not {"Thuong hieu", "Model", "Bao hanh", "Luot xem", "Mo ta"}.issubset(r):
                stale += 1
                continue
            self.results.append(r)
            seen.add(r.get("URL", ""))
        if stale:
            self.log(f"[SYS] Resume: bo qua {stale} record cu thieu schema moi -> crawl lai URL do")
        if recheck:
            self.log(f"[SYS] Resume: xac minh lai {recheck} Not Found/Unverified cu tren web")
        if seen:
            self.log(f"[SYS] Resume: {len(self.results)} record da xong -> bo qua")
        return seen

    def _save(self, key: str, recs: list[dict]) -> None:
        with self.ckpt(key).open("a", encoding="utf-8") as f:
            for r in recs:
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
            f.flush()

    def cleanup(self) -> None:
        """Xong sach -> xoa checkpoint. Bi Stop/crash -> GIU lai de lan sau resume."""
        if not self.key:
            return
        p = self.ckpt(self.key)
        has_failed = any(r.get("Status") == "Failed" for r in self.results)
        if p.exists() and not self.stop_evt.is_set() and not has_failed:
            p.unlink()
        elif has_failed:
            self.log("[SYS] Con record Failed -> giu checkpoint de run sau retry dung URL do")

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
                if r.status_code in (403, 429, 503):          # WAF / rate-limit -> lui dan
                    wait = min(2 ** i * 5, 60)
                    self.log(f"[WARN] {r.status_code} -> ngu {wait}s | {url[-60:]}")
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
                self.log(f"[ERR] {type(e).__name__} | {url[-60:]}")
                if self.fails >= CIRCUIT_BREAK:
                    self.log(f"[STOP] {CIRCUIT_BREAK} loi LIEN TIEP - nghi bi chan IP. DUNG.")
                    self.stop_evt.set()
                return None
        return None

    # ---- sitemap = MASTER LIST (khong BFS, khong category-crawl — xem docstring)
    async def sitemap_urls(self, cli: httpx.AsyncClient) -> list[str]:
        if self._sitemap_urls is not None:
            return self._sitemap_urls
        idx = await self.fetch(cli, f"{BASE}/sitemap.xml")
        subs = [u for u in LOC_RE.findall(idx or "") if "sitemap-products" in u] or [f"{BASE}/sitemap-products.xml"]
        xml = await self.fetch(cli, subs[0])
        urls = list(dict.fromkeys(LOC_RE.findall(xml or "")))
        urls = [u for u in urls if alias_from_url(u) and alias_from_url(u) != "san-pham"]
        self._sitemap_urls = urls
        return urls

    # ---- code-index: sku+mpn+ten -> URL, xay bang crawl HET sitemap (do that: ~45 phut)
    def _load_index(self) -> dict[str, dict]:
        if self._index is not None:
            return self._index
        self._index = {}
        if CODE_INDEX.exists() and (time.time() - CODE_INDEX.stat().st_mtime) / 86400 < INDEX_TTL_DAYS:
            try:
                self._index = json.loads(CODE_INDEX.read_text("utf-8"))
                self.log(f"[SYS] Code-index: {len(self._index)} ma da biet (cache)")
            except Exception:
                self._index = {}
        if self._index and INDEX_PROGRESS.exists():
            try:
                progress = json.loads(INDEX_PROGRESS.read_text("utf-8"))
                if progress.get("version") == 1:
                    self._index_done_urls = set(progress.get("completed_urls") or [])
                    self._index_complete = bool(progress.get("complete"))
                    self._index_progress_loaded = True
            except Exception:
                self.log("[WARN] Code-index progress hong, se kiem tra lai cac URL chua xac nhan")
        return self._index

    def _write_json(self, path: pathlib.Path, data: Any) -> None:
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False), "utf-8")
        os.replace(tmp, path)

    def save_index(self) -> None:
        if self._index is not None:
            self._write_json(CODE_INDEX, self._index)
        self._write_json(INDEX_PROGRESS, {"version": 1, "complete": self._index_complete,
                                           "completed_urls": sorted(self._index_done_urls)})

    def index_records(self, recs: list[dict]) -> None:
        idx = self._load_index()
        for r in recs:
            url, name = r.get("URL"), r.get("Page")
            if not url:
                continue
            for key in (r.get("Item No"), r.get("SKU")):
                if key:
                    idx[norm_code(key)] = {"url": url, "name": name}

    async def build_index(self, cli: httpx.AsyncClient, force: bool = False) -> dict[str, dict]:
        """Quet SACH sitemap -> code-index day du. Cache TTL 3 ngay (sitemap tu sinh moi
        ngay -> catalog doi nhanh). DA DO THAT: 4.082 trang ~45 phut (xem docstring dau
        file) — CHAM hon nhieu site khac trong workshop nay, bao truoc de nguoi dung
        khong tuong bi treo."""
        urls = await self.sitemap_urls(cli)
        self.log(f"[SYS] Xay code-index: quet {len(urls)} trang (do THAT ~45 phut cho "
                 f"4k trang — site nay khong nhanh, dung Stop neu can dung giua chung, "
                 f"resume se tiep tuc tu ma con thieu)...")
        idx = self._load_index()
        sitemap_set = set(urls)
        if force and self._index_complete:
            self.log("[SYS] Cache miss -> quet lai TOAN BO catalog live truoc khi ket luan Not Found")
            self._index_complete = False
            self._index_done_urls.clear()
            self._index_progress_loaded = True
        if not self._index_progress_loaded:
            # Legacy cache khong co progress: chi tin URL da co record, con lai kiem tra lai.
            self._index_done_urls = {v.get("url") for v in idx.values()
                                     if isinstance(v, dict) and v.get("url") in sitemap_set}
        todo = [u for u in urls if u not in self._index_done_urls]
        if not todo:
            self._index_complete = True
            self.save_index()
            return idx
        work: asyncio.Queue[str] = asyncio.Queue()
        for u in todo:
            work.put_nowait(u)
        done = 0

        async def worker() -> None:
            nonlocal done
            while not work.empty():
                if not await self._gate():
                    return
                try:
                    u = work.get_nowait()
                except asyncio.QueueEmpty:
                    return
                html = await self.fetch(cli, u)
                if html:
                    self.index_records(parse_page(html, u))
                    self._index_done_urls.add(u)
                done += 1
                if done % 50 == 0:
                    self.save_index()          # 45 phut la lau -> Stop giua chung van GIU duoc phan da quet
                    self.log(f"[SYS] index: {done}/{len(todo)} trang | {len(idx)} ma")

        await asyncio.gather(*(worker() for _ in range(self.conc)))
        self._index_complete = not self.stop_evt.is_set() and sitemap_set.issubset(self._index_done_urls)
        self.save_index()
        self.log(f"[SYS] Code-index xong: {len(idx)} ma"
                 + ("" if self._index_complete else " (bi Stop giua chung — chua du)"))
        return idx

    async def _resolve_codes(self, cli: httpx.AsyncClient, codes: list[str]) -> dict[str, list[str]]:
        """Tang 1: ma nam TRONG SLUG (0 request phu, chi can sitemap da tai).
        Tang 2: ma khong khop slug -> CODE_INDEX (build neu chua co/het han).

        Tra ve {url: [ma1, ma2, ...]} — KHONG dedup ve 1-ma-1-url: nhieu ma khac
        nhau (ma he thong "sp-323" + ma hang "DCD805N-B1") co the tro CUNG 1 san
        pham. Dedup theo url se lam ma nhap SAU bi nuot, dong ket qua khong con
        truy duoc theo cot Input (bay da dinh o insumart.vn — AGENTS.md §10)."""
        self._unverified_codes.clear()
        urls_all = await self.sitemap_urls(cli)
        slug_entries = [(u, alias_from_url(u)) for u in urls_all]

        url_codes: dict[str, list[str]] = {}
        todo: list[str] = []
        short_codes = 0
        for c in codes:
            hit, match_kind = select_slug_candidate(c, slug_entries)
            if hit:
                url_codes.setdefault(hit, []).append(c)
            else:
                if match_kind == "short":
                    short_codes += 1
                elif match_kind.startswith("ambiguous"):
                    key = slugify_match(c)
                    candidates = [u for u, slug in slug_entries if key in slugify_match(slug)]
                    self.log(f"[WARN] Ma '{c}' co nhieu slug khop ({match_kind}): "
                             + ", ".join(candidates[:3]) + " -> xac minh bang CODE_INDEX")
                todo.append(c)
        matched = sum(len(v) for v in url_codes.values())
        if matched:
            self.log(f"[SYS] {matched}/{len(codes)} ma khop truc tiep tu slug (0 request phu)")
        if short_codes:
            self.log(f"[SYS] {short_codes} ma ngan (<{MIN_SLUG_CODE_LEN} ky tu) bo qua slug -> CODE_INDEX")
        if not todo:
            return url_codes

        idx = self._load_index()
        index_was_complete = self._index_complete
        if not self._index_complete:
            idx = await self.build_index(cli)
        still: list[str] = []
        for c in todo:
            hit = idx.get(norm_code(c))
            if hit:
                url_codes.setdefault(hit["url"], []).append(c)
            else:
                still.append(c)
        if still and index_was_complete:
            self.log(f"[SYS] {len(still)} ma cache miss -> xac minh catalog live (co the ~45 phut)")
            idx = await self.build_index(cli, force=True)
            unresolved: list[str] = []
            for c in still:
                hit = idx.get(norm_code(c))
                if hit:
                    url_codes.setdefault(hit["url"], []).append(c)
                else:
                    unresolved.append(c)
            still = unresolved
        if still:
            if self._index_complete:
                self.log(f"[MISS] {len(still)} ma khong tim thay sau khi quet catalog live: {', '.join(still[:10])}")
            else:
                self._unverified_codes.update(still)
                self.log(f"[WARN] {len(still)} ma CHUA XAC MINH do quet catalog bi dung/loi: {', '.join(still[:10])}")
        return url_codes

    async def _run_codes(self, cli: httpx.AsyncClient, key: str, codes: list[str]) -> None:
        """1 URL co the phuc vu NHIEU ma (xem _resolve_codes) -> fetch URL 1 LAN,
        nhan record ra cho TUNG ma nhap (khac _crawl_urls: 1 URL = 1 input)."""
        url_codes = await self._resolve_codes(cli, codes)
        matched_codes = {c for cs in url_codes.values() for c in cs}
        unmatched = [c for c in codes if c not in matched_codes]

        seen = self._load_ckpt(key)                 # theo URL da xong
        todo = {u: cs for u, cs in url_codes.items() if u not in seen}
        self.total += len(todo) + len(unmatched)
        self.done += len(seen)
        sem = asyncio.Semaphore(self.conc)

        async def one(u: str, cs: list[str]) -> None:
            async with sem:
                if not await self._gate():
                    return
                html = await self.fetch(cli, u)
                if self.delay:
                    await asyncio.sleep(self.delay)
                base = parse_page(html, u) if html else []
                if base:
                    self.index_records(base)
                    recs = [{**r, "Input": c} for c in cs for r in base]
                else:
                    recs = [{**not_found(c), "Input": c} for c in cs]
                self.results.extend(recs)
                self._save(key, recs)
                self.done += 1
                self.log(f"[{'OK ' if base else 'MISS'}] {self.done}/{self.total} "
                         f"+{len(recs) if base else 0} | {','.join(cs)}")
                if self.done % GC_EVERY == 0:
                    gc.collect()

        try:
            await asyncio.gather(*(one(u, cs) for u, cs in todo.items()))
        except (KeyboardInterrupt, Exception) as e:   # crash-proof: record da nam trong JSONL
            self.log(f"[FATAL] {type(e).__name__}: {e}")

        for c in unmatched:                          # het duong tim -> Not Found that
            if not any(r.get("Input") == c for r in self.results):
                status = "Unverified - catalog scan interrupted" if c in self._unverified_codes else "Not Found"
                rec = {**not_found(c, status), "Input": c}
                self.results.append(rec)
                self._save(key, [rec])
                self.done += 1
                self.log(f"[MISS] {self.done}/{self.total} +0 | {c}")

    async def _crawl_urls(self, cli: httpx.AsyncClient, urls: list[str], key: str,
                          inputs: dict[str, str] | None = None) -> None:
        """Link / crawl ALL: 1 URL = 1 input (khac code mode, khong can nhan ban)."""
        seen = self._load_ckpt(key)
        urls = [u for u in dict.fromkeys(urls) if u not in seen][:MAX_ITEMS]
        self.total += len(urls)
        self.done += len(seen)
        sem = asyncio.Semaphore(self.conc)

        async def one(u: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                html = await self.fetch(cli, u)
                if self.delay:
                    await asyncio.sleep(self.delay)
                recs = parse_page(html, u) if html else []
                inp = (inputs or {}).get(u, "")
                for r in recs:
                    r["Input"] = inp
                if recs:
                    self.results.extend(recs)
                    self._save(key, recs)
                    self.index_records(recs)
                else:
                    recs = [empty_page_record(u, inp, fetched=html is not None)]
                    self.results.extend(recs)
                    self._save(key, recs)
                self.done += 1
                self.log(f"[{'OK ' if recs else 'MISS'}] {self.done}/{self.total} "
                         f"+{len(recs)} | {u[-58:]}")
                if self.done % GC_EVERY == 0:
                    gc.collect()

        try:
            await asyncio.gather(*(one(u) for u in urls))
        except (KeyboardInterrupt, Exception) as e:   # crash-proof: record da nam trong JSONL
            self.log(f"[FATAL] {type(e).__name__}: {e}")

    async def run(self, mode: str, target: str = "", excel: str = "", sheet: str = "",
                  col: str = "", limit: int = 0) -> list[dict]:
        headers = HEADERS
        key = self.key = f"{mode}_{target or excel}_{col}"
        async with httpx.AsyncClient(headers=headers, timeout=TIMEOUT, follow_redirects=True,
                                     max_redirects=MAX_REDIRECTS,
                                     limits=httpx.Limits(max_connections=self.conc + 4)) as cli:
            if mode == "all":
                sitemap_urls = await self.sitemap_urls(cli)
                urls = sitemap_urls[:limit or MAX_ITEMS]
                self.log(f"[SYS] Crawl ALL: {len(urls)} trang san pham")
                await self._crawl_urls(cli, urls, key)
                if (limit == 0 and not self.stop_evt.is_set()
                        and not any(r.get("Status") == "Failed" for r in self.results)):
                    self._index_complete = True
                    self._index_done_urls = set(sitemap_urls)
                    self._index_progress_loaded = True
                self.save_index()

            elif mode in ("code", "excel"):
                raw = [norm_code(target)] if mode == "code" else read_excel_col(excel, sheet, col)
                raw = [c for c in dict.fromkeys(raw) if c][:MAX_ITEMS]
                links = [c for c in raw if c.lower().startswith("http")]
                codes = [c for c in raw if c not in links]
                # LINK crawl RIENG, KHONG ap dung ket qua code (1 trang = 1 SP o day
                # nen it rui ro hon toptul, nhung van tach cho nhat quan)
                if links:
                    await self._crawl_urls(cli, links, key, inputs={u: u for u in links})
                if codes:
                    await self._run_codes(cli, key, codes)

            elif mode == "keyword":
                idx = self._load_index()
                if not self._index_complete:
                    idx = await self.build_index(cli)
                kw = target.strip().lower()
                hits = [v["url"] for v in idx.values() if kw in (v.get("name") or "").lower()]
                hits = list(dict.fromkeys(hits))[:limit or MAX_ITEMS]
                self.log(f"[SYS] '{target}': {len(hits)} SP khop ten")
                await self._crawl_urls(cli, hits, key, inputs={u: target for u in hits})

            elif mode == "link":
                links = [u.strip() for u in target.splitlines() if u.strip().startswith("http")]
                await self._crawl_urls(cli, links, key, inputs={u: u for u in links})

            else:
                self.log(f"[ERR] mode khong hop le: {mode}")
        return self.results


def read_excel_col(path: str, sheet: str, col: str) -> list[str]:
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
    root.title("Nam Toan Viet Crawler — namtoanviet.vn")
    root.geometry("1080x680")
    ttk.Style().theme_use("clam")

    q: queue.Queue[str] = queue.Queue()
    stop_evt, pause_evt = threading.Event(), threading.Event()
    state: dict[str, Any] = {"running": False, "results": [], "eng": None}

    left = ttk.Frame(root, padding=10)
    left.pack(side="left", fill="y")
    right = ttk.Frame(root, padding=(0, 10, 10, 10))
    right.pack(side="right", fill="both", expand=True)

    mode = tk.StringVar(value="code")
    ttk.Label(left, text="Pipeline", font=("Segoe UI", 10, "bold")).pack(anchor="w")
    for val, txt in [("code", "1. Theo MA"), ("keyword", "2. Theo TU KHOA"),
                     ("link", "3. Theo LINK (moi dong 1 link)"), ("excel", "4. Theo cot EXCEL"),
                     ("all", "5. Crawl ALL (sitemap ~4.1k SP)")]:
        ttk.Radiobutton(left, text=txt, value=val, variable=mode).pack(anchor="w")

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
            asyncio.run(eng.run(m, target, xls.get(), sheet.get(), col.get(), limit.get()))
            eng.cleanup()
        except BaseException as e:      # BaseException: SystemExit/KeyboardInterrupt cung phai bat,
            q.put(f"[FATAL] {type(e).__name__}: {e}")   # khong thi thread chet cam, GUI treo
        finally:                        # finally: LUON nha co, khong khoa START vinh vien
            state["results"] = eng.results
            state["running"] = False
            q.put(f"[SYS] XONG — {len(eng.results)} record. Bam EXPORT de xuat Excel.")

    def start() -> None:
        if state["running"]:
            messagebox.showinfo("Dang chay", "Dang crawl. Bam Stop truoc.")
            return
        m = mode.get()
        if m == "excel" and not (xls.get() and col.get()):
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
        p = ask_export_path(xls.get(), HERE, "namtoanviet.xlsx")
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
            os._exit(0)      # kill cung: khong de thread asyncio/console treo

    root.protocol("WM_DELETE_WINDOW", on_close)
    drain()
    root.mainloop()


# ================================================================ CLI
def main() -> None:
    ap = argparse.ArgumentParser(description="Nam Toan Viet crawler (namtoanviet.vn)")
    ap.add_argument("--mode", choices=["code", "keyword", "link", "excel", "all"])
    ap.add_argument("--code")
    ap.add_argument("--keyword")
    ap.add_argument("--link")
    ap.add_argument("--excel")
    ap.add_argument("--sheet", default="")
    ap.add_argument("--col", default="")
    ap.add_argument("--out", default="namtoanviet.xlsx")
    ap.add_argument("--conc", type=int, default=8)
    ap.add_argument("--delay", type=float, default=0.05)
    ap.add_argument("--limit", type=int, default=0)
    a = ap.parse_args()

    if not a.mode:
        run_gui()
        return

    target = a.code or a.keyword or a.link or ""
    eng = Engine(a.conc, a.delay)
    t0 = time.time()
    try:
        asyncio.run(eng.run(a.mode, target, a.excel or "", a.sheet, a.col, a.limit))
    finally:
        if eng.results:
            p = export_excel(eng.results, pathlib.Path(a.out))
            print(f"[SYS] {len(eng.results)} record -> {p} ({time.time() - t0:.1f}s)")
        eng.cleanup()


if __name__ == "__main__":
    main()
