"""
insumart.vn (INSUMART - Tan Dia Cau) — httpx + selectolax, NO browser, NO Selenium.

RECON (2026-08-07)
------------------
Nen tang : **Sapo** (robots.txt tu khai "we use Sapo as our ecommerce platform",
           CDN bizweb.dktcdn.net). => co ho JSON endpoint chuan Sapo/Shopify.
API      : CO, JSON thuan, khong can login. **Toan bo du lieu nam trong list JSON**
           (khong phai fetch detail tung SP khi crawl ALL):
    /products.json?limit=250&page=N
        -> {"products":[{id, name, alias, vendor, product_type, price, price_min/max,
             tags[], summary(HTML), content(HTML co BANG SPEC), featured_image, images[],
             options[], variants[], available, url}]}
    /products/<alias>.json  (= .js, = /<alias>.json)
        -> 1 san pham, GIONG HET record trong list (chi them ban copy long nhau "product")
           => KHONG can goi detail neu da co list. Chi dung cho tra 1 ma / va SP lot.
    /collections.json?limit=250&page=N          -> 1.361 collection
    /collections/<alias>/products.json?limit=250&page=N   -> full record nhu tren
    /sitemap.xml -> sitemap_products_1..5.xml   -> 40.511 URL san pham (CHUAN VANG)

    !! BAY #1 — CAP OFFSET 30.000 (quan trong nhat, de am tham mat 1/4 catalog):
       page*limit > 30000 -> HTTP **500** (120x250 chet, 600x50 cung chet => la cap
       theo OFFSET chu khong phai theo so trang). Catalog THAT co 40.511 SP.
       => Paginate /products.json den het chi lay duoc 30.000 SP = **SOT ~10.500 SP (26%)**
          MA KHONG BAO LOI GI. Da do that: sitemap 40.511 vs products.json 30.000.
       => Cach chua (facet-partition, AGENTS.md 1a/Buoc 2): quet theo TUNG COLLECTION.
          Do that: 1.361 collection -> 40.509 SP unique trong ~70s (conc 8), chi con
          **2 SP** khong thuoc collection nao -> va bang sitemap (fetch .json truc tiep).
          Ket qua: 40.511/40.511 = du 100%.
    BAY #2: /products.json co trang NGAN GIUA CHUNG (page 42 tra 204 < 250) nhung page 43
       VAN CON SP MOI => TUYET DOI khong dung phan trang khi len(page) < limit. Chi dung
       khi trang RONG hoac KHONG CO ID MOI (stuck-page).
    BAY #3: cheatsheet AGENTS.md ghi "Sapo: /search?q=&view=json" — o site NAY view=json
       **KHONG co tac dung**, van tra HTML (da thu ca view=suggest/ajax_result/result_json,
       /search/suggest.json, /search.json, /apps/search -> 404 hoac HTML). Dung ban HTML.
search   : /search?q=<tu khoa>&page=N — HTML **render san o server** (khong phai JS).
           Tra dung ca khi tra theo **SKU** (IS-040132), **barcode** (IS-TDC-TBVTHC-040132)
           va **MPN** (TG-50) => code mode chi ton 2 request/ma, KHOI dung index 40k.
           30 SP/trang; tong ghi trong "Co N ket qua tim kiem phu hop".
    BAY #4: link san pham trong trang search phai LAY TRONG LUOI KET QUA
       (doan tu "title-page-search" den "pagination"), neu quet ca trang se dinh
       carousel/menu cua theme -> 171 "ket qua" trong khi that su chi co 30.
       Ket qua Sapo dung the: <a class="image_thumb ..." href="/<alias>">
spec     : nam trong **content** (HTML) duoi dang <table> 2 cot "Ten: Gia tri" (95% SP),
           va trong **summary** (100% SP) duoi dang <p><strong>Ten:</strong> Gia tri</p>
           (summary la tap CON: MPN/VAT/Xuat Xu/Chung Chi). => doc table truoc, thieu thi
           bu bang summary. Key spec THEO NGANH HANG (khao sat 500 SP: 34 key khac nhau,
           14 key co o moi SP, con lai rieng tung nhom: "Duong Kinh Cang Khoa (mm)"...)
           => GOP 1 COT "Thong so" theo AGENTS.md 4.6, KHONG trai thanh cot dong.
variants : **23% SP co NHIEU variant** (do 500 SP: 115 SP >1 variant, nhieu nhat 15).
           MOI variant co sku + barcode + gia + don vi + ton kho RIENG.
           => xuat **1 DONG / VARIANT**. Neu xuat 1 dong/SP se mat sach SKU cua cac
           variant con lai (day la cho de "bo sot" nhat cua site nay).
tags     : "donvi_<don vi>", "minorder_<so luong toi thieu>" + ma nhom noi bo (TBVTHC...).
gia      : cong khai (VND, da gom VAT), khong can dang nhap.
robots   : Allow tat ca tru /admin /cart /search /checkout /account. (Ta doc /search cho
           chuc nang tra ma theo yeu cau nguoi dung — tan suat thap, co delay + rate limit.)
charset  : utf-8.

Run
---
  insumart.bat                                         -> GUI
  python insumart.py --mode all [--limit 500]          -> collections + va sitemap (~40.5k SP)
  python insumart.py --mode code --code IS-040132
  python insumart.py --mode keyword --keyword loctite
  python insumart.py --mode link --link https://insumart.vn/loctite-henkel-577-50ml-va-250ml
  python insumart.py --mode excel --excel parts.xlsx --sheet Sheet1 --col Part
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
from typing import Any
from urllib.parse import quote

import httpx
import pandas as pd

HERE = pathlib.Path(__file__).parent
BASE = "https://insumart.vn"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
HEADERS = {"User-Agent": UA, "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8"}

PAGE_SIZE = 250          # tran cua Sapo (limit>250 bi ep ve 250)
SEARCH_PER_PAGE = 30     # so SP/trang cua trang /search
OFFSET_CAP = 30000       # BAY #1: page*limit > 30000 -> HTTP 500

CIRCUIT_BREAK = 25       # N loi LIEN TIEP -> nghi bi chan IP -> dung han
MAX_ITEMS = 60000        # tran work-list (catalog that ~40.5k)
MAX_PAGES = 600          # tran phan trang
MAX_REDIRECTS = 5
TIMEOUT = 60.0
GC_EVERY = 2000

TAG_RE = re.compile(r"<[^>]+>")
TR_RE = re.compile(r"<tr[^>]*>(.*?)</tr>", re.S | re.I)
TD_RE = re.compile(r"<t[dh][^>]*>(.*?)</t[dh]>", re.S | re.I)
# BAY #4: chi lay link trong luoi ket qua, va Sapo dung class image_thumb cho anh SP
SEARCH_ITEM_RE = re.compile(r'class="image_thumb[^"]*"\s+href="/([^"?#]+)"')
SEARCH_TOTAL_RE = re.compile(r"Có\s+(\d+)\s+kết quả")
LOC_RE = re.compile(r"<loc>([^<]+)</loc>")


# ================================================================ CORE (khong dung GUI)
def norm_code(s: Any) -> str:
    """Chuan hoa ma de so khop. URL thi GIU NGUYEN hoa/thuong (path phan biet chu)."""
    v = re.sub(r"\s+", "", str(s or ""))
    return v if v.lower().startswith("http") else v.upper()


def alias_from_url(u: str) -> str:
    """https://insumart.vn/<alias>[?variant=..] -> <alias>. Chap ca /products/<alias>."""
    a = re.sub(r"^https?://[^/]+", "", str(u or "").strip()).split("?")[0].split("#")[0]
    a = a.strip("/")
    if a.startswith("products/"):
        a = a[len("products/"):]
    return re.sub(r"\.(json|js)$", "", a)


def product_json_url(alias: str) -> str:
    return f"{BASE}/products/{quote(alias)}.json"


def search_url(term: str, page: int = 1) -> str:
    return f"{BASE}/search?q={quote(term)}&page={page}"


def strip_html(s: str) -> str:
    return htmlmod.unescape(TAG_RE.sub("", s or "")).replace("\xa0", " ").strip()


def spec_from_content(content: str) -> dict[str, str]:
    """Bang spec 2 cot trong `content` -> {ten: gia tri}. Bo dong khong du 2 o."""
    out: dict[str, str] = {}
    for tr in TR_RE.findall(content or ""):
        tds = TD_RE.findall(tr)
        if len(tds) < 2:
            continue
        k = strip_html(tds[0]).rstrip(":").strip()
        v = strip_html(tds[1])
        if k and k not in out:
            out[k] = v
    return out


def spec_from_summary(summary: str) -> dict[str, str]:
    """`summary` dang <p><strong>Ten:</strong> Gia tri</p> -> {ten: gia tri}.
    Chi nhan dong CO dau ':' trong phan <strong> (tranh nuot link/tieu de quang cao)."""
    out: dict[str, str] = {}
    for blk in re.split(r"</p>|<br\s*/?>", summary or "", flags=re.I):
        m = re.search(r"<strong>(.*?)</strong>(.*)", blk, re.S | re.I)
        if not m:
            continue
        k = strip_html(m.group(1))
        if ":" not in k:                     # khong phai cap ten:gia tri
            continue
        k = k.rstrip(":").strip()
        v = strip_html(m.group(2))
        if k and v and k not in out:
            out[k] = v
    return out


def build_spec(p: dict) -> dict[str, str]:
    spec = spec_from_content(p.get("content") or "")
    for k, v in spec_from_summary(p.get("summary") or "").items():
        spec.setdefault(k, v)                # table la nguon chinh, summary bu cho ~5% thieu
    return spec


def parse_tags(tags: list) -> tuple[str, str, list[str]]:
    """tags -> (don vi, so luong toi thieu, tags con lai)."""
    unit = minorder = ""
    rest: list[str] = []
    for t in tags or []:
        s = str(t)
        if s.startswith("donvi_"):
            unit = s[len("donvi_"):]
        elif s.startswith("minorder_"):
            minorder = s[len("minorder_"):]
        else:
            rest.append(s)
    return unit, minorder, rest


def parse_product(p: dict, input_code: str = "") -> list[dict]:
    """1 san pham -> **1 record / VARIANT** (moi variant co SKU + gia + ton rieng;
    xuat 1 dong/SP se mat sach SKU cua cac variant con lai)."""
    if not isinstance(p, dict) or not p.get("alias"):
        return []
    spec = build_spec(p)
    mpn = spec.get("MPN") or ""
    unit_tag, minorder, rest_tags = parse_tags(p.get("tags"))
    url = f"{BASE}/{p['alias']}"
    imgs = [i for i in (p.get("images") or []) if i]
    if not imgs and p.get("featured_image"):
        imgs = [p["featured_image"]]
    cats = [c for c in (p.get("vendor"), p.get("product_type")) if c]

    variants = p.get("variants") or [{}]
    out: list[dict] = []
    for v in variants:
        vimg = v.get("featured_image")
        vimg = vimg.get("src") if isinstance(vimg, dict) else vimg
        out.append({
            "Input": input_code,
            "Item No": v.get("sku") or "",
            "Page": p.get("name") or "",
            "Status": "OK",
            "_cats": cats,
            "_spec": spec,
            "Bien the": v.get("title") if len(variants) > 1 else "",
            "MPN": mpn,
            "Barcode": v.get("barcode") or "",
            "Don vi": v.get("unit") or unit_tag,
            "Gia (VND)": v.get("price", p.get("price")),
            "Gia so sanh": v.get("compare_at_price") or "",
            "Ton kho": v.get("inventory_quantity"),
            "Con hang": "Co" if (v.get("available", p.get("available"))) else "Het",
            "Dat toi thieu": minorder,
            "Can nang (g)": v.get("weight"),
            "Nhan": ", ".join(rest_tags),
            "Product ID": p.get("id"),
            "Variant ID": v.get("id"),
            "Image": " | ".join([i for i in ([vimg] if vimg else []) + imgs if i][:5]),
            "URL": f"{url}?variant={v['id']}" if v.get("id") and len(variants) > 1 else url,
        })
    return out


def not_found(code: str, status: str = "Not Found") -> dict:
    return {"Input": code, "Item No": code, "Page": "", "Status": status,
            "_cats": [], "_spec": {}, "Image": "", "URL": ""}


def match_code(recs: list[dict], code: str) -> list[dict]:
    """§4.3 + BAY §10-5: search tra nhieu hit -> KHONG lay mu ket qua dau.
    So ma nhap voi CA BA: SKU, barcode, MPN (site VN thuong co ma he thong + ma hang)."""
    c = norm_code(code)
    exact = [r for r in recs
             if c in (norm_code(r.get("Item No")), norm_code(r.get("Barcode")), norm_code(r.get("MPN")))]
    if exact:
        return exact
    sub = [r for r in recs
           if any(c in norm_code(r.get(k)) for k in ("Item No", "Barcode", "MPN"))]
    for r in sub:                       # lay nhung PHAI danh dau de nguoi dung biet ma lech
        r["Status"] = "OK (khop mot phan)"
    return sub


def export_excel(results: list[dict], path: pathlib.Path) -> pathlib.Path:
    """Spec gop 1 COT 'Thong so' — key spec khac nhau theo nganh hang, trai ra se no
    hang tram cot rong (AGENTS.md 4.6)."""
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

    fixed_head = ["Input", "Item No", "Page", "Status"]
    fixed_tail = ["Bien the", "MPN", "Barcode", "Don vi", "Gia (VND)", "Gia so sanh",
                  "Ton kho", "Con hang", "Dat toi thieu", "Can nang (g)", "Nhan",
                  "Product ID", "Variant ID"]
    max_cat = max((len(r.get("_cats") or []) for r in results), default=0)
    rows = []
    for r in results:
        row = {k: r.get(k, "") for k in fixed_head}
        cats = r.get("_cats") or []
        for i in range(max_cat):
            row[f"Category {i + 1}"] = cats[i] if i < len(cats) else ""
        for k in fixed_tail:
            row[k] = r.get(k, "")
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
        # 1 object Alignment dung chung cho ca cot: crawl ALL co ~57k dong, tao
        # Alignment moi cho tung o vua cham vua ton RAM (va copy() da deprecated).
        align = Alignment(wrap_text=True, vertical="top")
        for c in ws[col][1:]:
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
        self._ckpt_loaded = False        # BAY §10-3: khong nap checkpoint 2 lan

    # ---- resume: JSONL sidecar (append tung item; chay lai = doc lai, bo qua da xong)
    def ckpt(self, key: str) -> pathlib.Path:
        return HERE / f"_raw_{re.sub(r'[^A-Za-z0-9]+', '_', key)[:40]}.jsonl"

    def _load_ckpt(self, key: str) -> set[str]:
        """Tra ve tap ALIAS da xong (dedup theo alias — 1 SP = nhieu dong variant)."""
        if self._ckpt_loaded:
            return {alias_from_url(r.get("URL", "")) for r in self.results if r.get("URL")}
        self._ckpt_loaded = True
        p = self.ckpt(key)
        if not p.exists():
            return set()
        seen: set[str] = set()
        for line in p.read_text("utf-8", errors="ignore").splitlines():
            try:
                r = json.loads(line)
            except Exception:
                continue
            if r.get("Status") == "Not Found":
                # MISS co the do loi mang tam (search that bai) -> lan sau PHAI THU LAI,
                # khong duoc coi la "da xong" roi bao khong co hang vinh vien
                continue
            self.results.append(r)
            if r.get("URL"):
                seen.add(alias_from_url(r["URL"]))
        if seen:
            self.log(f"[SYS] Resume: {len(self.results)} dong / {len(seen)} SP da xong -> bo qua")
        return seen

    def _save(self, key: str, recs: list[dict]) -> None:
        with self.ckpt(key).open("a", encoding="utf-8") as f:
            for r in recs:
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
            f.flush()

    def cleanup(self) -> None:
        """Xong sach -> xoa checkpoint. Bi Stop/crash -> GIU de lan sau resume."""
        if not self.key:
            return
        p = self.ckpt(self.key)
        if p.exists() and not self.stop_evt.is_set():
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
                if r.status_code in (403, 429, 503):          # WAF / rate-limit -> lui dan
                    wait = min(2 ** i * 5, 60)
                    self.log(f"[WARN] {r.status_code} -> ngu {wait}s | {url[-70:]}")
                    await asyncio.sleep(wait)
                    continue
                if r.status_code in (404, 500):
                    # 500 o day = cham CAP OFFSET 30.000 (BAY #1), khong phai loi mang
                    return None
                r.raise_for_status()
                self.fails = 0
                return r.text
            except Exception as e:
                if i < tries - 1:
                    await asyncio.sleep(min(2 ** i, 20))
                    continue
                self.fails += 1
                self.log(f"[ERR] {type(e).__name__} | {url[-70:]}")
                if self.fails >= CIRCUIT_BREAK:
                    self.log(f"[STOP] {CIRCUIT_BREAK} loi LIEN TIEP - nghi bi chan IP. DUNG.")
                    self.stop_evt.set()
                return None
        return None

    async def fetch_json(self, cli: httpx.AsyncClient, url: str) -> Any:
        t = await self.fetch(cli, url)
        if not t:
            return None
        try:
            return json.loads(t)
        except Exception:
            return None

    # ---- phan trang JSON dung chung cho /products.json va /collections/<a>/products.json
    async def _paginate(self, cli: httpx.AsyncClient, path: str, field: str,
                        limit_items: int = 0, label: str = "") -> list[dict]:
        """BAY #2: KHONG duoc dung khi len(page) < limit (site co trang ngan giua chung).
        Chi dung khi: trang rong / khong co id moi (stuck) / cham cap offset 30.000."""
        out: list[dict] = []
        seen_ids: set[Any] = set()
        page = 1
        while page <= MAX_PAGES:
            if not await self._gate():
                break
            if page * PAGE_SIZE > OFFSET_CAP:
                self.log(f"[WARN] {label or path}: cham CAP OFFSET {OFFSET_CAP} -> dung phan trang "
                         f"(dung mode 'all' de quet theo collection cho du)")
                break
            d = await self.fetch_json(cli, f"{BASE}{path}?limit={PAGE_SIZE}&page={page}")
            items = (d or {}).get(field) or []
            if not items:
                break
            new = [x for x in items if x.get("id") not in seen_ids]
            if not new:                                  # stuck-page
                break
            seen_ids.update(x.get("id") for x in new)
            out.extend(new)
            if self.delay:
                await asyncio.sleep(self.delay)
            if limit_items and len(out) >= limit_items:
                return out[:limit_items]
            page += 1
        return out

    # ---- discovery cho mode ALL: collection-partition (pha CAP 30.000) + va bang sitemap
    async def _all_collections(self, cli: httpx.AsyncClient) -> list[str]:
        cols = await self._paginate(cli, "/collections.json", "collections", label="collections")
        return [c["alias"] for c in cols if c.get("alias")]

    async def _sitemap_aliases(self, cli: httpx.AsyncClient) -> set[str]:
        """sitemap.xml (sitemapindex) -> cac file con -> alias san pham.
        KHONG doan day so file (BAY §10-9): doc that tu sitemapindex."""
        idx = await self.fetch(cli, f"{BASE}/sitemap.xml")
        if not idx:
            return set()
        subs = [u for u in LOC_RE.findall(idx) if "sitemap_products" in u]
        sem = asyncio.Semaphore(self.conc)

        async def one(u: str) -> list[str]:
            async with sem:
                x = await self.fetch(cli, u)
                return LOC_RE.findall(x) if x else []

        urls = [u for part in await asyncio.gather(*(one(s) for s in subs)) for u in part]
        return {a for a in (alias_from_url(u) for u in urls) if a}

    async def _crawl_products(self, cli: httpx.AsyncClient, products: list[dict], key: str,
                              seen: set[str], inputs: dict[str, str] | None = None) -> None:
        """Ghi record tu cac dict san pham DA CO SAN (list JSON da du du lieu -> khong
        can goi them detail)."""
        for p in products:
            if not await self._gate():
                return
            a = p.get("alias") or ""
            if not a or a in seen:
                continue
            seen.add(a)
            recs = parse_product(p, (inputs or {}).get(a, ""))
            if recs:
                self.results.extend(recs)
                self._save(key, recs)
            self.done += 1
            if self.done % 500 == 0:
                self.log(f"[OK ] {self.done}/{self.total or '?'} SP | {len(self.results)} dong")
            if self.done % GC_EVERY == 0:
                gc.collect()

    async def _crawl_aliases(self, cli: httpx.AsyncClient, aliases: list[str], key: str,
                             seen: set[str], inputs: dict[str, str] | None = None) -> None:
        """Fetch /products/<alias>.json cho tung alias (link / keyword / va SP sot).
        1 alias -> 1 lan fetch, KHONG loc theo ma (mode code dung _fetch_by_codes)."""
        todo = [a for a in dict.fromkeys(aliases) if a and a not in seen][:MAX_ITEMS]
        self.total += len(todo)
        sem = asyncio.Semaphore(self.conc)

        async def one(a: str) -> None:
            async with sem:
                if not await self._gate():           # re-check SAU khi gianh slot
                    return
                d = await self.fetch_json(cli, product_json_url(a))
                if self.delay:
                    await asyncio.sleep(self.delay)
                recs = parse_product(d, (inputs or {}).get(a, "")) if d else []
                if recs:
                    self.results.extend(recs)
                    self._save(key, recs)
                seen.add(a)
                self.done += 1
                self.log(f"[{'OK ' if recs else 'MISS'}] {self.done}/{self.total} "
                         f"+{len(recs)} | {a[:58]}")

        try:
            await asyncio.gather(*(one(a) for a in todo))
        except (KeyboardInterrupt, Exception) as e:   # crash-proof: record da o trong JSONL
            self.log(f"[FATAL] {type(e).__name__}: {e}")

    async def search_aliases(self, cli: httpx.AsyncClient, term: str, limit: int = 0) -> list[str]:
        """/search?q= HTML -> alias trong LUOI KET QUA (BAY #4). Dung khi stuck-page."""
        out: list[str] = []
        seen: set[str] = set()
        page = 1
        total = None
        while page <= MAX_PAGES:
            if not await self._gate():
                break
            h = await self.fetch(cli, search_url(term, page))
            if not h:
                break
            if total is None:
                m = SEARCH_TOTAL_RE.search(h)
                total = int(m.group(1)) if m else -1
            i = h.find("title-page-search")
            if i < 0:
                break
            j = h.find("pagination", i)
            if j < 0:
                j = h.find("</section>", i)
            seg = h[i:j if j > 0 else len(h)]
            new = [a for a in dict.fromkeys(SEARCH_ITEM_RE.findall(seg)) if a not in seen]
            if not new:                                   # stuck-page = het ket qua
                break
            seen.update(new)
            out.extend(new)
            if self.delay:
                await asyncio.sleep(self.delay)
            if limit and len(out) >= limit:
                return out[:limit]
            if total and total > 0 and len(out) >= total:
                break
            page += 1
        return out

    async def run(self, mode: str, target: str = "", excel: str = "", sheet: str = "",
                  col: str = "", limit: int = 0) -> list[dict]:
        key = self.key = f"{mode}_{target or excel}_{col}"
        async with httpx.AsyncClient(headers=HEADERS, timeout=TIMEOUT, follow_redirects=True,
                                     max_redirects=MAX_REDIRECTS,
                                     limits=httpx.Limits(max_connections=self.conc + 4)) as cli:
            seen = self._load_ckpt(key)

            if mode == "all":
                await self._run_all(cli, key, seen, limit)

            elif mode in ("code", "excel"):
                raw = [norm_code(target)] if mode == "code" else read_excel_col(excel, sheet, col)
                raw = [c for c in dict.fromkeys(raw) if c][:MAX_ITEMS]
                links = [c for c in raw if c.lower().startswith("http")]
                codes = [c for c in raw if c not in links]
                # BAY §10-2: LINK crawl RIENG, KHONG ap bo loc `codes` (Item No cua link
                # dau co nam trong danh sach ma -> se bi xoa sach record)
                if links:
                    al = [alias_from_url(u) for u in links]
                    await self._crawl_aliases(cli, al, key, seen,
                                              inputs={a: u for a, u in zip(al, links)})
                if codes:
                    await self._run_codes(cli, key, seen, codes)

            elif mode == "keyword":
                al = await self.search_aliases(cli, target, limit)
                self.log(f"[SYS] '{target}': {len(al)} san pham")
                await self._crawl_aliases(cli, al, key, seen, inputs={a: target for a in al})

            elif mode == "link":
                links = [u.strip() for u in target.splitlines() if u.strip().startswith("http")]
                al = [alias_from_url(u) for u in links]
                await self._crawl_aliases(cli, al, key, seen,
                                          inputs={a: u for a, u in zip(al, links)})

            else:
                self.log(f"[ERR] mode khong hop le: {mode}")
        return self.results

    async def _run_all(self, cli: httpx.AsyncClient, key: str, seen: set[str], limit: int) -> None:
        """CAP OFFSET 30.000 (BAY #1) => KHONG paginate /products.json. Quet theo tung
        collection roi DOI CHIEU sitemap de va SP khong thuoc collection nao."""
        cols = await self._all_collections(cli)
        self.log(f"[SYS] {len(cols)} collection -> quet song song (pha CAP {OFFSET_CAP})")
        sem = asyncio.Semaphore(self.conc)
        lock = asyncio.Lock()
        ncol = 0

        async def one(alias: str) -> None:
            nonlocal ncol
            async with sem:
                if not await self._gate():
                    return
                ps = await self._paginate(cli, f"/collections/{quote(alias)}/products.json",
                                          "products", label=alias)
                async with lock:                      # ghi record noi tiep -> khong tranh chap
                    await self._crawl_products(cli, ps, key, seen)
                    ncol += 1
                    if ncol % 100 == 0:
                        self.log(f"[SYS] {ncol}/{len(cols)} collection | {len(seen)} SP "
                                 f"| {len(self.results)} dong")

        await asyncio.gather(*(one(a) for a in cols))
        self.log(f"[SYS] Collection xong: {len(seen)} SP | {len(self.results)} dong")

        if self.stop_evt.is_set():
            return
        # DOI CHIEU sitemap (chuan vang): SP nao khong thuoc collection nao -> fetch le
        sm = await self._sitemap_aliases(cli)
        missing = [a for a in sm if a not in seen]
        self.log(f"[SYS] Sitemap {len(sm)} SP -> con thieu {len(missing)} -> va truc tiep")
        if missing:
            await self._crawl_aliases(cli, missing, key, seen)
        if limit:
            self.results = self.results[:limit]

    async def _run_codes(self, cli: httpx.AsyncClient, key: str, seen: set[str],
                         codes: list[str]) -> None:
        """1 ma -> /search?q=<ma> (tra dung ca SKU/barcode/MPN) -> alias -> detail JSON
        -> KHOP MA CHINH XAC (§4.3). Ma miss van co dong 'Not Found'.

        Resume/dedup o mode nay tinh theo **MA NHAP** (khong theo alias): NHIEU MA khac
        nhau co the tro CUNG 1 san pham (vd SKU 'IS-040132' + barcode
        'IS-TDC-TBVTHC-040132' + MPN 'TG-50' deu ra 1 SP). Neu dedup theo alias thi chi
        ma DAU TIEN co dong ket qua, cac ma sau bi nuot -> nguoi dung do Excel theo cot
        Input se tuong la "khong tim thay". Nen: fetch alias 1 LAN (cache), nhung PHAT
        SINH DONG RIENG cho TUNG ma."""
        done_inputs = {norm_code(r.get("Input")) for r in self.results if r.get("Input")}
        todo = [c for c in codes if norm_code(c) not in done_inputs]
        if len(todo) < len(codes):
            self.log(f"[SYS] Resume: bo qua {len(codes) - len(todo)} ma da xong")
        self.total += len(todo)

        sem = asyncio.Semaphore(self.conc)
        cache: dict[str, Any] = {}          # alias -> product json (fetch 1 lan duy nhat)
        cache_lock = asyncio.Lock()

        async def get_product(a: str) -> Any:
            async with cache_lock:
                if a in cache:
                    return cache[a]
            d = await self.fetch_json(cli, product_json_url(a))
            async with cache_lock:
                cache[a] = d
            return d

        async def one(c: str) -> None:
            async with sem:
                if not await self._gate():           # re-check SAU khi gianh slot
                    return
                aliases = await self.search_aliases(cli, c, limit=SEARCH_PER_PAGE)
                if self.delay:
                    await asyncio.sleep(self.delay)
                recs: list[dict] = []
                for a in aliases:
                    d = await get_product(a)
                    if not d:
                        continue
                    recs += match_code(parse_product(d, c), c)
                if not recs:
                    recs = [not_found(c)]
                self.results.extend(recs)
                self._save(key, recs)
                seen.update(aliases)
                self.done += 1
                ok = recs[0].get("Status") != "Not Found"
                self.log(f"[{'OK ' if ok else 'MISS'}] {self.done}/{self.total} "
                         f"+{len(recs) if ok else 0} | {c}")

        try:
            await asyncio.gather(*(one(c) for c in todo))
        except (KeyboardInterrupt, Exception) as e:   # crash-proof: record da o trong JSONL
            self.log(f"[FATAL] {type(e).__name__}: {e}")


def read_excel_col(path: str, sheet: str, col: str) -> list[str]:
    """Loi -> ValueError (KHONG SystemExit: no la BaseException, worker GUI bat
    `except Exception` se KHONG bat duoc -> thread chet cam, GUI treo running=True)."""
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

    # ha tang GUI chung o root (di nguoc len cay thu muc -> song sot khi doi nhom folder)
    root_dir = next(p for p in (HERE, *HERE.parents) if (p / "gui_common.py").exists())
    sys.path.insert(0, str(root_dir))
    from gui_common import ask_export_path, attach_excel_picker

    root = tk.Tk()
    root.title("INSUMART Crawler — insumart.vn")
    root.geometry("1080x700")
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
    for val, txt in [("code", "1. Theo MA (SKU / barcode / MPN)"),
                     ("keyword", "2. Theo TU KHOA"),
                     ("link", "3. Theo LINK (moi dong 1 link)"),
                     ("excel", "4. Theo cot EXCEL (ma hoac link)"),
                     ("all", "5. Crawl ALL (~40.5k SP, ~2-4 phut)")]:
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
            stat.config(text=f"{eng.done}/{eng.total} SP | {len(eng.results)} dong")
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
            q.put(f"[SYS] XONG — {len(eng.results)} dong. Bam EXPORT de xuat Excel.")

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
        p = ask_export_path(xls.get(), HERE, "insumart.xlsx")
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
    ap = argparse.ArgumentParser(description="INSUMART crawler (insumart.vn)")
    ap.add_argument("--mode", choices=["code", "keyword", "link", "excel", "all"])
    ap.add_argument("--code")
    ap.add_argument("--keyword")
    ap.add_argument("--link")
    ap.add_argument("--excel")
    ap.add_argument("--sheet", default="")
    ap.add_argument("--col", default="")
    ap.add_argument("--out", default="insumart.xlsx")
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
            print(f"[SYS] {len(eng.results)} dong -> {p} ({time.time() - t0:.1f}s)")
        eng.cleanup()


if __name__ == "__main__":
    main()
