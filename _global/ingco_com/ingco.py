"""
ingco.com (INGCO Vietnam, www.ingco.com/vn) — httpx, NO browser, NO Selenium.

RECON (2026-08-07)
-------------------
API      : KHONG co API rieng biet cong khai. Site la React SPA (umi 3.5.41, "togroup"
           CMS) SSR full du lieu trang vao 1 the <script>:
               window.g_initialProps = { ...,"allCategory":[...], "data": {...} }
           -> doc JSON nhung, KHOI parse HTML. (kieu __NEXT_DATA__ cua DigiKey, nhung
           key la "data" o TOP-LEVEL cua g_initialProps, KHONG nam trong reqConfig).
    BAY #1 (quan trong nhat): curl voi User-Agent don gian ("Mozilla/5.0") ma KHONG co
    header Accept -> server tra ve trang SSR "rut gon" (thieu het "data"). PHAI gui
    Accept: text/html... + Accept-Language giong browser that thi "data" moi day du.
    BAY #1b (QUAN TRONG — anh huong toc do that): server backend site nay THUC SU
    KHONG ON DINH — quan sat thuc te (2026-08-07): crawl 350 SP, ty le loi SSR (200 OK
    nhung "data" khong thay HOAC tra nham "data":null cho ca ma THAT) dao dong 25-75%
    MOI LAN CHAY, khong on dinh theo thoi gian/concurrency ro rang (da do conc=1..8, cai
    thien it). Vi crawler khong the sua server nguoi ta, thiet ke buoc vong bang 2 lop:
      1. fetch_data() retry NGAY tai cho (10 lan, backoff 0.5-4s) truoc khi bo cuoc.
      2. Van con that bai -> ghi Status="Failed" (KHAC "Not Found" that — that su chi
         gan khi biet chac SP khong ton tai, vd tra ve "data":null cho ma KHONG nam
         trong category-index). "Failed" KHONG duoc tinh la "da xong" trong checkpoint
         -> CHAY LAI CUNG LENH (resume) se TU DONG thu lai CHI cac ma Failed, giu nguyen
         cac ma da OK. Quan sat thuc te: 3 lan chay lien tiep --mode all hoi tu
         90 -> 145 -> 184 OK / 350 (khong tuyen tinh, co the can 5-8 lan chay moi het
         Failed). NGUOI DUNG: neu thay nhieu "Failed", cu CHAY LAI y nguyen lenh vai lan.
listing  : GET /vn/products/<slug>?pageSize=2000
           -> data.list[] (productNo, productName, pathReplacement, keydata1-3, price,
              productThumbnail...) + data.total.
    BAY #2: slug SAI/khong ton tai (vd "hardware","generators" - doan tu sitemap GLOBAL
    khong prefix /vn) KHONG loi, KHONG rong -> IM LANG tra ve fallback total=350 (= TONG
    CA CATALOG). Vietnam site CHI CO 6 category that (xem allCategory), tong 350 SP.
    Chi tin cac slug lay tu "allCategory" cua chinh site, KHONG doan tu sitemap goc.
detail   : GET /vn/product/<slug>/<PRODUCTNO>  -> data = 1 product day du (parameter =
           HTML spec full, keydata1-3 = 3 thong so noi bat, productPicList, price...).
    PHAT HIEN HAY: <slug> chi la trang tri SEO — server resolve THEO PRODUCTNO, slug SAI
    van ra dung san pham (da test cheo). -> tra ma TRUC TIEP KHONG can biet slug that:
    GET /vn/product/x/<CODE>, roi doc pathReplacement THAT trong response de ghi URL
    chuan vao Excel. (giong Anex/Tsunoda "slug = ma -> direct lookup", o day con manh
    hon vi slug la wildcard).
    San pham/slug khong ton tai -> data = null (JSON) -> "Not Found" sach, KHONG redirect.
sitemap  : /sitemap.xml (domain goc, KHONG /vn) that, moi (512 URL /product/<slug>/<no>)
           nhung la catalog TOAN CAU (nhieu nuoc), KHONG khop voi catalog rieng /vn (350
           SP, 6 category). -> dung "allCategory" cua /vn lam nguon that cho VN, KHONG
           dung sitemap global de doan slug (xem BAY #2).
gia      : co gia VND cong khai (originalPrice/price), khong can dang nhap.
search   : /vn/products?search=<tu>  -> KHONG loc gi ca, luon tra fallback total=350
           giong slug sai (BAY #2 phien ban khac) -> KHONG dung endpoint nay. Keyword
           mode o day tu lam: crawl category-index (6 category, ret ve productName) roi
           loc substring tren client.
charset  : utf-8, khong quirk.

Run
---
  ingco.bat                                        -> GUI
  python ingco.py --mode all [--limit 500]         -> quet toan bo 6 category VN (~350 SP)
  python ingco.py --mode code --code CIRLI1201
  python ingco.py --mode keyword --keyword "may khoan"
  python ingco.py --mode link --link https://www.ingco.com/vn/product/cordless-impact-driver/CIRLI1201
  python ingco.py --mode excel --excel parts.xlsx --sheet Sheet1 --col Part

  Server flaky (xem BAY #1b) -> neu Excel co nhieu dong Status="Failed", CHAY LAI Y
  NGUYEN lenh vua roi — resume tu dong bo qua ma da OK, chi thu lai ma Failed.
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
BASE = "https://www.ingco.com"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
HEADERS = {
    "User-Agent": UA,
    # BAY #1: thieu Accept -> server tra trang SSR rut gon, mat het "data". BAT BUOC.
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
}

CAT_INDEX = HERE / "_category_index.json"   # productNo -> {name, slug, category} (6 category VN)
INDEX_TTL_DAYS = 7          # catalog nho (350 SP) nhung van dat TTL, tranh cache mu (§10 bay #8)
PAGE_SIZE = 2000            # 1 request lay HET 1 category (tong catalog VN chi ~350 SP)

CIRCUIT_BREAK = 25
MAX_ITEMS = 5000
MAX_REDIRECTS = 5
TIMEOUT = 30.0
GC_EVERY = 200

BR_RE = re.compile(r"<br\s*/?>", re.I)
TAG_RE = re.compile(r"<[^>]+>")


# ================================================================ CORE (khong dung GUI)
def norm_code(s: Any) -> str:
    v = re.sub(r"\s+", "", str(s or ""))
    return v if v.lower().startswith("http") else v.upper()


def detail_url(code: str, slug: str = "") -> str:
    """Slug chi trang tri SEO — server resolve theo CODE, slug sai/rong van ra dung SP."""
    return f"{BASE}/vn/product/{quote(slug or code.lower())}/{quote(code)}"


def category_url(slug: str, page_size: int = PAGE_SIZE) -> str:
    return f"{BASE}/vn/products/{quote(slug)}?pageSize={page_size}"


def extract_json_key(text: str, key: str, start: int = 0, rfind: bool = False) -> tuple[str | None, int]:
    """Trich JSON con tro key "KEY": {...}/[...] tu blob HTML bang can bang ngoac
    (KHONG the dung regex "</script>" - g_initialProps chua ca route-config voi function,
    khong phai JSON thuan; can tim dung vi tri roi can bang rieng tung nhanh)."""
    marker = f'"{key}":'
    i = text.rfind(marker) if rfind else text.find(marker, start)
    if i == -1:
        return None, -1
    p = i + len(marker)
    while p < len(text) and text[p] in " \n\r\t":
        p += 1
    if p >= len(text) or text[p] not in "{[":
        return None, -1
    open_c, close_c = ("{", "}") if text[p] == "{" else ("[", "]")
    depth, in_str, esc, bs = 0, False, False, chr(92)
    j = p
    while j < len(text):
        c = text[j]
        if in_str:
            if esc:
                esc = False
            elif c == bs:
                esc = True
            elif c == '"':
                in_str = False
        else:
            if c == '"':
                in_str = True
            elif c == open_c:
                depth += 1
            elif c == close_c:
                depth -= 1
                if depth == 0:
                    return text[p:j + 1], j + 1
        j += 1
    return None, -1


def extract_page_data(html: str) -> tuple[Any, bool]:
    """window.g_initialProps.data (top-level, SAU allCategory neu co) -> (value, found).
    "data" xuat hien >=2 lan trong blob (reqHeadFoot.data la cai KHAC, sai) -> lay lan
    CUOI (rfind) moi dung cai chua noi dung trang (list SP hoac 1 SP).
    found=False = SSR rong/hong (key "data" khong thay) -> DANG retry (xem fetch_data).
    found=True, value=None = server tra that "data":null = SP/slug that su khong ton tai."""
    if "g_initialProps" not in html:
        return None, False
    raw, _ = extract_json_key(html, "data", rfind=True)
    if not raw:
        return None, False
    try:
        return json.loads(raw), True
    except Exception:
        return None, False


def extract_all_category(html: str) -> list[dict]:
    raw, _ = extract_json_key(html, "allCategory")
    if not raw:
        return []
    try:
        return json.loads(raw)
    except Exception:
        return []


def clean_spec(parameter: str | None) -> str:
    """"parameter" la HTML (<br> ngan dong) -> text nhieu dong, unescape entity."""
    if not parameter:
        return ""
    t = BR_RE.sub("\n", parameter)
    t = TAG_RE.sub("", t)
    t = htmlmod.unescape(t)
    lines = [ln.strip() for ln in t.split("\n")]
    return "\n".join(ln for ln in lines if ln)


def parse_detail(data: dict, input_code: str = "") -> dict:
    code = data.get("productNo") or data.get("productExtNo") or ""
    pics = data.get("productPicList") or []
    if not pics and data.get("productThumbnail"):
        pics = [data["productThumbnail"]]
    spec = clean_spec(data.get("parameter"))
    if not spec:  # trang thieu "parameter" -> fallback 3 thong so noi bat
        spec = "\n".join(f"{v}" for v in
                         (data.get("keydata1"), data.get("keydata2"), data.get("keydata3")) if v)
    price = data.get("price") if data.get("price") is not None else data.get("originalPrice")
    return {
        "Input": input_code, "Item No": code,
        "Page": data.get("productName") or data.get("productExtName") or "",
        "Status": "OK",
        "_cats": [data.get("categoryName")] if data.get("categoryName") else [],
        "_spec": spec, "Price": price,
        "Image": " | ".join(pics[:5]),
        "URL": detail_url(code, data.get("pathReplacement") or ""),
    }


def not_found(code: str, status: str = "Not Found") -> dict:
    """status="Failed" (khac "Not Found" that) khi server flaky het budget retry ma
    khong xac dinh duoc — xem fetch_data(). _load_ckpt() KHONG tinh "Failed" la da xong
    -> resume se tu dong thu lai, "Not Found" that thi giu nguyen (khong thu lai vo ich)."""
    return {"Input": code, "Item No": code, "Page": "", "Status": status,
            "_cats": [], "_spec": "", "Price": None, "Image": "", "URL": ""}


def export_excel(results: list[dict], path: pathlib.Path) -> pathlib.Path:
    """Spec gop 1 COT 'Thong so' — khong trai union key thanh hang tram cot rong (§4.6)."""
    from openpyxl.utils import get_column_letter

    max_cat = max((len(r.get("_cats") or []) for r in results), default=0)
    rows = []
    for r in results:
        row = {"Input": r.get("Input", ""), "Item No": r.get("Item No", ""),
               "Page": r.get("Page", ""), "Status": r.get("Status", "")}
        cats = r.get("_cats") or []
        for i in range(max_cat):
            row[f"Category {i + 1}"] = cats[i] if i < len(cats) else ""
        row["Gia (VND)"] = r.get("Price")
        row["Thong so"] = r.get("_spec") or ""
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
    """Crawler core — chay headless/CLI duoc, khong phu thuoc GUI."""

    def __init__(self, conc: int = 3, delay: float = 0.05, log=print,
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
        self._cat_index: dict[str, dict] | None = None

    # ---- resume: JSONL sidecar
    def ckpt(self, key: str) -> pathlib.Path:
        return HERE / f"_raw_{re.sub(r'[^A-Za-z0-9]+', '_', key)[:40]}.jsonl"

    def _load_ckpt(self, key: str) -> set[str]:
        """"Failed" (server flaky, khac "Not Found" that — xem not_found()) KHONG duoc
        tinh la da xong: bi LOAI khoi self.results va khoi `seen` de duoc crawl lai tu
        dau o lan resume nay, tranh dong bang ket qua sai vi bi ket qua chua thi lai."""
        p = self.ckpt(key)
        if self._ckpt_loaded:
            return {r.get("Item No", "") for r in self.results if r.get("Status") != "Failed"}
        self._ckpt_loaded = True
        seen: set[str] = set()
        if not p.exists():
            return seen
        failed = 0
        for line in p.read_text("utf-8", errors="ignore").splitlines():
            try:
                r = json.loads(line)
            except Exception:
                continue
            if r.get("Status") == "Failed":
                failed += 1
                continue
            self.results.append(r)
            seen.add(r.get("Item No", ""))
        if seen or failed:
            self.log(f"[SYS] Resume: {len(self.results)} record da xong -> bo qua"
                     + (f" ({failed} 'Failed' se thu lai)" if failed else ""))
        return seen

    def _save(self, key: str, recs: list[dict]) -> None:
        with self.ckpt(key).open("a", encoding="utf-8") as f:
            for r in recs:
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
            f.flush()

    def cleanup(self) -> None:
        """Xong sach (khong Stop, khong con "Failed" do server flaky) -> xoa checkpoint.
        Con "Failed" -> GIU checkpoint de resume tu dong thu lai cac ma do lan sau."""
        if not self.key:
            return
        p = self.ckpt(self.key)
        has_failed = any(r.get("Status") == "Failed" for r in self.results)
        if has_failed:
            self.log(f"[SYS] Con {sum(1 for r in self.results if r.get('Status') == 'Failed')} "
                     f"'Failed' (server flaky) -> GIU checkpoint, chay LAI lenh nay se tu thu tiep.")
        if p.exists() and not self.stop_evt.is_set() and not has_failed:
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
                    self.log(f"[WARN] {r.status_code} -> ngu {wait}s | {url[-70:]}")
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
                self.log(f"[ERR] {type(e).__name__} | {url[-70:]}")
                if self.fails >= CIRCUIT_BREAK:
                    self.log(f"[STOP] {CIRCUIT_BREAK} loi LIEN TIEP - nghi bi chan IP. DUNG.")
                    self.stop_evt.set()
                return None
        return None

    # ---- category index: 6 category That cua VN (allCategory), 350 SP, TTL 7 ngay
    def _load_index(self) -> dict[str, dict]:
        if self._cat_index is not None:
            return self._cat_index
        self._cat_index = {}
        if CAT_INDEX.exists() and (time.time() - CAT_INDEX.stat().st_mtime) / 86400 < INDEX_TTL_DAYS:
            try:
                self._cat_index = json.loads(CAT_INDEX.read_text("utf-8"))
                self.log(f"[SYS] Category-index: {len(self._cat_index)} SP da biet (cache)")
            except Exception:
                self._cat_index = {}
        return self._cat_index

    def save_index(self) -> None:
        if self._cat_index:
            CAT_INDEX.write_text(json.dumps(self._cat_index, ensure_ascii=False), "utf-8")

    async def fetch_data(self, cli: httpx.AsyncClient, url: str, tries: int = 10) -> tuple[Any, bool]:
        """BAY RIENG cua site nay: server backend KHONG ON DINH — ty le SSR rong (200 OK
        nhung "data" khong thay) co luc rat cao (quan sat thuc te: co thoi diem >90% du
        da retry), khong hoan toan do concurrency cua crawler (da do o conc=1..8, van
        flaky). -> PHAI retry o TANG PARSE (khong chi tang HTTP) VOI budget lon, VA phan
        biet ro 2 that bai KHAC NHAU cho phia goi:
          - ok=True,  data=None  : server tra that "data":null -> SP/slug KHONG TON TAI that.
          - ok=False, data=None  : het budget retry ma KHONG LAN NAO xac dinh duoc (server
            flaky) -> KHONG PHAI Not Found that, phia goi phai danh dau "Failed" (khac
            "Not Found") de LAN CHAY SAU (resume) tu dong thu lai, khong bi ket qua sai
            dong bang vinh vien trong checkpoint."""
        for i in range(tries):
            html = await self.fetch(cli, url)
            if not html:
                await asyncio.sleep(min(0.5 * (i + 1), 4.0))
                continue
            data, found = extract_page_data(html)
            if found:
                return data, True
            await asyncio.sleep(min(0.5 * (i + 1), 4.0))
        return None, False

    async def _discover_categories(self, cli: httpx.AsyncClient) -> list[dict]:
        for i in range(4):
            html = await self.fetch(cli, f"{BASE}/vn/products/cordless-tools")
            cats = extract_all_category(html) if html else []
            if cats:
                return cats
            await asyncio.sleep(0.3 * (i + 1))
        self.log("[ERR] Khong doc duoc danh sach category VN (allCategory rong).")
        return []

    async def build_index(self, cli: httpx.AsyncClient) -> dict[str, dict]:
        """Quet 6 category That cua VN (KHONG doan slug tu sitemap global — xem BAY #2
        trong docstring) -> productNo -> {name, slug, category}."""
        idx = self._load_index()
        cats = await self._discover_categories(cli)
        self.log(f"[SYS] {len(cats)} category VN -> quet toan bo (~350 SP, 1 request/category)")
        sem = asyncio.Semaphore(self.conc)
        incomplete: list[str] = []

        async def one(cat: dict) -> None:
            async with sem:
                if not await self._gate():
                    return
                slug = cat.get("pathReplacement") or ""
                data, items = None, []
                # tat ca 6 category VN luon >=1 SP (xem recon) -> "list" rong = SSR
                # tra dict "hop le" nhung noi dung rong (khac case "data" khong thay
                # ma fetch_data() da retry) -> RETRY CA REQUEST o tang nay.
                for attempt in range(3):
                    data, _ = await self.fetch_data(cli, category_url(slug))
                    items = (data or {}).get("list") or []
                    if items:
                        break
                    await asyncio.sleep(0.8 * (attempt + 1))
                expected = (data or {}).get("total")
                if not items:
                    incomplete.append(slug)   # het budget retry -> DUNG cache lai
                for it in items:
                    no = it.get("productNo")
                    if no:
                        idx[no] = {"name": it.get("productName") or it.get("productExtName") or "",
                                   "slug": it.get("pathReplacement") or "",
                                   "category": cat.get("title") or slug}
                self.log(f"[SYS] {slug}: +{len(items)} SP (total bao {expected})")

        await asyncio.gather(*(one(c) for c in cats))
        if incomplete:
            self.log(f"[WARN] Category-index THIEU ({', '.join(incomplete)}) - KHONG luu cache, "
                     f"lan sau se quet lai.")
        else:
            self.save_index()
        self.log(f"[SYS] Category-index xong: {len(idx)} SP")
        return idx

    async def _crawl_codes(self, cli: httpx.AsyncClient, codes: list[str], key: str,
                           slugs: dict[str, str] | None = None,
                           inputs: dict[str, str] | None = None,
                           known_valid: bool = False) -> None:
        """known_valid=True (goi tu 'all'/'keyword', ma lay tu category-index that) ->
        server flaky co the tra "data":null NGAY CA VOI MA THAT (khong chi thieu key SSR)
        -> KHONG duoc tin "null" la Not Found that trong truong hop nay, phai coi la
        "Failed" de resume thu lai. code/excel (nguoi dung tu go, khong doi chieu) van
        giu "data":null = Not Found that (khong co gi de doi chieu them)."""
        seen = self._load_ckpt(key)
        codes = [c for c in dict.fromkeys(codes) if c not in seen][:MAX_ITEMS]
        self.total = len(codes) + len(seen)
        self.done = len(seen)
        sem = asyncio.Semaphore(self.conc)

        async def one(code: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                slug = (slugs or {}).get(code, "")
                data, ok = await self.fetch_data(cli, detail_url(code, slug))
                if self.delay:
                    await asyncio.sleep(self.delay)
                if data and (data.get("productNo") or "").upper() == code.upper():
                    rec = parse_detail(data, (inputs or {}).get(code, code))
                elif ok and not known_valid:
                    rec = not_found((inputs or {}).get(code, code))
                    rec["Item No"] = code
                else:
                    rec = not_found((inputs or {}).get(code, code), status="Failed")
                    rec["Item No"] = code
                self.results.append(rec)
                self._save(key, [rec])
                self.done += 1
                tag = {"OK": "OK ", "Not Found": "MISS", "Failed": "FAIL"}[rec["Status"]]
                self.log(f"[{tag}] {self.done}/{self.total} | {code}")
                if self.done % GC_EVERY == 0:
                    gc.collect()

        try:
            await asyncio.gather(*(one(c) for c in codes))
        except (KeyboardInterrupt, Exception) as e:
            self.log(f"[FATAL] {type(e).__name__}: {e}")

    async def run(self, mode: str, target: str = "", excel: str = "", sheet: str = "",
                  col: str = "", limit: int = 0) -> list[dict]:
        key = self.key = f"{mode}_{target or excel}_{col}"
        # DA DO (AGENTS.md 5.1): server SSR nay flaky manh duoi tai — conc cao (8) ->
        # ty le SSR-rong TANG VOT (~90% MISS du da retry). max_keepalive_connections=0
        # tranh dinh 1 ket noi vao 1 backend instance hong (retry ma van no lai chinh
        # no thi vo ich). fetch_data() da co retry rieng (10 lan) bu lai phan flaky con lai.
        limits = httpx.Limits(max_connections=self.conc + 2, max_keepalive_connections=0)
        async with httpx.AsyncClient(headers=HEADERS, timeout=TIMEOUT, follow_redirects=True,
                                     max_redirects=MAX_REDIRECTS, limits=limits) as cli:
            if mode == "all":
                idx = await self.build_index(cli)
                codes = list(idx.keys())[:limit or MAX_ITEMS]
                slugs = {c: idx[c]["slug"] for c in codes}
                self.log(f"[SYS] Crawl ALL: {len(codes)} SP")
                await self._crawl_codes(cli, codes, key, slugs=slugs, known_valid=True)

            elif mode in ("code", "excel"):
                raw = [norm_code(target)] if mode == "code" else read_excel_col(excel, sheet, col)
                raw = [c for c in dict.fromkeys(raw) if c][:MAX_ITEMS]
                links = [c for c in raw if c.lower().startswith("http")]
                codes = [c for c in raw if c not in links]
                if links:
                    await self._crawl_links(cli, links, key)
                if codes:
                    # slug wildcard -> khoi can category-index de resolve, tra thang
                    await self._crawl_codes(cli, codes, key, inputs={c: c for c in codes})

            elif mode == "keyword":
                idx = await self.build_index(cli)
                kw = target.strip().lower()
                hits = [no for no, info in idx.items() if kw in (info.get("name") or "").lower()]
                hits = hits[:limit or MAX_ITEMS]
                self.log(f"[SYS] '{target}': {len(hits)}/{len(idx)} SP khop ten")
                slugs = {c: idx[c]["slug"] for c in hits}
                await self._crawl_codes(cli, hits, key, slugs=slugs, inputs={c: target for c in hits},
                                       known_valid=True)

            elif mode == "link":
                links = [u.strip() for u in target.splitlines() if u.strip().startswith("http")]
                await self._crawl_links(cli, links, key)

            else:
                self.log(f"[ERR] mode khong hop le: {mode}")
        return self.results

    async def _crawl_links(self, cli: httpx.AsyncClient, links: list[str], key: str) -> None:
        self._load_ckpt(key)   # nap self.results (loc san "Failed") — link mode dedup theo URL, khong theo Item No
        seen_url = {r.get("URL", "") for r in self.results if r.get("Status") != "Failed"} | \
                   {r.get("Input", "") for r in self.results if r.get("Status") == "Not Found"}
        links = [u for u in dict.fromkeys(links) if u not in seen_url][:MAX_ITEMS]
        self.total += len(links)
        sem = asyncio.Semaphore(self.conc)

        async def one(url: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                data, ok = await self.fetch_data(cli, url)
                if self.delay:
                    await asyncio.sleep(self.delay)
                if data and data.get("productNo"):
                    rec = parse_detail(data, url)
                else:
                    rec = {"Input": url, "Item No": "", "Page": "", "Status": "Not Found" if ok else "Failed",
                          "_cats": [], "_spec": "", "Price": None, "Image": "", "URL": url}
                self.results.append(rec)
                self._save(key, [rec])
                self.done += 1
                tag = {"OK": "OK ", "Not Found": "MISS", "Failed": "FAIL"}[rec["Status"]]
                self.log(f"[{tag}] {self.done}/{self.total} | {url[-60:]}")

        try:
            await asyncio.gather(*(one(u) for u in links))
        except (KeyboardInterrupt, Exception) as e:
            self.log(f"[FATAL] {type(e).__name__}: {e}")


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
    root.title("INGCO Crawler — ingco.com/vn")
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
    for val, txt in [("code", "1. Theo MA"), ("keyword", "2. Theo TU KHOA (ten SP)"),
                     ("link", "3. Theo LINK (moi dong 1 link)"), ("excel", "4. Theo cot EXCEL"),
                     ("all", "5. Crawl ALL (~350 SP, 6 category VN)")]:
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

    conc, delay, limit = tk.IntVar(value=3), tk.DoubleVar(value=0.05), tk.IntVar(value=0)
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
        except BaseException as e:
            q.put(f"[FATAL] {type(e).__name__}: {e}")
        finally:
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
        p = ask_export_path(xls.get(), HERE, "ingco.xlsx")
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
    ap = argparse.ArgumentParser(description="INGCO crawler (ingco.com/vn)")
    ap.add_argument("--mode", choices=["code", "keyword", "link", "excel", "all"])
    ap.add_argument("--code")
    ap.add_argument("--keyword")
    ap.add_argument("--link")
    ap.add_argument("--excel")
    ap.add_argument("--sheet", default="")
    ap.add_argument("--col", default="")
    ap.add_argument("--out", default="ingco.xlsx")
    ap.add_argument("--conc", type=int, default=3)  # DA DO: cao hon -> ty le SSR-rong tang vot
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
