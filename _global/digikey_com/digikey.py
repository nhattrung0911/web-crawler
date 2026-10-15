"""
DigiKey crawler (digikey.com) — WEB scrape via Playwright, NO login / NO registration.

Cloudflare-gated web; a real Chromium (Playwright) passes it. Every page is Next.js:
full data sits in embedded __NEXT_DATA__ (props.pageProps.envelope.data) -> read
structured JSON, no HTML parsing.

  result : /en/products/result?keywords=X               -> categories (topResults)
  listing: /en/products/filter/{slug}/{id}?...&page=N    -> data.products (rows)
  detail : /en/products/detail/{mfr}/{part}/{id}         -> data.productOverview + attrs
  index  : /en/products                                  -> data.categories (tree)

5 MODES:
  keyword : search keyword           -> san pham
  code    : 1 ma linh kien           -> san pham
  excel   : Excel (cot Part/Code)    -> tung ma -> san pham
  link    : dan link detail (1..n)   -> chi tiet day du
  all     : crawl toan bo catalog    -> san pham (gioi han --limit)

Run:
  digikey.bat                        -> GUI
  python digikey.py --mode keyword --input LM358 [--limit 300] [--headed]
  python digikey.py --mode code --input LM358P
  python digikey.py --mode excel --excel codes.xlsx
  python digikey.py --mode link --input "https://www.digikey.com/en/products/detail/..."
  python digikey.py --mode all --limit 5000
Read-only (chi render trang cong khai).
"""
from __future__ import annotations

import argparse
import asyncio
import os
import queue
import subprocess
import sys
import threading
from pathlib import Path

import pandas as pd

HERE = Path(__file__).parent
BASE = "https://www.digikey.com"
PAGE_SIZE = 100

_NEXT = """() => {
  const el = document.getElementById('__NEXT_DATA__');
  if (!el) return null;
  try { return JSON.parse(el.textContent).props.pageProps.envelope; }
  catch (e) { return null; }
}"""


async def _new_page(pw, headed):
    browser = await pw.chromium.launch(
        headless=not headed, args=["--disable-blink-features=AutomationControlled"])
    ctx = await browser.new_context(
        viewport={"width": 1400, "height": 900}, locale="en-US",
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    page = await ctx.new_page()
    # warm up: clear Cloudflare on the homepage so its clearance cookie is set for
    # the whole context (deep-link detail pages get challenged harder otherwise).
    try:
        await page.goto(BASE, wait_until="domcontentloaded", timeout=45000)
        await page.wait_for_timeout(2500)
    except Exception:
        pass
    return browser, page


async def _envelope(page, url):
    await page.goto(url, wait_until="domcontentloaded", timeout=45000)
    for _ in range(6):
        env = await page.evaluate(_NEXT)
        if env:
            return env
        try:
            await page.wait_for_function("document.getElementById('__NEXT_DATA__')!==null", timeout=3000)
        except Exception:
            await page.wait_for_timeout(1500)  # Cloudflare challenge settling
    return None


# ---------------- parsing ----------------

_CORE = ("manufacturerPartNumber", "productNumber", "packaging", "price",
         "shortDescription", "description", "datasheetUrl", "quantityAvailable",
         "productId", "productUrl", "minimumOrderQuantity")


def parse_rows(products, query=None):
    rows = []
    for cells in products or []:
        r = {}
        for cell in cells if isinstance(cells, list) else []:
            v = cell.get("value") if isinstance(cell, dict) else None
            if not isinstance(v, dict):
                continue
            for k in _CORE:
                if k in v and not r.get(k):
                    r[k] = v[k]
            m = v.get("manufacturer")
            if isinstance(m, dict) and not r.get("manufacturer"):
                r["manufacturer"] = m.get("name")
        if r.get("manufacturerPartNumber") or r.get("productNumber"):
            row = {
                "Part": r.get("manufacturerPartNumber"),
                "DigiKeyPart": r.get("productNumber"),
                "Manufacturer": r.get("manufacturer"),
                "Description": r.get("shortDescription") or r.get("description"),
                "Packaging": r.get("packaging"),
                "UnitPrice": r.get("price"),
                "Stock": r.get("quantityAvailable"),
                "MinQty": r.get("minimumOrderQuantity"),
                "Datasheet": r.get("datasheetUrl"),
                "URL": (BASE + r["productUrl"]) if str(r.get("productUrl", "")).startswith("/") else r.get("productUrl"),
                "ProductId": r.get("productId"),
            }
            if query:
                row["Query"] = query
            rows.append(row)
    return rows


def _clean(x):
    """Unwrap DigiKey wrapped values (stringList / {label,value} / lists) -> plain str."""
    if x is None or isinstance(x, (str, int, float, bool)):
        return x
    if isinstance(x, dict):
        if x.get("type") == "stringList" and isinstance(x.get("value"), list):
            return ", ".join(str(i.get("label") or i.get("value")) for i in x["value"] if isinstance(i, dict))
        if "label" in x:
            return str(x["label"])
        if "value" in x:
            return _clean(x["value"])
        return None
    if isinstance(x, list):
        return ", ".join(s for s in (str(_clean(i)) for i in x) if s and s != "None")
    return str(x)


def _attrs(pa):
    """productAttributes = {attributes:[{label, values:[{value}]}]} -> {label: value}."""
    out = {}
    node = pa.get("attributes") if isinstance(pa, dict) else pa
    for a in node or []:
        if not isinstance(a, dict):
            continue
        label = a.get("label") or a.get("name")
        vals = a.get("values")
        if isinstance(vals, list):
            v = ", ".join(str(x.get("value")) for x in vals if isinstance(x, dict) and x.get("value"))
        else:
            v = _clean(vals if vals is not None else a.get("value"))
        if label and v:
            out[str(label)] = v
    return out


def _price(pq):
    """priceQuantity -> (stock, unit_price@1, breaks, packaging, min_qty) as plain strings."""
    if not isinstance(pq, dict):
        return None, None, None, None, None
    pr = (pq.get("pricing") or [{}])[0]
    tiers = pr.get("mergedPricingTiers") or []
    unit = tiers[0].get("unitPrice") if tiers else None
    breaks = "; ".join(f"{t.get('brkQty')}:{t.get('unitPrice')}" for t in tiers[:8])
    return pq.get("qtyAvailable"), unit, breaks, pr.get("packaging"), pr.get("minOrderQuantity")


def parse_detail(data):
    ov = data.get("productOverview") or {}
    mfr = ov.get("manufacturer")
    mfr = mfr.get("name") if isinstance(mfr, dict) else mfr
    stock, unit, breaks, packaging, minq = _price(data.get("priceQuantity"))
    bc = " > ".join(b.get("text", "") for b in (data.get("breadcrumb") or []) if b.get("text"))
    row = {
        "Part": _clean(ov.get("manufacturerProductNumber")),
        "DigiKeyPart": _clean(ov.get("digikeyProductNumbers")),
        "Manufacturer": mfr,
        "Description": _clean(ov.get("description")),
        "DetailedDescription": _clean(ov.get("detailedDescription")),
        "Category": bc,
        "UnitPrice": unit,
        "PriceBreaks": breaks,
        "Stock": stock,
        "MinQty": minq,
        "Packaging": packaging,
        "LeadTime": _clean(ov.get("standardLeadTime")),
        "Datasheet": _clean(ov.get("datasheetUrl")),
    }
    row.update(_attrs(data.get("productAttributes")))
    return row


# ---------------- shared crawl helpers ----------------

async def _listing_pages(page, curl, keywords, limit, rows, label=""):
    pg = 1
    while len(rows) < limit:
        sep = "&" if "?" in curl else "?"
        kw = f"keywords={keywords}&" if keywords else ""
        url = f"{BASE}{curl}{sep}{kw}page={pg}&pageSize={PAGE_SIZE}"
        env = await _envelope(page, url)
        d = (env or {}).get("data") or {}
        got = parse_rows(d.get("products"), keywords or None)
        if not got:
            break
        rows.extend(got)
        total = int(str(d.get("productCount", "0")).replace(",", "") or 0)
        print(f"  {label} p{pg}: +{len(got)} ({len(rows)})", flush=True)
        pg += 1
        if pg * PAGE_SIZE > total or total == 0:
            break


async def _resolve(page, term):
    """Resolve a keyword/code search. Returns (kind, payload):
    ('detail', data) if it redirected straight to a product page (exact part),
    ('listing', [url]) if it landed on a product listing,
    ('cats', [urls]) for the category buckets, or (None, None) if blocked."""
    env = await _envelope(page, f"{BASE}/en/products/result?keywords={term}")
    if not env:
        return None, None
    d = env.get("data") or {}
    if d.get("productOverview"):
        return "detail", d
    if d.get("exactMatch"):                    # exact part -> category page w/ exact matches
        return "exact", d["exactMatch"]
    if d.get("products"):
        return "listing", [page.url.replace(BASE, "").split("?")[0]]
    return "cats", [c["categoryUrl"] for c in (d.get("topResults") or []) if c.get("categoryUrl")]


async def _exact_details(page, matches, limit):
    """exactMatch list -> load each detail page for full data (fallback to summary)."""
    out = []
    for e in matches[:limit]:
        url = e.get("detailUrl", "")
        d = (await _envelope(page, BASE + url) or {}).get("data") if url else None
        if d and d.get("productOverview"):
            out.append(parse_detail(d))
        else:
            out.append({"Part": e.get("mfrProduct"), "Manufacturer": e.get("mfr"),
                        "Description": e.get("description"), "UnitPrice": e.get("unitPrice"),
                        "URL": BASE + url})
        print(f"    exact {e.get('mfrProduct')}", flush=True)
    return out


def _dedup(rows, limit):
    seen, out = set(), []
    for r in rows:
        k = r.get("DigiKeyPart") or r.get("Part")
        if k and k not in seen:
            seen.add(k)
            out.append(r)
        if len(out) >= limit:
            break
    return out


# ---------------- modes ----------------

async def run_search(term, limit, headed, output, label="search"):
    from playwright.async_api import async_playwright
    rows = []
    async with async_playwright() as pw:
        browser, page = await _new_page(pw, headed)
        try:
            kind, payload = await _resolve(page, term)
            if kind is None:
                sys.exit("Blocked (Cloudflare). Thu --headed.")
            if kind == "detail":                       # exact part -> straight to detail page
                _save([parse_detail(payload)], output)
                return
            if kind == "exact":                        # exact part -> exactMatch -> full detail
                _save(await _exact_details(page, payload, limit), output)
                return
            print(f"'{term}': {len(payload)} category, target <= {limit}", flush=True)
            for cu in payload:
                if len(rows) >= limit:
                    break
                await _listing_pages(page, cu, term, limit, rows, cu.split("/")[-1])
        finally:
            await browser.close()
    _save(_dedup(rows, limit), output)


async def run_excel(excel, per_code, headed, output):
    from playwright.async_api import async_playwright
    df = pd.read_excel(excel)
    col = next((c for c in df.columns if str(c).lower() in ("part", "code", "ma", "mpn")), df.columns[0])
    codes = [str(x).strip() for x in df[col].dropna()]
    rows = []
    async with async_playwright() as pw:
        browser, page = await _new_page(pw, headed)
        try:
            for i, code in enumerate(codes, 1):
                kind, payload = await _resolve(page, code)
                if kind == "detail":
                    rows.append(parse_detail(payload))
                    print(f"  [{i}/{len(codes)}] {code}: detail", flush=True)
                elif kind == "exact":
                    got = await _exact_details(page, payload, per_code)
                    rows.extend(got)
                    print(f"  [{i}/{len(codes)}] {code}: +{len(got)} exact", flush=True)
                elif kind in ("listing", "cats") and payload:
                    sub = []
                    await _listing_pages(page, payload[0], code, per_code, sub, code)
                    rows.extend(sub)
                    print(f"  [{i}/{len(codes)}] {code}: +{len(sub)}", flush=True)
                else:
                    print(f"  [{i}/{len(codes)}] x {code}", flush=True)
        finally:
            await browser.close()
    _save(rows, output)


async def run_link(links, headed, output):
    from playwright.async_api import async_playwright
    urls = [u.strip() for u in links.replace(",", "\n").splitlines() if "digikey.com" in u]
    if not urls:
        sys.exit("Khong co link digikey hop le.")
    rows = []
    async with async_playwright() as pw:
        browser, page = await _new_page(pw, headed)
        try:
            for i, url in enumerate(urls, 1):
                d = (await _envelope(page, url) or {}).get("data")
                if d and d.get("productOverview"):
                    rows.append(parse_detail(d))
                    print(f"  [{i}/{len(urls)}] ok {rows[-1].get('Part')}", flush=True)
                else:
                    rows.append({"URL": url, "Part": "NO_DATA/BLOCKED"})
                    print(f"  [{i}/{len(urls)}] x {url}", flush=True)
        finally:
            await browser.close()
    _save(rows, output)


async def run_all(limit, headed, output):
    from playwright.async_api import async_playwright
    rows = []
    async with async_playwright() as pw:
        browser, page = await _new_page(pw, headed)
        try:
            env = await _envelope(page, f"{BASE}/en/products")
            if not env:
                sys.exit("Blocked (Cloudflare). Thu --headed.")
            cats = (env.get("data") or {}).get("categories") or []
            subs = [(s.get("label"), s.get("url")) for c in cats for s in (c.get("subCategories") or []) if s.get("url")]
            print(f"CATALOG: {len(subs)} sub-category, target <= {limit}", flush=True)
            for label, cu in subs:
                if len(rows) >= limit:
                    break
                await _listing_pages(page, cu, "", limit, rows, (label or "")[:24])
        finally:
            await browser.close()
    _save(_dedup(rows, limit), output)


def _save(rows, output):
    if not rows:
        print("Khong co du lieu.", flush=True)
        return
    out = (HERE / output) if not os.path.isabs(output) else Path(output)
    pd.DataFrame(rows).to_excel(out, index=False)
    print(f"DONE -> {out} ({len(rows)} rows)", flush=True)


# ---------------- GUI ----------------

def launch_gui():
    import customtkinter as ctk
    from tkinter import filedialog, messagebox

    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
    MODES = [("Keyword", "keyword"), ("Ma (code)", "code"), ("Excel (nhieu ma)", "excel"),
             ("Link detail", "link"), ("Crawl ALL", "all")]

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title("DigiKey Crawler (web / no login)")
            self.geometry("920x600")
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(3, weight=1)
            self.proc = None
            self.log_q: queue.Queue = queue.Queue()

            top = ctk.CTkFrame(self)
            top.grid(row=0, column=0, padx=16, pady=(16, 4), sticky="ew")
            self.mode = ctk.StringVar(value="keyword")
            for i, (t, v) in enumerate(MODES):
                ctk.CTkRadioButton(top, text=t, variable=self.mode, value=v,
                                   command=self._toggle).grid(row=0, column=i, padx=8, pady=8)

            f = ctk.CTkFrame(self)
            f.grid(row=1, column=0, padx=16, pady=6, sticky="ew")
            f.grid_columnconfigure(0, weight=1)
            self.hint = ctk.CTkLabel(f, text="", anchor="w", text_color="#7fb0ff")
            self.hint.grid(row=0, column=0, columnspan=2, padx=8, pady=(6, 0), sticky="w")
            self.val = ctk.CTkTextbox(f, height=120, font=("Consolas", 12))
            self.val.grid(row=1, column=0, padx=8, pady=6, sticky="ew")
            self.pick = ctk.CTkButton(f, text="Chon Excel", width=90, command=self._pick)
            self.pick.grid(row=1, column=1, padx=6, sticky="n")

            opt = ctk.CTkFrame(self, fg_color="transparent")
            opt.grid(row=2, column=0, padx=18, sticky="ew")
            ctk.CTkLabel(opt, text="Limit").grid(row=0, column=0, padx=(0, 4))
            self.limit = ctk.CTkEntry(opt, width=80)
            self.limit.insert(0, "300")
            self.limit.grid(row=0, column=1, padx=4)
            self.headed = ctk.CTkCheckBox(opt, text="Headed (neu bi chan)")
            self.headed.grid(row=0, column=2, padx=16)
            self.btn = ctk.CTkButton(opt, text="START", width=120, command=self.start)
            self.btn.grid(row=0, column=3, padx=6)
            self.btn_stop = ctk.CTkButton(opt, text="STOP", width=90, fg_color="#a33",
                                          hover_color="#c44", command=self.stop, state="disabled")
            self.btn_stop.grid(row=0, column=4, padx=6)

            self.console = ctk.CTkTextbox(self, state="disabled", font=("Consolas", 12))
            self.console.grid(row=3, column=0, padx=16, pady=12, sticky="nsew")
            self._toggle()
            self.after(100, self._drain)

        def _toggle(self):
            m = self.mode.get()
            self.pick.configure(state="normal" if m == "excel" else "disabled")
            hints = {
                "keyword": "Nhap keyword (vd: LM358, op amp, STM32...)",
                "code": "Nhap 1 ma linh kien (vd: LM358P) -> lay chi tiet",
                "excel": "Bam 'Chon Excel' -> file co cot Part / Code (nhieu ma)",
                "link": "DAN NHIEU LINK SAN PHAM — moi dong 1 link:\n"
                        "  https://www.digikey.com/en/products/detail/.../.../123\n"
                        "  https://www.digikey.com/en/products/detail/.../.../456",
                "all": "Crawl TOAN BO catalog DigiKey — chi can dat Limit (khong nhap gi)",
            }
            self.hint.configure(text=hints[m])
            self.val.configure(state="normal")
            self.val.delete("1.0", "end")
            if m == "all":
                self.val.configure(state="disabled")

        def _pick(self):
            p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
            if p:
                self.val.delete("1.0", "end")
                self.val.insert("1.0", p)

        def _log(self, m):
            self.console.configure(state="normal")
            self.console.insert("end", m)
            self.console.see("end")
            self.console.configure(state="disabled")

        def _drain(self):
            try:
                while True:
                    self._log(self.log_q.get_nowait())
            except queue.Empty:
                pass
            self.after(100, self._drain)

        def start(self):
            if self.proc:
                return
            m = self.mode.get()
            v = self.val.get("1.0", "end").strip()
            if m != "all" and not v:
                messagebox.showwarning("Thieu", "Nhap keyword / ma / link / chon Excel.")
                return
            args = ["--mode", m, "--limit", self.limit.get().strip() or "300"]
            if m == "excel":
                args += ["--excel", v]
            elif m != "all":
                args += ["--input", v]
            if self.headed.get():
                args.append("--headed")
            self.btn.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            threading.Thread(target=self._run, args=(args,), daemon=True).start()

        def _run(self, args):
            env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"}
            self.log_q.put(f"\n===== digikey {args[1]} =====\n")
            try:
                self.proc = subprocess.Popen(
                    [sys.executable, str(HERE / "digikey.py"), *args], cwd=str(HERE),
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace", env=env)
            except OSError as e:
                self.log_q.put(f"[launch error] {e}\n")
            else:
                for line in self.proc.stdout:
                    self.log_q.put(line)
                self.proc.wait()
                self.log_q.put(f"[exit {self.proc.returncode}]\n")
            self.proc = None
            self.after(0, lambda: (self.btn.configure(state="normal"),
                                   self.btn_stop.configure(state="disabled")))

        def stop(self):
            if self.proc:
                self.proc.terminate()
                self.log_q.put("\n[da dung]\n")

    App().mainloop()


def main():
    ap = argparse.ArgumentParser(description="DigiKey web crawler (GUI default)")
    ap.add_argument("--mode", choices=["keyword", "code", "excel", "link", "all"])
    ap.add_argument("--input")
    ap.add_argument("--excel")
    ap.add_argument("--limit", type=int, default=300)
    ap.add_argument("--headed", action="store_true")
    ap.add_argument("--output")
    a = ap.parse_args()
    o = a.output
    if a.mode == "keyword" and a.input:
        asyncio.run(run_search(a.input, a.limit, a.headed, o or "digikey_keyword.xlsx", "keyword"))
    elif a.mode == "code" and a.input:
        asyncio.run(run_search(a.input, a.limit, a.headed, o or "digikey_code.xlsx", "code"))
    elif a.mode == "excel" and a.excel:
        asyncio.run(run_excel(a.excel, a.limit, a.headed, o or "digikey_excel.xlsx"))
    elif a.mode == "link" and a.input:
        asyncio.run(run_link(a.input, a.headed, o or "digikey_link.xlsx"))
    elif a.mode == "all":
        asyncio.run(run_all(a.limit, a.headed, o or "digikey_all.xlsx"))
    else:
        launch_gui()


if __name__ == "__main__":
    main()
