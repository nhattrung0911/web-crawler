"""
MPE crawler — all-in-one (mpe.com.vn + mpemall.com). 5 modes.

  data     : Excel (cot 'part')  -> mpe.com.vn: tim SP -> bang thong so -> xlsx
  anh-vn   : Excel (cot 'Part')  -> mpe.com.vn: anh chinh SP -> images/ + report
  anh-mall : Excel (cot 'Part')  -> mpemall.com: anh chinh SP -> output/
  anh-link : Excel (cot ID, Product Link) -> tai anh tu link truc tiep (Selenium)
  pdf      : PDF catalog MPE -> trich moi bang -> xlsx tho

Run:
  mpe.bat                            -> GUI
  python mpe.py                      -> GUI
  python mpe.py --mode data --input parts.xlsx [--output MPE_data.xlsx]
Read-only (GET).
"""
from __future__ import annotations

import argparse
import asyncio
import os
import queue
import re
import subprocess
import sys
import threading
from pathlib import Path

import pandas as pd

HERE = Path(__file__).parent
MPE = "https://www.mpe.com.vn"
MALL = "https://mpemall.com"


def _norm(s):
    return re.sub(r"[^A-Z0-9/-]", "", str(s).upper().strip())


# ============ mode: data (mpe.com.vn specs) ============

async def _data_detail(page, url):
    if not url:
        return {"error": "no_url"}
    if url.startswith("/"):
        url = MPE + url
    await page.goto(url, timeout=60000)
    r = {"url": url}
    for key, sel in (("Danh mục", "div.product-info div.product-category p.category a"),
                     ("Mã SP", "div.product-info div.product-code p.code")):
        try:
            r[key] = (await page.locator(sel).first.inner_text(timeout=3000)).strip()
        except Exception:
            r[key] = ""
    try:
        await page.wait_for_selector("div.full-spec table.spec-table tr", timeout=6000)
    except Exception:
        r["error"] = "no_spec"
        return r
    rows = page.locator("div.full-spec table.spec-table tr")
    for i in range(await rows.count()):
        try:
            label = (await rows.nth(i).locator("td").nth(0).inner_text()).strip().rstrip(":").strip()
            val = " ".join((await rows.nth(i).locator("td").nth(1).inner_text()).split())
            if label:
                r[" ".join(label.split())] = val
        except Exception:
            pass
    return r


async def _data_search(page, part):
    await page.goto(MPE + "/", timeout=60000)
    await page.click("div.search.lh-1 button.search-action")
    await page.fill("input[placeholder='Vui lòng nhập từ khoá...']", part)
    await page.press("input[placeholder='Vui lòng nhập từ khoá...']", "Enter")
    try:
        await page.wait_for_selector("ul.content-list li.search-result-item a", timeout=6000)
    except Exception:
        return None
    items = page.locator("ul.content-list li.search-result-item a")
    n = await items.count()
    if n == 0:
        return None
    for i in range(n):
        if part.lower() in (await items.nth(i).inner_text()).lower():
            return await items.nth(i).get_attribute("href")
    return await items.first.get_attribute("href")


async def run_data(inp, output="MPE_data.xlsx"):
    from playwright.async_api import async_playwright
    df = pd.read_excel(inp)
    if "part" not in df.columns:
        sys.exit("Excel can cot 'part'.")
    parts = df["part"].astype(str).tolist()
    results = []
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)

        async def worker(part, wid):
            ctx = await browser.new_context()
            page = await ctx.new_page()
            try:
                d = await _data_detail(page, await _data_search(page, part))
                d["part"] = part
                results.append(d)
                print(f"  ok {part}", flush=True)
            except Exception as e:
                results.append({"part": part, "error": str(e)})
                print(f"  err {part}: {e}", flush=True)
            finally:
                await ctx.close()

        for i in range(0, len(parts), 10):
            batch = parts[i:i + 10]
            print(f"[batch {i // 10 + 1}] {len(batch)} ma", flush=True)
            await asyncio.gather(*(worker(p, j) for j, p in enumerate(batch)))
        await browser.close()
    fixed = ["part", "url", "Danh mục", "Mã SP"]
    keys = set().union(*[r.keys() for r in results]) if results else set()
    cols = fixed + sorted(k for k in keys if k not in fixed)
    out = (HERE / output) if not os.path.isabs(output) else Path(output)
    pd.DataFrame(results).reindex(columns=cols).to_excel(out, index=False)
    print(f"DONE data -> {out}", flush=True)


# ============ mode: anh-vn (mpe.com.vn images) ============

def _exact(text, code):
    text = _norm(text)
    alt = code.replace("-", "/")
    return any(re.search(rf"(^|[^A-Z0-9]){re.escape(c)}([^A-Z0-9]|$)", text) for c in (code, alt))


async def run_anh_vn(inp, outdir="images"):
    import requests
    from playwright.async_api import async_playwright
    df = pd.read_excel(inp)
    col = "Part" if "Part" in df.columns else df.columns[0]
    root = (HERE / outdir) if not os.path.isabs(outdir) else Path(outdir)
    root.mkdir(parents=True, exist_ok=True)
    results = []
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False, slow_mo=100)
        page = await (await browser.new_context(viewport={"width": 1280, "height": 900})).new_page()
        for raw in df[col]:
            code = _norm(raw)
            try:
                await page.goto(MPE + "/", timeout=60000)
                await page.click(".search-action")
                await page.fill(".search-form-wrap input", code)
                await page.wait_for_selector(".search-result-item", timeout=10000)
                href = None
                for li in await page.query_selector_all(".search-result-item"):
                    a = await li.query_selector("a")
                    if not a:
                        continue
                    h = await a.get_attribute("href") or ""
                    sp = await li.query_selector("span.highligh")
                    if _exact((await sp.inner_text()) if sp else "", code) or _exact(h, code) \
                            or _exact(await a.inner_text(), code):
                        href = h
                        break
                if not href:
                    results.append({"Part": code, "ImageLink": "not found"})
                    print(f"  x {code}", flush=True)
                    continue
                if href.startswith("/"):
                    href = MPE + href
                await page.goto(href, timeout=60000)
                img = await page.query_selector("a[data-fancybox] img")
                src = await img.get_attribute("src") if img else None
                if not src:
                    results.append({"Part": code, "ImageLink": "no image"})
                    continue
                url = MPE + src if src.startswith("/") else src
                with open(root / f"{code}.png", "wb") as f:
                    f.write(requests.get(url, timeout=30).content)
                results.append({"Part": code, "ImageLink": url})
                print(f"  ok {code}", flush=True)
            except Exception as e:
                results.append({"Part": code, "ImageLink": "error"})
                print(f"  err {code}: {e}", flush=True)
        await browser.close()
    pd.DataFrame(results).to_excel(root / "report.xlsx", index=False)
    print(f"DONE anh-vn -> {root}", flush=True)


# ============ mode: anh-mall (mpemall.com via WooCommerce Store API) ============
# mpemall.com is WooCommerce -> public Store API returns name/price/sku/images as
# JSON in ONE request. No browser needed (fast). Gets product DATA + images together.
MALL_API = MALL + "/wp-json/wc/store/v1/products"


def _nz(t):
    return re.sub(r"[^a-z0-9]", "", (t or "").lower())


async def run_anh_mall(inp, outdir="output_mall"):
    import httpx
    df = pd.read_excel(inp)
    col = "Part" if "Part" in df.columns else df.columns[0]
    parts = [str(p).strip() for p in df[col].dropna()]
    root = (HERE / outdir) if not os.path.isabs(outdir) else Path(outdir)
    root.mkdir(parents=True, exist_ok=True)
    results = []
    sem = asyncio.Semaphore(8)

    async def one(client, part):
        async with sem:
            res = {"part": part, "status": "unknown"}
            try:
                r = await client.get(MALL_API, params={"search": part, "per_page": 10})
                r.raise_for_status()
                products = r.json()
                if not products:
                    res["status"] = "no_result"
                    results.append(res)
                    print(f"  x {part}", flush=True)
                    return
                # best match by sku/name, else first
                best = next((p for p in products if _nz(part) in _nz(p.get("sku"))
                             or _nz(part) in _nz(p.get("name"))), products[0])
                res.update({"product_id": best.get("id"), "name": best.get("name"),
                            "sku": best.get("sku"), "permalink": best.get("permalink"),
                            "price": (best.get("prices") or {}).get("price")})
                folder = root / part.replace("/", "-")
                folder.mkdir(parents=True, exist_ok=True)
                imgs = best.get("images") or []
                for i, im in enumerate(imgs, 1):
                    src = im.get("src")
                    if not src:
                        continue
                    ir = await client.get(src)
                    if ir.status_code == 200:
                        ext = os.path.splitext(src.split("?")[0])[1] or ".jpg"
                        with open(folder / f"{part.replace('/', '-')}_{i}{ext}", "wb") as f:
                            f.write(ir.content)
                res["images"] = len(imgs)
                res["status"] = "ok"
                print(f"  ok {part} ({len(imgs)} anh)", flush=True)
            except Exception as e:
                res.update({"status": "error", "error": str(e)})
                print(f"  err {part}: {e}", flush=True)
            results.append(res)

    async with httpx.AsyncClient(timeout=30, headers={"User-Agent": "Mozilla/5.0"},
                                 follow_redirects=True) as client:
        await asyncio.gather(*(one(client, p) for p in parts))
    pd.DataFrame(results).to_excel(root / "report.xlsx", index=False)
    print(f"DONE anh-mall (Store API) -> {root}", flush=True)


# ============ mode: anh-link (direct product links, Selenium) ============

def run_anh_link(inp, outdir="Product_Images"):
    import requests
    from urllib.parse import urljoin, urlparse
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    def sane(n):
        return re.sub(r'[\\/*?:"<>|]', "_", str(n)).strip()[:100]

    root = (HERE / outdir) if not os.path.isabs(outdir) else Path(outdir)
    root.mkdir(parents=True, exist_ok=True)
    opts = webdriver.ChromeOptions()
    opts.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(options=opts)
    session = requests.Session()
    df = pd.read_excel(inp)
    if "ID" not in df.columns or "Product Link" not in df.columns:
        driver.quit()
        sys.exit("Excel can cot ID, Product Link.")
    try:
        for _, row in df.iterrows():
            pid, url = str(row["ID"]), row["Product Link"]
            dest = root / pid
            dest.mkdir(parents=True, exist_ok=True)
            print(f"  {pid}", flush=True)
            try:
                driver.get(url)
                WebDriverWait(driver, 30).until(EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div.product-images__item img")))
                base = f"{urlparse(url).scheme}://{urlparse(url).netloc}"
                for i, img in enumerate(driver.find_elements(By.CSS_SELECTOR, "div.product-images__item img"), 1):
                    src = img.get_attribute("src") or img.get_attribute("data-src")
                    if not src:
                        continue
                    src = urljoin(base, src)
                    name = sane(img.get_attribute("alt") or os.path.basename(src))
                    r = session.get(src, timeout=25)
                    if r.status_code == 200:
                        with open(dest / f"{name}_{pid}_{i}.jpg", "wb") as f:
                            f.write(r.content)
            except Exception as e:
                print(f"  err {pid}: {e}", flush=True)
    finally:
        driver.quit()
    print(f"DONE anh-link -> {root}", flush=True)


# ============ mode: pdf (catalog -> raw tables) ============

def run_pdf(pdf_path, output=None):
    import pdfplumber
    out = output or (Path(pdf_path).with_suffix("").as_posix() + "_tables.xlsx")
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for pageno, page in enumerate(pdf.pages, 1):
            try:
                for t in page.extract_tables():
                    if t and len(t) >= 2:
                        d = pd.DataFrame(t[1:], columns=t[0])
                        d["__page__"] = pageno
                        tables.append(d)
            except Exception as e:
                print(f"  page {pageno} err: {e}", flush=True)
    if not tables:
        print("Khong trich duoc bang nao tu PDF.", flush=True)
        return
    pd.concat(tables, ignore_index=True).to_excel(out, index=False)
    print(f"DONE pdf -> {out}", flush=True)


# ---------------- GUI ----------------

def launch_gui():
    import customtkinter as ctk
    from tkinter import filedialog, messagebox

    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
    MODES = [("Data (mpe.vn specs)", "data"), ("Anh mpe.vn", "anh-vn"),
             ("Anh mpemall", "anh-mall"), ("Anh tu link", "anh-link"), ("PDF catalog", "pdf")]

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title("MPE Crawler")
            self.geometry("900x580")
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(2, weight=1)
            self.proc = None
            self.log_q: queue.Queue = queue.Queue()

            top = ctk.CTkFrame(self)
            top.grid(row=0, column=0, padx=16, pady=(16, 4), sticky="ew")
            self.mode = ctk.StringVar(value="data")
            for i, (t, v) in enumerate(MODES):
                ctk.CTkRadioButton(top, text=t, variable=self.mode, value=v).grid(
                    row=0, column=i, padx=8, pady=8)

            f = ctk.CTkFrame(self)
            f.grid(row=1, column=0, padx=16, pady=6, sticky="ew")
            f.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(f, text="Input (Excel / PDF)").grid(row=0, column=0, padx=8, pady=8, sticky="w")
            self.inp = ctk.CTkEntry(f, placeholder_text="chon file")
            self.inp.grid(row=0, column=1, padx=8, sticky="ew")
            ctk.CTkButton(f, text="...", width=40, command=self._pick).grid(row=0, column=2, padx=6)
            self.btn = ctk.CTkButton(f, text="START", width=110, command=self.start)
            self.btn.grid(row=0, column=3, padx=6)
            self.btn_stop = ctk.CTkButton(f, text="STOP", width=90, fg_color="#a33",
                                          hover_color="#c44", command=self.stop, state="disabled")
            self.btn_stop.grid(row=0, column=4, padx=6)

            self.console = ctk.CTkTextbox(self, state="disabled", font=("Consolas", 12))
            self.console.grid(row=2, column=0, padx=16, pady=12, sticky="nsew")
            self.after(100, self._drain)

        def _pick(self):
            ft = [("PDF", "*.pdf")] if self.mode.get() == "pdf" else [("Excel", "*.xlsx *.xls")]
            p = filedialog.askopenfilename(filetypes=ft)
            if p:
                self.inp.delete(0, "end")
                self.inp.insert(0, p)

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
            inp = self.inp.get().strip()
            if not inp:
                messagebox.showwarning("Thieu file", "Chon input.")
                return
            self.btn.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            args = ["--mode", self.mode.get(), "--input", inp]
            threading.Thread(target=self._run, args=(args,), daemon=True).start()

        def _run(self, args):
            env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"}
            self.log_q.put(f"\n===== mpe {' '.join(args)} =====\n")
            try:
                self.proc = subprocess.Popen(
                    [sys.executable, str(HERE / "mpe.py"), *args], cwd=str(HERE),
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
    ap = argparse.ArgumentParser(description="MPE crawler (GUI default)")
    ap.add_argument("--mode", choices=["data", "anh-vn", "anh-mall", "anh-link", "pdf"])
    ap.add_argument("--input")
    ap.add_argument("--output")
    a = ap.parse_args()
    if not a.mode or not a.input:
        return launch_gui()
    if a.mode == "data":
        asyncio.run(run_data(a.input, a.output or "MPE_data.xlsx"))
    elif a.mode == "anh-vn":
        asyncio.run(run_anh_vn(a.input))
    elif a.mode == "anh-mall":
        asyncio.run(run_anh_mall(a.input))
    elif a.mode == "anh-link":
        run_anh_link(a.input)
    elif a.mode == "pdf":
        run_pdf(a.input, a.output)


if __name__ == "__main__":
    main()
