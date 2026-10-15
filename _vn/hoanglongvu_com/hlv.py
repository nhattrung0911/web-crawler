"""
hoanglongvu.com — 2-stage crawler (Playwright).

  link : Excel (cot 'part') -> search -> products_links.csv (part, product_url)
  data : CSV/Excel (cot product_url, part) -> chi tiet (gia/specs/anh) -> xlsx + images/

Run:
  hlv.bat                       -> GUI
  python hlv.py                 -> GUI
  python hlv.py --mode link --excel parts.xlsx [--output products_links.csv] [--headless]
  python hlv.py --mode data --input links.csv [--output HLV_Report.xlsx] [--headless]
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
import time
from pathlib import Path

import aiohttp
import pandas as pd
from playwright.async_api import async_playwright

HERE = Path(__file__).parent
BASE = "https://hoanglongvu.com/"


# ================= stage: link =================

async def _block_static(route):
    if any(e in route.request.url for e in
           (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".css",
            ".woff", ".woff2", ".ttf", ".eot", "fonts.gstatic")):
        await route.abort()
    else:
        await route.continue_()


async def search_product(page, part_code: str):
    try:
        await page.goto(BASE, timeout=25000)
        await page.wait_for_timeout(200)
        await page.fill("input[name='s']", part_code)
        await page.press("input[name='s']", "Enter")
        await page.wait_for_timeout(400)
        if "/san-pham/" in page.url:
            return [{"part": part_code, "product_url": page.url, "category": ""}]
        section = page.locator("section.taxonomy-product-listing")
        await section.wait_for(timeout=3000)
        items = section.locator(".product-item")
        out = []
        for i in range(await items.count()):
            link = await items.nth(i).locator("a").first.get_attribute("href")
            if link:
                out.append({"part": part_code, "product_url": link, "category": ""})
        return out
    except Exception:
        return []


async def link_run(excel: str, output: str = "products_links.csv", headless: bool = True):
    df = pd.read_excel(excel)
    if "part" not in df.columns:
        sys.exit("Excel can cot 'part'.")
    parts = df["part"].astype(str).tolist()
    results, done = [], 0
    print(f"Tim link {len(parts)} ma...", flush=True)

    async def worker(browser, batch, wid):
        nonlocal done
        ctx = await browser.new_context()
        await ctx.route("**/*", _block_static)
        page = await ctx.new_page()
        for part in batch:
            links = await search_product(page, part)
            results.extend(links)
            done += 1
            print(f"[{done}/{len(parts)}] {part} -> {len(links)} link", flush=True)
        await ctx.close()

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=headless)
        w = 10
        batches = [parts[i:i + w] for i in range(0, len(parts), w)]
        await asyncio.gather(*(worker(browser, b, i + 1) for i, b in enumerate(batches)))
        await browser.close()
    out = (HERE / output) if not os.path.isabs(output) else Path(output)
    pd.DataFrame(results).to_csv(out, index=False, encoding="utf-8-sig")
    print(f"DONE link -> {out}", flush=True)


# ================= stage: data =================

def _sanitize(name):
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() if name else "unknown"


def _largest_srcset(srcset):
    if not srcset:
        return None
    items = []
    for chunk in srcset.split(","):
        part = chunk.strip().rsplit(" ", 1)
        if len(part) == 2 and part[1].endswith("w"):
            try:
                items.append((int(part[1][:-1]), part[0].strip()))
            except ValueError:
                pass
    return sorted(items, reverse=True)[0][1] if items else None


async def _dl_img(session, url, path):
    if not url or url.startswith("data:"):
        return
    try:
        async with session.get(url, timeout=25) as r:
            if r.status == 200:
                with open(path, "wb") as f:
                    f.write(await r.read())
    except Exception as e:
        print(f"  ! img {url}: {e}", flush=True)


async def _extract_price(page):
    try:
        t = await page.locator(".regular-price").text_content()
        if t and t.strip():
            return t.strip()
    except Exception:
        pass
    try:
        css = await page.evaluate(
            "()=>{let el=document.querySelector('.regular-price');if(!el)return '';"
            "return window.getComputedStyle(el,'::before').getPropertyValue('content').replace(/\"/g,'').trim();}")
        if css and css.strip():
            return css.strip()
    except Exception:
        pass
    try:
        block = " ".join((await page.locator(".price-wrap").inner_text()).split())
        found = re.findall(r"[\d\.]+[₫đ]", block)
        if found:
            return found[0]
    except Exception:
        pass
    return ""


async def _crawl_detail(page, session, row, outdir):
    url, part = row["product_url"], str(row["part"])
    result = {"part": part, "product_url": url}
    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    await page.wait_for_timeout(80)
    try:
        cats = await page.locator(".rank-math-breadcrumb p a").all_inner_texts()
        result["category"] = cats[-1] if cats else ""
    except Exception:
        result["category"] = ""
    for key, sel in (("title", "h1.product_title"), ("brand", ".brand a")):
        try:
            result[key] = (await page.inner_text(sel)).strip()
        except Exception:
            result[key] = ""
    result["price"] = await _extract_price(page)
    try:
        btn = page.locator("//a[span[contains(text(),'Xem đầy đủ')]]")
        if await btn.count() > 0:
            await btn.first.click()
            await page.wait_for_timeout(120)
    except Exception:
        pass
    try:
        trs = page.locator("table.w-100 tbody tr")
        for i in range(await trs.count()):
            tds = trs.nth(i).locator("td")
            if await tds.count() >= 2:
                label = (await tds.nth(0).inner_text()).replace(":", "").strip()
                value = (await tds.nth(1).inner_text()).strip()
                if label and value:
                    result[_sanitize(label)] = value
    except Exception:
        pass
    urls = set()
    main_img = page.locator(".woocommerce-product-gallery img.wp-post-image").first
    if await main_img.count() == 1:
        u = await main_img.get_attribute("data-large_image")
        if not u:
            u = _largest_srcset(await main_img.get_attribute("srcset"))
        if not u:
            u = await main_img.get_attribute("src")
        if u and u.startswith("http"):
            urls.add(u)
    folder = outdir / _sanitize(result.get("category") or "Unknown") / _sanitize(part)
    folder.mkdir(parents=True, exist_ok=True)
    tasks = []
    for idx, u in enumerate(urls, 1):
        local = folder / f"{_sanitize(part)}_img{idx}.jpg"
        result[f"image_{idx}"] = u
        result[f"image_local_{idx}"] = str(local)
        tasks.append(_dl_img(session, u, local))
    await asyncio.gather(*tasks)
    result["status"] = "ok"
    return result


async def data_run(input_path: str, output: str = "HLV_Report.xlsx", headless: bool = False):
    p = Path(input_path)
    df = pd.read_csv(p) if p.suffix.lower() == ".csv" else pd.read_excel(p)
    if "product_url" not in df.columns:
        sys.exit("Input can cot 'product_url' (va 'part').")
    if "part" not in df.columns:
        df["part"] = ""
    df = df.drop_duplicates(subset=["product_url"]).reset_index(drop=True)
    rows = df.to_dict("records")
    outdir = HERE / "images"
    results, done, start = [], 0, time.time()
    print(f"Crawl {len(rows)} san pham...", flush=True)

    async def worker(browser, batch, wid, session):
        nonlocal done
        ctx = await browser.new_context()

        async def block(route):
            if route.request.resource_type in ("stylesheet", "font", "media"):
                await route.abort()
            else:
                await route.continue_()

        await ctx.route("**/*", block)
        page = await ctx.new_page()
        page.set_default_timeout(15000)
        for row in batch:
            ok = False
            for attempt in range(3):
                try:
                    await page.goto(row["product_url"], wait_until="domcontentloaded")
                    await page.wait_for_timeout(80)
                    results.append(await _crawl_detail(page, session, row, outdir))
                    ok = True
                    break
                except Exception as e:
                    print(f"  W{wid} err {row['product_url']} ({attempt + 1}): {e}", flush=True)
                    page = await ctx.new_page()
            if not ok:
                results.append({"part": row.get("part", ""), "product_url": row["product_url"], "status": "failed"})
            done += 1
            if done % 5 == 0 or done == len(rows):
                sp = done / (time.time() - start) if time.time() > start else 0
                print(f"[{done}/{len(rows)}] {sp:.2f} sp/s", flush=True)
        await ctx.close()

    import multiprocessing
    w = min(20, multiprocessing.cpu_count() * 2)
    batches = [rows[i:i + w] for i in range(0, len(rows), w)]
    async with async_playwright() as pw, aiohttp.ClientSession() as session:
        browser = await pw.chromium.launch(headless=headless)
        await asyncio.gather(*(worker(browser, b, i + 1, session) for i, b in enumerate(batches)))
        await browser.close()
    out = pd.DataFrame(results)
    base = ["part", "product_url", "status", "category", "title", "price", "brand"]
    specs = sorted(c for c in out.columns if c not in base and not c.startswith("image_"))
    imgs = sorted(c for c in out.columns if c.startswith("image_"))
    out = out.reindex(columns=[c for c in base if c in out.columns] + specs + imgs)
    dest = (HERE / output) if not os.path.isabs(output) else Path(output)
    out.to_excel(dest, index=False)
    print(f"DONE data -> {dest}", flush=True)


# ================= GUI =================

def launch_gui():
    import customtkinter as ctk
    from tkinter import filedialog, messagebox

    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title("Hoang Long Vu Crawler")
            self.geometry("880x580")
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(3, weight=1)
            self.proc = None
            self.log_q: queue.Queue = queue.Queue()

            top = ctk.CTkFrame(self)
            top.grid(row=0, column=0, padx=16, pady=(16, 4), sticky="ew")
            self.mode = ctk.StringVar(value="link")
            ctk.CTkRadioButton(top, text="Link (part -> URL)", variable=self.mode,
                               value="link", command=self._toggle).grid(row=0, column=0, padx=10, pady=8)
            ctk.CTkRadioButton(top, text="Data (URL -> chi tiet + anh)", variable=self.mode,
                               value="data", command=self._toggle).grid(row=0, column=1, padx=10)
            self.headless = ctk.CTkCheckBox(top, text="Headless")
            self.headless.grid(row=0, column=2, padx=16)

            f = ctk.CTkFrame(self)
            f.grid(row=1, column=0, padx=16, pady=6, sticky="ew")
            f.grid_columnconfigure(1, weight=1)
            self.lbl = ctk.CTkLabel(f, text="Excel input")
            self.lbl.grid(row=0, column=0, padx=8, pady=6, sticky="w")
            self.inp = ctk.CTkEntry(f, placeholder_text="chon file")
            self.inp.grid(row=0, column=1, padx=8, sticky="ew")
            ctk.CTkButton(f, text="...", width=40, command=self._pick).grid(row=0, column=2, padx=6)

            bar = ctk.CTkFrame(self, fg_color="transparent")
            bar.grid(row=2, column=0, padx=18, sticky="ew")
            self.btn = ctk.CTkButton(bar, text="START", width=120, command=self.start)
            self.btn.grid(row=0, column=0, padx=6)
            self.btn_stop = ctk.CTkButton(bar, text="STOP", width=90, fg_color="#a33",
                                          hover_color="#c44", command=self.stop, state="disabled")
            self.btn_stop.grid(row=0, column=1, padx=6)

            self.console = ctk.CTkTextbox(self, state="disabled", font=("Consolas", 12))
            self.console.grid(row=3, column=0, padx=16, pady=12, sticky="nsew")
            self._toggle()
            self.after(100, self._drain)

        def _toggle(self):
            self.lbl.configure(text="Excel (cot part)" if self.mode.get() == "link"
                               else "CSV/Excel (cot product_url)")

        def _pick(self):
            p = filedialog.askopenfilename(filetypes=[("Data", "*.xlsx *.xls *.csv")])
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
            m, inp = self.mode.get(), self.inp.get().strip()
            if not inp:
                messagebox.showwarning("Thieu file", "Chon input.")
                return
            args = ["--mode", m, ("--excel" if m == "link" else "--input"), inp]
            if self.headless.get():
                args.append("--headless")
            self.btn.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            threading.Thread(target=self._run, args=(args,), daemon=True).start()

        def _run(self, args):
            env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"}
            self.log_q.put(f"\n===== hlv {' '.join(args)} =====\n")
            try:
                self.proc = subprocess.Popen(
                    [sys.executable, str(HERE / "hlv.py"), *args], cwd=str(HERE),
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
    ap = argparse.ArgumentParser(description="hoanglongvu.com crawler (GUI default)")
    ap.add_argument("--mode", choices=["link", "data"])
    ap.add_argument("--excel")
    ap.add_argument("--input")
    ap.add_argument("--output")
    ap.add_argument("--headless", action="store_true")
    a = ap.parse_args()
    if a.mode == "link" and a.excel:
        asyncio.run(link_run(a.excel, a.output or "products_links.csv", a.headless or True))
    elif a.mode == "data" and a.input:
        asyncio.run(data_run(a.input, a.output or "HLV_Report.xlsx", a.headless))
    else:
        launch_gui()


if __name__ == "__main__":
    main()
