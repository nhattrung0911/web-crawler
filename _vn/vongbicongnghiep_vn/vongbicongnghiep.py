"""
vongbicongnghiep.vn — resolve product URLs by keyword (Playwright).

Run:
  vongbicongnghiep.bat                 -> GUI
  python vongbicongnghiep.py           -> GUI
  python vongbicongnghiep.py --excel in.xlsx --output out.xlsx [--headless] [--concurrency 3]

Reads an Excel with a 'URL' column of keywords, searches the site, writes
'Product URL' + 'Search Status' columns to the output Excel. Read-only (GET).
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
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

HERE = Path(__file__).parent
BASE_URL = "https://vongbicongnghiep.vn/"


async def resolve_product_url(browser, keyword: str):
    page = await browser.new_page()
    try:
        await page.goto(BASE_URL, timeout=30_000)
        await page.wait_for_selector("#search", timeout=10_000)
        await page.fill("#search", keyword)
        await page.press("#search", "Enter")
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(1.2)
        if await page.locator("section.related.products").count() > 0:
            return page.url, "Found - Direct"
        products = page.locator(
            'ul[data-bs-toggle="shop-products"] li.product a.woocommerce-LoopProduct-link')
        if await products.count() > 0:
            return await products.first.get_attribute("href"), "Found - From Listing"
        return "", "Not Found"
    except PWTimeout:
        return "", "Timeout"
    except Exception as e:
        return "", f"Error: {e}"
    finally:
        await page.close()


async def crawl(excel_path: str, output_path: str, headless: bool = True, concurrency: int = 3):
    df = pd.read_excel(excel_path)
    df.columns = df.columns.str.strip()
    if "URL" not in df.columns:
        sys.exit("Excel phai co cot 'URL' (chua keyword/ma can tim).")
    df["Product URL"] = ""
    df["Search Status"] = ""
    sem = asyncio.Semaphore(concurrency)
    done = 0

    async def worker(browser, idx):
        nonlocal done
        async with sem:
            kw = str(df.at[idx, "URL"])
            url, status = await resolve_product_url(browser, kw)
            df.at[idx, "Product URL"] = url
            df.at[idx, "Search Status"] = status
            done += 1
            print(f"[{done}/{len(df)}] {kw} -> {status}", flush=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        await asyncio.gather(*(worker(browser, i) for i in df.index))
        await browser.close()
    df.to_excel(output_path, index=False)
    print(f"DONE -> {output_path}", flush=True)


# ---------------- GUI ----------------

def launch_gui():
    import customtkinter as ctk
    from tkinter import filedialog, messagebox

    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title("Vong Bi Cong Nghiep - Resolve URL")
            self.geometry("860x560")
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(3, weight=1)
            self.proc = None
            self.log_q: queue.Queue = queue.Queue()

            f = ctk.CTkFrame(self)
            f.grid(row=0, column=0, padx=16, pady=(16, 6), sticky="ew")
            f.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(f, text="Excel input (cot URL)").grid(row=0, column=0, padx=8, pady=6, sticky="w")
            self.excel = ctk.CTkEntry(f, placeholder_text="chon file .xlsx")
            self.excel.grid(row=0, column=1, padx=8, pady=6, sticky="ew")
            ctk.CTkButton(f, text="...", width=40, command=self._pick_in).grid(row=0, column=2, padx=6)
            ctk.CTkLabel(f, text="Excel output").grid(row=1, column=0, padx=8, pady=6, sticky="w")
            self.out = ctk.CTkEntry(f, placeholder_text="ket qua .xlsx")
            self.out.grid(row=1, column=1, padx=8, pady=6, sticky="ew")
            ctk.CTkButton(f, text="...", width=40, command=self._pick_out).grid(row=1, column=2, padx=6)

            o = ctk.CTkFrame(self, fg_color="transparent")
            o.grid(row=1, column=0, padx=18, sticky="ew")
            self.headless = ctk.CTkCheckBox(o, text="Headless")
            self.headless.grid(row=0, column=0, padx=6)
            ctk.CTkLabel(o, text="Concurrency").grid(row=0, column=1, padx=(16, 4))
            self.conc = ctk.CTkEntry(o, width=60)
            self.conc.insert(0, "3")
            self.conc.grid(row=0, column=2)
            self.btn = ctk.CTkButton(o, text="START", width=120, command=self.start)
            self.btn.grid(row=0, column=3, padx=(16, 6))
            self.btn_stop = ctk.CTkButton(o, text="STOP", width=90, fg_color="#a33",
                                          hover_color="#c44", command=self.stop, state="disabled")
            self.btn_stop.grid(row=0, column=4)

            self.status = ctk.CTkLabel(self, text="San sang.", anchor="w")
            self.status.grid(row=2, column=0, padx=18, pady=(6, 0), sticky="ew")
            self.console = ctk.CTkTextbox(self, state="disabled", font=("Consolas", 12))
            self.console.grid(row=3, column=0, padx=16, pady=12, sticky="nsew")
            self.after(100, self._drain)

        def _pick_in(self):
            p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
            if p:
                self.excel.delete(0, "end")
                self.excel.insert(0, p)

        def _pick_out(self):
            p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if p:
                self.out.delete(0, "end")
                self.out.insert(0, p)

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
            excel, out = self.excel.get().strip(), self.out.get().strip()
            if not excel or not out:
                messagebox.showwarning("Thieu file", "Chon Excel input + output.")
                return
            args = ["--excel", excel, "--output", out, "--concurrency", self.conc.get().strip() or "3"]
            if self.headless.get():
                args.append("--headless")
            self.btn.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            threading.Thread(target=self._run, args=(args,), daemon=True).start()

        def _run(self, args):
            env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"}
            self.log_q.put(f"\n===== {' '.join(args)} =====\n")
            try:
                self.proc = subprocess.Popen(
                    [sys.executable, str(HERE / "vongbicongnghiep.py"), *args], cwd=str(HERE),
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
    ap = argparse.ArgumentParser(description="vongbicongnghiep.vn URL resolver (GUI default)")
    ap.add_argument("--excel")
    ap.add_argument("--output")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--concurrency", type=int, default=3)
    a = ap.parse_args()
    if a.excel and a.output:
        asyncio.run(crawl(a.excel, a.output, a.headless, a.concurrency))
    else:
        launch_gui()


if __name__ == "__main__":
    main()
