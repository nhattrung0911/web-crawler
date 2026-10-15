"""
skf.com — crawl product tables from a list of SKF product/listing URLs (Playwright).

Run:
  skf.bat                       -> GUI
  python skf.py                 -> GUI
  python skf.py --excel input_skf.xlsx [--output skf_report.xlsx]

Reads an Excel with a URL column, opens each page, extracts the product table
(breadcrumb -> cate_lv*, headers -> columns) across all pages, writes one xlsx.
Read-only (GET).
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


def _clean(t: str) -> str:
    return " ".join(t.replace("\xa0", " ").split())


async def crawl_url(page, url: str) -> list[dict]:
    results = []
    await page.goto(url, wait_until="networkidle")
    breadcrumb = [_clean(x) for x in await page.locator("ol.breadcrumb li").all_inner_texts() if _clean(x)]
    cate = {f"cate_lv{i}": v for i, v in enumerate(breadcrumb, 1)}
    await page.wait_for_selector("table.table-sm", timeout=20000)
    headers = [_clean(x) for x in await page.locator("table thead tr").last.locator("th").all_inner_texts()]
    page_index = 1
    while True:
        rows = page.locator("table tbody tr")
        for i in range(await rows.count()):
            cells = [_clean(x) for x in await rows.nth(i).locator("td").all_inner_texts()]
            rd = {**cate, "source_url": url, "page_index": page_index, "row_index": i + 1}
            rd.update(dict(zip(headers, cells)))
            results.append(rd)
        nxt = page.locator("nav.pagination .nav-item[rel='next']")
        if await nxt.count() == 0:
            break
        cls = await nxt.get_attribute("class")
        if cls and "disabled" in cls:
            break
        await nxt.click()
        await page.wait_for_selector("table tbody tr", timeout=20000)
        page_index += 1
    return results


async def crawl(excel: str, output: str = "skf_report.xlsx"):
    df = pd.read_excel(excel)
    url_col = next((c for c in df.columns if "url" in c.lower()), df.columns[0])
    urls = df[url_col].dropna().unique().tolist()
    all_rows = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True, args=["--disable-gpu", "--no-sandbox"])
        ctx = await browser.new_context()
        await ctx.route("**/*", lambda r: r.abort()
                        if r.request.resource_type in ("image", "media", "font") else r.continue_())
        page = await ctx.new_page()
        for idx, url in enumerate(urls, 1):
            print(f"[{idx}/{len(urls)}] {url}", flush=True)
            try:
                all_rows.extend(await crawl_url(page, url))
            except PWTimeout:
                print(f"  timeout: {url}", flush=True)
            except Exception as e:
                print(f"  error: {url} -> {e}", flush=True)
        await browser.close()
    dest = (HERE / output) if not os.path.isabs(output) else Path(output)
    pd.DataFrame(all_rows).to_excel(dest, index=False)
    print(f"DONE -> {dest} ({len(all_rows)} rows)", flush=True)


# ---------------- GUI ----------------

def launch_gui():
    import customtkinter as ctk
    from tkinter import filedialog, messagebox

    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title("SKF Crawler (skf.com)")
            self.geometry("860x540")
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(2, weight=1)
            self.proc = None
            self.log_q: queue.Queue = queue.Queue()

            f = ctk.CTkFrame(self)
            f.grid(row=0, column=0, padx=16, pady=(16, 6), sticky="ew")
            f.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(f, text="Excel (cot URL)").grid(row=0, column=0, padx=8, pady=8, sticky="w")
            self.excel = ctk.CTkEntry(f, placeholder_text="chon .xlsx")
            self.excel.grid(row=0, column=1, padx=8, sticky="ew")
            ctk.CTkButton(f, text="...", width=40, command=self._pick).grid(row=0, column=2, padx=6)
            self.btn = ctk.CTkButton(f, text="START", width=120, command=self.start)
            self.btn.grid(row=0, column=3, padx=6)
            self.btn_stop = ctk.CTkButton(f, text="STOP", width=90, fg_color="#a33",
                                          hover_color="#c44", command=self.stop, state="disabled")
            self.btn_stop.grid(row=0, column=4, padx=6)

            self.status = ctk.CTkLabel(self, text="Output: skf_report.xlsx", anchor="w")
            self.status.grid(row=1, column=0, padx=18, sticky="ew")
            self.console = ctk.CTkTextbox(self, state="disabled", font=("Consolas", 12))
            self.console.grid(row=2, column=0, padx=16, pady=12, sticky="nsew")
            self.after(100, self._drain)

        def _pick(self):
            p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
            if p:
                self.excel.delete(0, "end")
                self.excel.insert(0, p)

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
            excel = self.excel.get().strip()
            if not excel:
                messagebox.showwarning("Thieu file", "Chon Excel (cot URL).")
                return
            self.btn.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            threading.Thread(target=self._run, args=(["--excel", excel],), daemon=True).start()

        def _run(self, args):
            env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"}
            self.log_q.put(f"\n===== skf {' '.join(args)} =====\n")
            try:
                self.proc = subprocess.Popen(
                    [sys.executable, str(HERE / "skf.py"), *args], cwd=str(HERE),
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
    ap = argparse.ArgumentParser(description="skf.com crawler (GUI default)")
    ap.add_argument("--excel")
    ap.add_argument("--output", default="skf_report.xlsx")
    a = ap.parse_args()
    if a.excel:
        asyncio.run(crawl(a.excel, a.output))
    else:
        launch_gui()


if __name__ == "__main__":
    main()
