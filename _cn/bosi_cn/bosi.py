"""
Bosi (bosi.cn) crawler — all-in-one (crawler + GUI).

Run:
  bosi.bat                          -> open GUI (default)
  python bosi.py                    -> open GUI
  python bosi.py --start 1 --end 999 [--output bosi_data.csv]  -> crawl id range

Iterates product ids on https://www.bosi.cn/index/pro_list/id/{id}.html,
parses each product page (selectolax) and appends rows to a CSV.
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

import httpx
import pandas as pd
from selectolax.parser import HTMLParser

HERE = Path(__file__).parent
BASE = "https://www.bosi.cn"
SEMAPHORE = asyncio.Semaphore(15)
TIMEOUT = httpx.Timeout(15.0, connect=10.0)


async def fetch_product(client: httpx.AsyncClient, prod_id: int):
    url = f"{BASE}/index/pro_list/id/{prod_id}.html"
    async with SEMAPHORE:
        for attempt in range(3):
            try:
                res = await client.get(url)
                if res.status_code == 200:
                    return prod_id, url, res.text
                if res.status_code == 404:
                    return prod_id, url, None
            except httpx.RequestError as e:
                print(f"Error fetching {prod_id} (attempt {attempt + 1}): {e}", flush=True)
                await asyncio.sleep(2)
        return prod_id, url, None


def parse_html(html: str):
    tree = HTMLParser(html)
    data = {}
    breadcrumbs = tree.css(".conmmon_nav3 .left a")
    if not breadcrumbs:
        return None  # not a product page
    data["Breadcrumb"] = " > ".join(a.text(strip=True) for a in breadcrumbs)

    name_node = tree.css_first(".right.fr span")
    data["Name"] = name_node.text(strip=True) if name_node else ""

    item_no_node = tree.css_first(".right.fr h2")
    data["Item_No"] = item_no_node.text(strip=True).replace("产品货号：", "").strip() if item_no_node else ""

    images = []
    for img in tree.css(".gallery-top .swiper-slide img"):
        src = img.attributes.get("src")
        if src:
            images.append(f"{BASE}{src}")
    data["Images"] = " | ".join(images)

    detail_imgs = []
    for img in tree.css(".pro_bar2 .b1 img"):
        src = img.attributes.get("src")
        if src:
            detail_imgs.append(src if src.startswith("http") else f"{BASE}{src}")
    data["Detail_Images"] = " | ".join(detail_imgs)
    return data


async def crawl(start_id: int, end_id: int, output_file: str = "bosi_data.csv") -> None:
    out = HERE / output_file
    file_exists = out.is_file()
    saved = 0
    async with httpx.AsyncClient(timeout=TIMEOUT, verify=False) as client:
        tasks = [fetch_product(client, i) for i in range(start_id, end_id + 1)]
        for coro in asyncio.as_completed(tasks):
            prod_id, url, html = await coro
            if not html:
                continue
            parsed = parse_html(html)
            if parsed and parsed.get("Item_No"):
                parsed["URL"] = url
                parsed["ID"] = prod_id
                pd.DataFrame([parsed]).to_csv(
                    out, mode="a", header=not file_exists, index=False, encoding="utf-8-sig")
                file_exists = True
                saved += 1
                print(f"[{saved}] id {prod_id}: {parsed['Item_No']}", flush=True)
    print(f"Done. Saved {saved} products -> {out.name}", flush=True)


# =====================================================
# GUI
# =====================================================

def launch_gui() -> None:
    import customtkinter as ctk
    from tkinter import messagebox

    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title("Bosi Crawler (bosi.cn)")
            self.geometry("820x560")
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(2, weight=1)
            self.proc = None
            self.log_q: queue.Queue = queue.Queue()

            top = ctk.CTkFrame(self)
            top.grid(row=0, column=0, padx=16, pady=(16, 6), sticky="ew")
            ctk.CTkLabel(top, text="ID tu").grid(row=0, column=0, padx=(10, 4), pady=10)
            self.start = ctk.CTkEntry(top, width=90, placeholder_text="1")
            self.start.insert(0, "1")
            self.start.grid(row=0, column=1, padx=4)
            ctk.CTkLabel(top, text="den").grid(row=0, column=2, padx=4)
            self.end = ctk.CTkEntry(top, width=90, placeholder_text="999")
            self.end.insert(0, "999")
            self.end.grid(row=0, column=3, padx=4)
            self.btn_start = ctk.CTkButton(top, text="START", width=110, command=self.start_crawl)
            self.btn_start.grid(row=0, column=4, padx=(16, 4))
            self.btn_stop = ctk.CTkButton(top, text="STOP", width=90, fg_color="#a33",
                                          hover_color="#c44", command=self.stop, state="disabled")
            self.btn_stop.grid(row=0, column=5, padx=4)

            self.status = ctk.CTkLabel(self, text="San sang. Output: bosi_data.csv", anchor="w")
            self.status.grid(row=1, column=0, padx=18, sticky="ew")
            self.console = ctk.CTkTextbox(self, state="disabled", font=("Consolas", 12))
            self.console.grid(row=2, column=0, padx=16, pady=12, sticky="nsew")

            self.after(100, self._drain)

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

        def start_crawl(self):
            if self.proc:
                return
            try:
                s, e = int(self.start.get()), int(self.end.get())
                assert s <= e
            except (ValueError, AssertionError):
                messagebox.showwarning("Sai range", "ID tu/den phai la so, tu <= den.")
                return
            self.btn_start.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            threading.Thread(target=self._run, args=(s, e), daemon=True).start()

        def _run(self, s, e):
            env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"}
            self.log_q.put(f"\n===== crawl id {s}-{e} =====\n")
            try:
                self.proc = subprocess.Popen(
                    [sys.executable, str(HERE / "bosi.py"), "--start", str(s), "--end", str(e)],
                    cwd=str(HERE), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace", env=env)
            except OSError as ex:
                self.log_q.put(f"[launch error] {ex}\n")
            else:
                for line in self.proc.stdout:
                    self.log_q.put(line)
                self.proc.wait()
                self.log_q.put(f"[exit {self.proc.returncode}]\n")
            self.proc = None
            self.after(0, lambda: (self.btn_start.configure(state="normal"),
                                   self.btn_stop.configure(state="disabled")))

        def stop(self):
            if self.proc:
                self.proc.terminate()
                self.log_q.put("\n[da dung]\n")

    App().mainloop()


def main() -> None:
    ap = argparse.ArgumentParser(description="Bosi crawler (GUI default)")
    ap.add_argument("--start", type=int, help="start id")
    ap.add_argument("--end", type=int, help="end id")
    ap.add_argument("--output", default="bosi_data.csv", help="output CSV")
    a = ap.parse_args()
    if a.start is not None and a.end is not None:
        asyncio.run(crawl(a.start, a.end, a.output))
    else:
        launch_gui()


if __name__ == "__main__":
    main()
