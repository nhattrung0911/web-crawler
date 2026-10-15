"""totalpowertools.in crawler.

Recon 2026-08-13: Lightspeed/Ecwid store 62425538. Public v3 API is 403, but
the storefront uses POST /storefront/api/v1/62425538/catalog/product. Sitemap
has product clean URLs and is the live master list. Detail JSON contains SKU,
price, stock, category paths, HTML description/spec and original image URLs.
Benchmark: HTTP detail at concurrency 4 was 24/24 200; 8 produced 19/24 429.
"""
from __future__ import annotations
import argparse, asyncio, gc, json, os, pathlib, re, threading, time
from typing import Any
from urllib.parse import urlparse
import httpx
import pandas as pd
from lxml import etree, html
import sys

HERE = pathlib.Path(__file__).parent
BASE = "https://totalpowertools.in"
API = "https://eu-fra3-storefront-api.ecwid.com/storefront/api/v1/62425538/catalog/product"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122 Safari/537.36"
CIRCUIT_BREAK, MAX_ITEMS, MAX_BFS_PAGES, MAX_PAGES, MAX_REDIRECTS = 25, 5000, 1500, 600, 5
TIMEOUT, DEFAULT_CONC = 35.0, 4

def norm_code(s: Any) -> str:
    v = re.sub(r"\s+", "", str(s or "").strip())
    return v if v.lower().startswith("http") else v.upper()

def parse_sitemap(text: str) -> list[str]:
    root = etree.fromstring(text.encode())
    out, seen = [], set()
    for loc in root.xpath("//*[local-name()='loc']/text()"):
        if "/products/" in loc and re.search(r"-p\d+$", loc) and loc not in seen:
            seen.add(loc); out.append(loc)
    return out

def product_id(url: str) -> int | None:
    m = re.search(r"-p(\d+)(?:$|[?#])", url)
    return int(m.group(1)) if m else None

def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()

def description_spec(value: str) -> tuple[dict[str, str], str]:
    doc = html.fromstring(value or "<div></div>")
    lines = [clean_text(x) for x in doc.xpath("//li//text()") if clean_text(x)]
    if not lines: lines = [clean_text(x) for x in doc.xpath("//text()") if clean_text(x)]
    spec, notes = {}, []
    for line in lines:
        if line.lower().startswith(("about this item", "item name")): continue
        k, sep, v = line.partition(":")
        if sep and k and v and len(k) <= 60: spec[k.strip()] = v.strip()
        else: notes.append(line)
    return spec, " | ".join(notes)

def parse_product(data: Any, url: str) -> dict:
    # HTML fallback keeps unit tests and harmless metadata pages readable.
    if isinstance(data, str):
        doc = html.fromstring(data); title = clean_text(" ".join(doc.xpath("//h1//text()")))
        sku = re.search(r"\bSKU\s*[:#]?\s*([\w.-]+)", doc.text_content(), re.I)
        image = (doc.xpath("//meta[@property='og:image']/@content") or [""])[0]
        if image.startswith("//"): image = "https:" + image
        spec, notes = description_spec(data)
        cats = [clean_text(x) for x in doc.xpath("//a[contains(@href,'-c')]/text()") if clean_text(x)]
        return {"Input":"", "Item No":sku.group(1) if sku else "", "Ten":title, "Status":"OK" if title else "Failed", "Gia":"", "Ton kho":"", "Ghi chu":notes, "_cats":cats, "_spec":spec, "Image":image, "URL":url}
    overrides = ((data.get("defaultOptionsOverrides") or {}).get("variationOverrides") or {})
    price = ((data.get("defaultOptionsOverrides") or {}).get("pricesOverrides") or {})
    spec, notes = description_spec(data.get("description") or "")
    paths = data.get("categoryPaths") or []
    default = next((x.get("categoryPath") for x in paths if x.get("isDefault")), paths[0].get("categoryPath", []) if paths else [])
    cats = [x.get("name", "") for x in default if x.get("name")]
    images = [x.get("imageOriginalUrl") or x.get("image1500pxUrl") for x in overrides.get("mediaItems") or [] if x.get("type") == "PICTURE"]
    return {"Input":"", "Item No":overrides.get("sku") or "", "Ten":data.get("name") or "", "Status":"OK", "Gia":price.get("basePrice", ""), "Ton kho":"Out of stock" if overrides.get("isSoldOut") else "In stock", "Ghi chu":notes, "_cats":cats, "_spec":spec, "Image":" | ".join(images[:5]), "URL":url}

def export_excel(results: list[dict], path: pathlib.Path) -> pathlib.Path:
    max_cat=max((len(r.get("_cats") or []) for r in results),default=0); rows=[]
    for r in results:
        row={k:r.get(k,"") for k in ("Input","Ten","Item No","Status","Gia","Ton kho")}
        for i in range(max_cat): row[f"Category {i+1}"]=(r.get("_cats") or [""]*max_cat)[i] if i<len(r.get("_cats") or []) else ""
        row["Thong so"]="\n".join([f"{k}: {v}" for k,v in (r.get("_spec") or {}).items()]+[x for x in (r.get("Ghi chu") or "").split(" | ") if x])
        row["Image"],row["URL"]=r.get("Image",""),r.get("URL",""); rows.append(row)
    with pd.ExcelWriter(path,engine="openpyxl") as xw:
        pd.DataFrame(rows).to_excel(xw,index=False); ws=xw.sheets["Sheet1"]; ws.column_dimensions["H"].width=60
        for c in ws["H"][1:]: c.alignment=c.alignment.copy(wrap_text=True,vertical="top")
    return path

class Engine:
    def __init__(self, conc=DEFAULT_CONC, delay=.15, log=print, stop_evt=None, pause_evt=None):
        self.conc,self.delay,self.log=max(1,min(int(conc),4)),max(0.,delay),log; self.stop_evt=stop_evt or threading.Event(); self.pause_evt=pause_evt or threading.Event(); self.results=[];self.seen=set();self.fails=0;self.key="";self.total=0;self.done=0
    def ckpt(self,key): return HERE/f"_raw_{re.sub(r'[^A-Za-z0-9]+','_',key)[:40]}.jsonl"
    def load(self,key):
        p=self.ckpt(key)
        if p.exists():
            for line in p.read_text('utf8',errors='ignore').splitlines():
                try: r=json.loads(line); self.results.append(r); self.seen.add(r['URL'])
                except Exception: pass
            self.log(f"[SYS] Resume {len(self.results)} records")
    def add(self,r):
        if r.get('URL') in self.seen:return
        self.seen.add(r['URL']);self.results.append(r);self.done+=1
        with self.ckpt(self.key).open('a',encoding='utf8') as f:f.write(json.dumps(r,ensure_ascii=False)+'\n');f.flush()
    async def fetch_product(self,cli,url):
        pid=product_id(url)
        if not pid:return {"Input":"","Item No":"","Ten":"","Status":"Not Found","Gia":"","Ton kho":"","Ghi chu":"URL khong co product ID","_cats":[],"_spec":{},"Image":"","URL":url}
        body={"lang":"en","productIdentifier":{"type":"PUBLISHED","productId":pid},"urlParams":{"urlType":"CLEAN_URL","baseUrl":"/products","canonicalBaseUrl":BASE+"/products","isCleanUrls":True,"isCanonicalUrlsEnabled":True,"isSlugsWithoutIds":False,"isTrailingSlash":False}}
        for attempt in range(4):
            while self.pause_evt.is_set() and not self.stop_evt.is_set():await asyncio.sleep(.2)
            if self.stop_evt.is_set():return None
            try:
                r=await cli.post(API,json=body,headers={"Referer":BASE+"/"})
                if r.status_code==404:return {"Input":"","Item No":"","Ten":"","Status":"Not Found","Gia":"","Ton kho":"","Ghi chu":"API 404","_cats":[],"_spec":{},"Image":"","URL":url}
                if r.status_code in (429,403,503): await asyncio.sleep(min(5*2**attempt,60));continue
                r.raise_for_status();self.fails=0;return parse_product(r.json(),url)
            except Exception as e:
                if attempt<3:await asyncio.sleep(2**attempt);continue
                self.fails+=1; self.log(f"[ERR] {type(e).__name__} {url}")
                if self.fails>=CIRCUIT_BREAK:self.stop_evt.set()
        return {"Input":"","Item No":"","Ten":"","Status":"Failed","Gia":"","Ton kho":"","Ghi chu":"retry exhausted","_cats":[],"_spec":{},"Image":"","URL":url}
    async def sitemap(self,cli):
        r=await cli.get(BASE+'/sitemap.xml');r.raise_for_status();return parse_sitemap(r.text)
    async def run(self,mode,target='',limit=0):
        self.key=f"{mode}_{target or 'all'}";self.load(self.key)
        async with httpx.AsyncClient(headers={"User-Agent":UA,"Origin":BASE},timeout=TIMEOUT,follow_redirects=True,http2=True) as cli:
            urls=target.splitlines() if mode=='link' else await self.sitemap(cli)
            url_code={}  # map url -> code da khop, de gan Input cho dong ket qua
            if mode=='code':
                codes={norm_code(x) for x in target.splitlines() if norm_code(x)}
                url_codes={}  # map url -> TAT CA ma khop (1 slug co the khop nhieu ma)
                for u in urls:
                    U=u.upper()  # ban sao IN HOA chi de so sanh, khong doi u goc
                    m={c for c in codes if re.search(rf"(?<![A-Z0-9]){re.escape(c)}(?![A-Z0-9])",U)}
                    if m:url_codes[u]=m
                urls=[u for u in urls if u in url_codes]
                url_code={u:sorted(m,key=lambda c:(-len(c),c))[0] for u,m in url_codes.items()}  # Input: ma dai nhat khop, hoa giai bang sorted -> deterministic
                found={c for m in url_codes.values() for c in m}
                for c in codes-found:self.add({"Input":c,"Item No":"","Ten":"","Status":"Not Found","Gia":"","Ton kho":"","Ghi chu":"Khong thay SKU trong sitemap live","_cats":[],"_spec":{},"Image":"","URL":f"{BASE}/products/search?keyword={c}"})
                for c in found-set(url_code.values()):self.add({"Input":c,"Item No":"","Ten":"","Status":"Not Found","Gia":"","Ton kho":"","Ghi chu":"Ma chi xuat hien trong slug cua san pham khac","_cats":[],"_spec":{},"Image":"","URL":f"{BASE}/products/search?keyword={c}"})
            elif mode=='keyword':urls=[u for u in urls if norm_code(target) in norm_code(u)]
            if limit:urls=urls[:limit]
            self.total=len(urls)
            sem=asyncio.Semaphore(self.conc)
            async def one(u):
                async with sem:
                    if u in self.seen:return
                    r=await self.fetch_product(cli,u)
                    if r:r["Input"]=url_code.get(u,"");self.add(r);await asyncio.sleep(self.delay)
                    elif u in url_code:self.add({"Input":url_code[u],"Item No":"","Ten":"","Status":"Not Found","Gia":"","Ton kho":"","Ghi chu":"Khong doc duoc trang chi tiet","_cats":[],"_spec":{},"Image":"","URL":u})
            await asyncio.gather(*(one(u) for u in urls[:MAX_ITEMS]))
        if not self.stop_evt.is_set() and self.ckpt(self.key).exists():self.ckpt(self.key).unlink()
        return self.results

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--mode',choices=['code','keyword','link','excel','all']);ap.add_argument('--code',default='');ap.add_argument('--keyword',default='');ap.add_argument('--link',default='');ap.add_argument('--excel');ap.add_argument('--sheet',default=0);ap.add_argument('--col');ap.add_argument('--limit',type=int,default=0);ap.add_argument('--out',default='') ;ap.add_argument('--conc',type=int,default=4);args=ap.parse_args()
    if not args.mode:
        run_gui(); return
    target={'code':args.code,'keyword':args.keyword,'link':args.link}.get(args.mode,'')
    if args.mode=='excel':
        if not args.excel or not args.col: ap.error('--excel and --col required')
        target='\n'.join(str(x) for x in pd.read_excel(args.excel,sheet_name=args.sheet)[args.col].dropna());args.mode='code'
    results=asyncio.run(Engine(args.conc).run(args.mode,target,args.limit));out=pathlib.Path(args.out or (pathlib.Path(args.excel).with_name(pathlib.Path(args.excel).stem+'_ketqua.xlsx') if args.excel else HERE/'totalpowertools_ketqua.xlsx'));export_excel(results,out);print(f'[OK] {len(results)} records -> {out}')

def run_gui() -> None:
    """Layout chuan WORX GREEN: pipeline trai, console/progress phai."""
    import tkinter as tk
    from tkinter import messagebox, ttk
    import queue
    root_dir = next(p for p in (HERE, *HERE.parents) if (p / 'gui_common.py').exists())
    sys.path.insert(0, str(root_dir))
    from gui_common import ask_export_path, attach_excel_picker

    root = tk.Tk(); root.title('TOTAL Power Tools India Crawler'); root.geometry('1080x680')
    ttk.Style().theme_use('clam')
    logs: queue.Queue[str] = queue.Queue()
    stop_evt, pause_evt = threading.Event(), threading.Event()
    state: dict[str, Any] = {'running': False, 'results': [], 'engine': None}
    left = ttk.Frame(root, padding=10); left.pack(side='left', fill='y')
    right = ttk.Frame(root, padding=(0, 10, 10, 10)); right.pack(side='right', fill='both', expand=True)
    mode = tk.StringVar(value='code')
    ttk.Label(left, text='Pipeline', font=('Segoe UI', 10, 'bold')).pack(anchor='w')
    for value, label in [('code', '1. Theo MA'), ('keyword', '2. Theo TU KHOA'), ('link', '3. Theo LINK (moi dong 1 link)'), ('excel', '4. Theo cot EXCEL'), ('all', '5. Crawl ALL (sitemap live)')]:
        ttk.Radiobutton(left, text=label, value=value, variable=mode).pack(anchor='w')
    ttk.Label(left, text='\nMa / Tu khoa:').pack(anchor='w')
    entry = ttk.Entry(left, width=36); entry.pack(anchor='w')
    ttk.Label(left, text='Link (moi dong 1 link):').pack(anchor='w')
    link_box = tk.Text(left, width=36, height=5); link_box.pack(anchor='w')
    xls, sheet, col, outdir = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar(value='images')
    attach_excel_picker(left, xls, sheet, col, outdir)
    conc, delay, limit = tk.IntVar(value=4), tk.DoubleVar(value=.15), tk.IntVar(value=0)
    for label, var in [('Concurrency (max 4):', conc), ('Delay (s):', delay), ('Limit (0=all):', limit)]:
        row = ttk.Frame(left); row.pack(anchor='w', pady=1)
        ttk.Label(row, text=label, width=20).pack(side='left'); ttk.Entry(row, textvariable=var, width=8).pack(side='left')
    console = tk.Text(right, bg='#12161c', fg='#d6e2ee', insertbackground='#fff', wrap='none')
    console.pack(fill='both', expand=True)
    progress = ttk.Progressbar(right, mode='determinate'); progress.pack(fill='x', pady=4)
    stats = ttk.Label(right, text='0/0 | 0 record'); stats.pack(anchor='w')

    def drain() -> None:
        while not logs.empty(): console.insert('end', logs.get() + '\n'); console.see('end')
        engine: Engine | None = state['engine']
        if engine:
            progress['maximum'] = max(1, engine.total); progress['value'] = min(engine.done, max(1, engine.total))
            stats.config(text=f'{engine.done}/{engine.total} | {len(engine.results)} record')
        root.after(100, drain)

    def worker() -> None:
        current = mode.get(); target = link_box.get('1.0', 'end').strip() if current == 'link' else entry.get().strip()
        try:
            if current == 'excel': target = '\n'.join(str(x) for x in pd.read_excel(xls.get(), sheet_name=sheet.get())[col.get()].dropna()); current = 'code'
            engine = Engine(conc.get(), delay.get(), logs.put, stop_evt, pause_evt); state['engine'] = engine
            asyncio.run(engine.run(current, target, limit.get()))
        except BaseException as exc: logs.put(f'[FATAL] {type(exc).__name__}: {exc}')
        finally:
            engine = state['engine']; state['results'] = engine.results if engine else []; state['running'] = False
            logs.put(f"[SYS] XONG -- {len(state['results'])} record. Bam EXPORT de xuat Excel.")

    def start() -> None:
        if state['running']: messagebox.showinfo('Dang chay', 'Dang crawl. Bam Stop truoc.'); return
        current = mode.get()
        if current == 'excel' and not (xls.get() and col.get()): messagebox.showwarning('Thieu du lieu', 'Chon Excel va cot truoc.'); return
        if current in ('code', 'keyword') and not entry.get().strip(): messagebox.showwarning('Thieu du lieu', 'Nhap ma / tu khoa truoc.'); return
        if current == 'link' and not link_box.get('1.0', 'end').strip(): messagebox.showwarning('Thieu du lieu', 'Dan link vao o (moi dong 1 link).'); return
        stop_evt.clear(); pause_evt.clear(); state['running'] = True; console.delete('1.0', 'end'); threading.Thread(target=worker, daemon=True).start()

    def export() -> None:
        if not state['results']: messagebox.showinfo('Export', 'Chua co du lieu'); return
        path = ask_export_path(xls.get(), HERE, 'totalpowertools.xlsx')
        if path: export_excel(state['results'], pathlib.Path(path)); messagebox.showinfo('Export', f'Da luu:\n{path}')
    controls = ttk.Frame(left); controls.pack(anchor='w', pady=8)
    ttk.Button(controls, text='START', command=start).pack(side='left')
    ttk.Button(controls, text='Pause', command=lambda: pause_evt.clear() if pause_evt.is_set() else pause_evt.set()).pack(side='left')
    ttk.Button(controls, text='Stop', command=stop_evt.set).pack(side='left')
    ttk.Button(controls, text='EXPORT', command=export).pack(side='left')
    def on_close() -> None:
        if state['running'] and not messagebox.askyesno('Thoat', 'Dang crawl. Thoat va tat luon?'): return
        stop_evt.set()
        try: root.destroy()
        finally: os._exit(0)
    root.protocol('WM_DELETE_WINDOW', on_close); drain(); root.mainloop()
if __name__=='__main__':main()
