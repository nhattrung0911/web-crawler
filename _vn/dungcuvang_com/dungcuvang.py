"""dungcuvang.com — WordPress REST API-first crawler (recon 2026-08-10)."""
from __future__ import annotations
import argparse, asyncio, html, json, os, pathlib, queue, re, sys, threading
from typing import Any
import httpx, pandas as pd
from selectolax.lexbor import LexborHTMLParser

HERE=pathlib.Path(__file__).parent; BASE="https://dungcuvang.com"; API=BASE+"/wp-json/wp/v2/product"
HEADERS={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36","Accept-Language":"vi-VN,vi;q=0.9"}
CIRCUIT_BREAK=25; MAX_ITEMS=20000; MAX_PAGES=600; TIMEOUT=30.0

def clean(s: Any)->str:
    return re.sub(r"\s+([,.;:!?])",r"\1"," ".join(html.unescape(str(s or "")).replace("\xa0"," ").split()))
def norm(s: Any)->str:
    v=str(s or "").strip(); return v if v.lower().startswith("http") else re.sub(r"\s+","",v).upper()
def text(node: Any)->str: return clean(node.text(separator=" ",strip=True) if node else "")
def parse_api_product(p: dict)->dict:
    tree=LexborHTMLParser(p.get("content",{}).get("rendered","") or "")
    spec={}
    for tr in tree.css("tr"):
        cells=tr.css("th,td")
        if len(cells)==2 and text(cells[0]) and text(cells[1]): spec.setdefault(text(cells[0]).rstrip(":"),text(cells[1]))
    desc="\n".join(dict.fromkeys(text(x) for x in tree.css("p,h1,h2,h3,h4,li") if text(x))) or text(tree)
    emb=p.get("_embedded",{}) or {}; media=emb.get("wp:featuredmedia",[]) or []; terms=emb.get("wp:term",[]) or []
    cats=[clean(t.get("name")) for group in terms for t in group if clean(t.get("name"))]
    return {"Input":"","Item No":"","SKU":"","Page":clean(p.get("title",{}).get("rendered")),"Status":"OK","Thuong hieu":"","Model":"","Gia (VND)":None,"Ton kho":"","_cats":list(dict.fromkeys(cats)),"_spec":spec,"Mo ta":desc,"Image":" | ".join(m.get("source_url","") for m in media if m.get("source_url")),"URL":p.get("link","")}
def parse_pdp(record:dict, html_page:str)->dict:
    tree=LexborHTMLParser(html_page); product=None; cats=[]
    for n in tree.css('script[type="application/ld+json"]'):
        try: data=json.loads(n.text() or "")
        except Exception: continue
        items=data.get("@graph",[]) if isinstance(data,dict) and "@graph" in data else [data]
        for d in items:
            if d.get("@type")=="Product": product=d
            if d.get("@type")=="BreadcrumbList": cats=[x.get("item",{}).get("name","") for x in d.get("itemListElement",[])[1:-1] if x.get("item",{}).get("name")]
    if not product: return record
    offer=product.get("offers",{}) or {}; brand=product.get("brand",{}) or {}
    record.update({"Item No":clean(product.get("mpn") or product.get("sku")),"SKU":clean(product.get("sku")),"Model":clean(product.get("mpn")),"Thuong hieu":clean(brand.get("name") if isinstance(brand,dict) else brand),"Gia (VND)":int(re.sub(r"\D","",str(offer.get("price",""))) or 0) or None,"Ton kho":clean(str(offer.get("availability",""))).rsplit("/",1)[-1]})
    record["_cats"]=list(dict.fromkeys([*record["_cats"],*cats])); imgs=product.get("image",[]) or []
    if isinstance(imgs,str): imgs=[imgs]
    record["Image"]=" | ".join(dict.fromkeys([*filter(None,record["Image"].split(" | ")),*imgs])); return record
def exact(code:str,r:dict)->bool:
    c=norm(code); hay=" ".join(str(r.get(k,"")) for k in ("Page","Item No","SKU","Model","URL")).upper(); return bool(re.search(rf"(?<![A-Z0-9]){re.escape(c)}(?![A-Z0-9])",hay))
def not_found(inp:str,status="Not Found")->dict: return {"Input":inp,"Item No":inp,"SKU":"","Page":"","Status":status,"Thuong hieu":"","Model":"","Gia (VND)":None,"Ton kho":"","_cats":[],"_spec":{},"Mo ta":"","Image":"","URL":""}
def export_excel(rows:list[dict], path:pathlib.Path)->pathlib.Path:
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter
    maxcat=max((len(r["_cats"]) for r in rows),default=0); out=[]
    for r in rows:
        d={k:r.get(k,"") for k in ("Input","Item No","SKU","Page","Status","Thuong hieu","Model","Gia (VND)","Ton kho")}
        d.update({f"Category {i+1}":r["_cats"][i] if i<len(r["_cats"]) else "" for i in range(maxcat)}); d["Thong so"]="\n".join(f"{k}: {v}" for k,v in r["_spec"].items()); d["Mo ta"]=r["Mo ta"]; d["Image"]=r["Image"]; d["URL"]=r["URL"]; out.append(d)
    df=pd.DataFrame(out)
    with pd.ExcelWriter(path,engine="openpyxl") as w:
        df.to_excel(w,index=False); ws=w.sheets["Sheet1"]; al=Alignment(wrap_text=True,vertical="top")
        for name,width in (("Thong so",60),("Mo ta",80)):
            col=df.columns.get_loc(name)+1; ws.column_dimensions[get_column_letter(col)].width=width
            for cell in list(ws.columns)[col-1][1:]: cell.alignment=al
    return path
class Engine:
 def __init__(self,conc=8,log=print,stop_evt=None,pause_evt=None): self.conc=max(1,conc); self.log=log; self.stop_evt=stop_evt or threading.Event(); self.pause_evt=pause_evt or threading.Event(); self.results=[]; self.done=0; self.total=0; self.fails=0; self.key=""; self.loaded=False
 def ckpt(self): return HERE/f"_raw_{re.sub(r'[^A-Za-z0-9]+','_',self.key)[:40]}.jsonl"
 def load(self):
  if self.loaded:return {r.get("URL") for r in self.results}
  self.loaded=True; seen=set(); p=self.ckpt()
  if p.exists():
   for line in p.read_text("utf8",errors="ignore").splitlines():
    try:r=json.loads(line)
    except:continue
    if r.get("Status")=="OK":self.results.append(r);seen.add(r.get("URL"))
  return seen
 def save(self,rows):
  with self.ckpt().open("a",encoding="utf8") as f:
   for r in rows:f.write(json.dumps(r,ensure_ascii=False)+"\n")
 def clean(self):
  if self.key and self.ckpt().exists() and not self.stop_evt.is_set():self.ckpt().unlink()
 async def get(self,c,url,params=None):
  for i in range(4):
   while self.pause_evt.is_set() and not self.stop_evt.is_set():await asyncio.sleep(.2)
   if self.stop_evt.is_set():return None
   try:
    r=await c.get(url,params=params);
    if r.status_code in (403,429,503):
     wait=min(5*2**i,60);self.log(f"[WARN] {r.status_code} -> ngu {wait}s | {url[-60:]}");await asyncio.sleep(wait);continue
    r.raise_for_status();self.fails=0;return r
   except Exception as e:
    if i<3:await asyncio.sleep(2**i);continue
    self.fails+=1;self.log(f"[ERR] {type(e).__name__} | {url[-60:]}")
    if self.fails>=CIRCUIT_BREAK:self.log(f"[STOP] {CIRCUIT_BREAK} loi LIEN TIEP - nghi bi chan IP. DUNG.");self.stop_evt.set()
 async def detail(self,c,p,inp):
  r=parse_api_product(p);r["Input"]=inp; page=await self.get(c,r["URL"])
  return parse_pdp(r,page.text) if page else r
 async def run(self,mode,target="",excel="",sheet="",col="",limit=0):
  self.key=f"{mode}_{target or excel}_{col}"; seen=self.load()
  if seen:self.log(f"[SYS] Resume: {len(seen)} record da xong -> bo qua")
  async with httpx.AsyncClient(headers=HEADERS,timeout=TIMEOUT,follow_redirects=True,limits=httpx.Limits(max_connections=self.conc+2)) as c:
   if mode in ("code","keyword","excel"):
    vals=[norm(target)] if mode=="code" else ([target] if mode=="keyword" else [norm(x) for x in pd.read_excel(excel,sheet_name=sheet or 0,dtype=str)[col].dropna()])
    vals=vals[:MAX_ITEMS]; self.total=len(vals)
    self.log(f"[SYS] Mode {mode}: {self.total} input")
    for inp in vals:
     if inp.lower().startswith("http"):  # BUG1 fix: link input -> crawl truc tiep qua slug (nhu mode link), khong loc bang exact()
      slug=inp.strip().rstrip("/").rsplit("/",1)[-1]; data=await self.get(c,API,{"slug":slug,"per_page":1,"_embed":1}); pr=data.json() if data else []
      self.done+=1
      if not pr:self.results.append(not_found(inp));self.log(f"[MISS] {self.done}/{self.total} | {inp}");continue
      rec=await self.detail(c,pr[0],inp);self.results.append(rec);self.save([rec]);self.log(f"[OK ] {self.done}/{self.total} | {inp}");continue
     data=await self.get(c,API,{"search":inp,"per_page":100,"_embed":1}); products=data.json() if data else []
     hits=[p for p in products if mode=="keyword" or exact(inp,parse_api_product(p))]
     self.done+=1
     if not hits:self.results.append(not_found(inp));self.log(f"[MISS] {self.done}/{self.total} | {inp}");continue
     for p in hits:
      if p.get("link") in seen:continue
      rec=await self.detail(c,p,inp); self.results.append(rec);self.save([rec])
     self.log(f"[OK ] {self.done}/{self.total} +{len(hits)} | {inp}")
   elif mode=="link":
    urls=target.splitlines(); self.total=len(urls)
    self.log(f"[SYS] Mode link: {self.total} URL")
    for u in urls:
     slug=u.strip().rstrip("/").rsplit("/",1)[-1]; data=await self.get(c,API,{"slug":slug,"per_page":1,"_embed":1})
     self.done+=1
     if not data or not data.json():self.results.append(not_found(u));self.log(f"[MISS] {self.done}/{self.total} | {u[-60:]}");continue
     rec=await self.detail(c,data.json()[0],u);self.results.append(rec);self.save([rec])
     self.log(f"[OK ] {self.done}/{self.total} | {u[-60:]}")
   elif mode=="all":
    page=1
    self.log("[SYS] Mode all: quet API theo trang")
    while page<=MAX_PAGES and not self.stop_evt.is_set():
     r=await self.get(c,API,{"page":page,"per_page":100,"_embed":1}); products=r.json() if r else []
     if not products:break
     rows=[parse_api_product(p) for p in products if p.get("link") not in seen]; self.results.extend(rows);self.save(rows)
     self.done+=len(rows); self.log(f"[SYS] trang {page}: +{len(rows)} record (tong {self.done})")
     if len(self.results)>=(limit or MAX_ITEMS):self.results=self.results[:limit or MAX_ITEMS];break  # BUG3 fix: limit ap dung ca run, dung phan trang khi du
     page+=1
  return self.results

def run_gui():
 import tkinter as tk
 from tkinter import ttk, messagebox
 root_dir=next(p for p in (HERE,*HERE.parents) if (p/'gui_common.py').exists());sys.path.insert(0,str(root_dir))
 from gui_common import ask_export_path,attach_excel_picker
 root=tk.Tk();root.title('Dung Cu Vang Crawler');root.geometry('1000x650');ttk.Style().theme_use('clam')
 q=queue.Queue();stop,pause=threading.Event(),threading.Event();state={'run':False,'rows':[],'eng':None};left=ttk.Frame(root,padding=10);left.pack(side='left',fill='y');console=tk.Text(root,bg='#12161c',fg='#d6e2ee');console.pack(side='right',fill='both',expand=True)
 mode=tk.StringVar(value='code')
 for v,t in [('code','1. Theo MA'),('keyword','2. Theo TU KHOA'),('link','3. Theo LINK'),('excel','4. Theo EXCEL'),('all','5. Crawl ALL API')]:ttk.Radiobutton(left,text=t,value=v,variable=mode).pack(anchor='w')
 ent=ttk.Entry(left,width=36);ent.pack(pady=4);box=tk.Text(left,width=36,height=5);box.pack();xls,sheet,col,outdir=tk.StringVar(),tk.StringVar(),tk.StringVar(),tk.StringVar(value='images');attach_excel_picker(left,xls,sheet,col,outdir)
 def drain():
  while not q.empty():console.insert('end',q.get()+'\n');console.see('end')
  root.after(100,drain)
 def start():
  if state['run']:return
  m=mode.get();target=box.get('1.0','end').strip() if m=='link' else ent.get().strip()
  if m in ('code','keyword') and not target:messagebox.showwarning('Thieu du lieu','Nhap ma/tu khoa');return
  if m=='excel' and not(xls.get() and col.get()):messagebox.showwarning('Thieu du lieu','Chon Excel va cot');return
  state['run']=True;stop.clear()
  def worker():
   e=Engine(log=q.put,stop_evt=stop,pause_evt=pause);state['eng']=e
   try:state['rows']=asyncio.run(e.run(m,target,xls.get(),sheet.get(),col.get()))
   except BaseException as ex:q.put(f'[FATAL] {type(ex).__name__}: {ex}')
   finally:e.clean();state['run']=False;q.put(f'[SYS] XONG {len(state["rows"])} record')
  threading.Thread(target=worker,daemon=True).start()
 def export():
  if not state['rows']:return
  p=ask_export_path(xls.get(),HERE,'dungcuvang.xlsx')
  if p:export_excel(state['rows'],pathlib.Path(p))
 ttk.Button(left,text='START',command=start).pack(side='left');ttk.Button(left,text='Pause',command=lambda:pause.set() if not pause.is_set() else pause.clear()).pack(side='left');ttk.Button(left,text='Stop',command=stop.set).pack(side='left');ttk.Button(left,text='EXPORT',command=export).pack(side='left')
 def close():
  stop.set();root.destroy();os._exit(0)
 root.protocol('WM_DELETE_WINDOW',close);drain();root.mainloop()

def main():
 ap=argparse.ArgumentParser();ap.add_argument('--mode',choices=['code','keyword','link','excel','all']);ap.add_argument('--code');ap.add_argument('--keyword');ap.add_argument('--link');ap.add_argument('--excel');ap.add_argument('--sheet',default='');ap.add_argument('--col',default='');ap.add_argument('--out',default='dungcuvang.xlsx');ap.add_argument('--conc',type=int,default=8);ap.add_argument('--limit',type=int,default=0);a=ap.parse_args()
 if not a.mode:return run_gui()
 e=Engine(a.conc);rows=asyncio.run(e.run(a.mode,a.code or a.keyword or a.link or '',a.excel or '',a.sheet,a.col,a.limit));export_excel(rows,pathlib.Path(a.out));e.clean();print(f'[SYS] {len(rows)} record -> {a.out}')
if __name__=='__main__':main()
