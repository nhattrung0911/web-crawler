"""
dientudanxuan.py — crawler Điện Tử Đan Xuân (dientudanxuan.com).
Sử dụng httpx + selectolax (LexborHTMLParser), KHÔNG Selenium, KHÔNG browser.

RECON (2026-07-20)
------------------
CMS      : WordPress + Haravan (shop ID: 1000325124, cdn.hstatic.net).
API      : CÓ — Haravan JSON API (.js endpoint cho chi tiết sản phẩm):
             /products/<handle>.js -> JSON đầy đủ (variants, sku, barcode, price, description, images, vendor, type).
             Giá trị price trong .js được lưu theo đơn vị nhỏ nhất (nhân 100), nên chia 100 ra VND chuẩn.
robots   : Cho phép thu thập sitemap và sản phẩm.
sitemap  : /sitemap.xml -> /sitemap_products_1.xml, /sitemap_products_2.xml (tổng ~1,263 sản phẩm).
HTML     : Fallback khi cần hoặc khi có cờ --html: /products/<handle> -> parse h1.product-title, .product-price, #tab-description.
Tìm kiếm : /search?q=<kw>&type=product -> trả HTML danh sách liên kết /products/<handle>.

Chạy lệnh (CLI / GUI)
---------------------
  dientudanxuan.bat                                           -> Mở GUI
  python dientudanxuan.py --mode code --code CXR-41H
  python dientudanxuan.py --mode keyword --keyword "Ruột mỏ hàn" --limit 10
  python dientudanxuan.py --mode link --link "https://dientudanxuan.com/products/loi-nhiet-thay-the-goot-cxr-41h"
  python dientudanxuan.py --mode excel --excel parts.xlsx --col "Mã SP"
  python dientudanxuan.py --mode all [--limit 500]            -> Quét toàn bộ (~1.26k SP)
  python dientudanxuan.py --mode images --excel parts.xlsx --col "Mã SP" --outdir out
  thêm --html để ép parse HTML (bỏ qua .js API)
"""
from __future__ import annotations

import argparse
import asyncio
import gc
import json
import os
import pathlib
import queue
import re
import threading
import time
from typing import Any
from urllib.parse import quote, urljoin, urlparse

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser

HERE = pathlib.Path(__file__).parent
BASE = "https://dientudanxuan.com"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")

CIRCUIT_BREAK = 25        # N lỗi LIÊN TIẾP -> nghi bị block IP -> dừng
MAX_ITEMS = 10000         # Trang có ~1,263 SP
MAX_PAGES = 600
MAX_REDIRECTS = 5
TIMEOUT = 30.0
GC_EVERY = 500


# ================================================================ CORE (không dùng GUI)
def norm_code(s: Any) -> str:
    """Chuẩn hoá mã để so khớp. URL thì GIỮ NGUYÊN hoa/thường (path phân biệt chữ)."""
    v = re.sub(r"\s+", "", str(s or ""))
    return v if v.lower().startswith("http") else v.upper()


def slugify_sku(sku: str) -> str:
    """Chuyển mã/tên sang định dạng slug url Haravan cơ bản."""
    s = re.sub(r"[^a-z0-9]+", "-", sku.lower()).strip("-")
    return s


def slug_urls(code: str) -> list[str]:
    """Các ứng viên URL từ mã/slug."""
    s = slugify_sku(code)
    return [
        f"{BASE}/products/{s}",
        f"{BASE}/products/loi-nhiet-thay-the-goot-{s}",
        f"{BASE}/products/mo-han-goot-{s}"
    ]


def extract_html_meta(html: str) -> dict[str, str]:
    """Bóc tách metadata từ .product-meta và swatches (#variant-swatch-X) trên trang HTML."""
    t = LexborHTMLParser(html)
    res = {"SKU": "", "Ton kho": "", "Tinh trang": "", "Thuong hieu": "", "Bao hanh": "", "Xuat xu": "", "Ma vach": "", "Loai san pham": ""}

    meta_div = t.css_first("#detail-product > div.wrapbox-detail > div.product-heading > div, .product-heading .product-meta")
    if meta_div:
        sku_el = meta_div.css_first("#pro_sku strong, #pro_sku")
        if sku_el:
            res["SKU"] = re.sub(r"^SKU:\s*", "", sku_el.text(strip=True), flags=re.I).strip()
        status_el = meta_div.css_first(".pro-soldold strong, .pro-soldold")
        if status_el:
            tt_val = re.sub(r"^Tình trạng:\s*", "", status_el.text(strip=True), flags=re.I).strip()
            res["Tinh trang"] = tt_val
            if not res["Ton kho"]:
                res["Ton kho"] = "Còn hàng" if "sẵn" in tt_val.lower() or "còn" in tt_val.lower() else ("Hết hàng" if "hết" in tt_val.lower() else tt_val)
        vendor_el = meta_div.css_first(".pro-vendor strong a, .pro-vendor strong, .pro-vendor")
        if vendor_el:
            res["Thuong hieu"] = re.sub(r"^Thương hiệu:\s*", "", vendor_el.text(strip=True), flags=re.I).strip()
        bh_el = meta_div.css_first(".product__baohanh-time strong, .product__baohanh-time")
        if bh_el:
            res["Bao hanh"] = re.sub(r"^Bảo hành:\s*", "", bh_el.text(strip=True), flags=re.I).strip()

        for el in meta_div.css("span, div"):
            txt = re.sub(r"\s+", " ", el.text(strip=True))
            if ":" in txt:
                parts = txt.split(":", 1)
                k, v = parts[0].strip().lower(), parts[1].strip()
                if not v: continue
                if "sku" in k or "mã" in k:
                    if not res["SKU"]: res["SKU"] = v
                elif "tình trạng" in k or "status" in k or "tồn" in k:
                    if not res["Tinh trang"]: res["Tinh trang"] = v
                    if not res["Ton kho"]: res["Ton kho"] = "Còn hàng" if "sẵn" in v.lower() or "còn" in v.lower() else ("Hết hàng" if "hết" in v.lower() else v)
                elif "thương hiệu" in k or "hãng" in k or "vendor" in k:
                    if not res["Thuong hieu"]: res["Thuong hieu"] = v
                elif "bảo hành" in k or "baohanh" in k or "warranty" in k:
                    if not res["Bao hanh"]: res["Bao hanh"] = v
                elif "xuất xứ" in k or "origin" in k:
                    if not res["Xuat xu"]: res["Xuat xu"] = v

    for sw in t.css('.swatch, [id^="variant-swatch-"]'):
        hdr = sw.css_first('.header, .title-swap')
        if not hdr: continue
        opt_name = re.sub(r":\s*$", "", hdr.text(strip=True)).strip()
        if not opt_name: continue
        chk = sw.css_first('input:checked, input[checked]')
        val = chk.attributes.get("value") or chk.attributes.get("data-value") if chk else ""
        if not val:
            el = sw.css_first('.swatch-element, [data-value]')
            if el: val = el.attributes.get("data-value") or el.text(strip=True)
        if opt_name and val:
            kl = opt_name.lower()
            if kl in ("xuất xứ", "xuất xứ :", "xuat xu", "origin"):
                if not res["Xuat xu"]: res["Xuat xu"] = val
            elif kl in ("thương hiệu", "thương hiệu :", "thuong hieu", "hãng sản xuất"):
                if not res["Thuong hieu"]: res["Thuong hieu"] = val
            elif kl in ("bảo hành", "bảo hành :", "bao hanh", "warranty"):
                if not res["Bao hanh"]: res["Bao hanh"] = val
    return res


def parse_record_from_js(p: dict, url: str) -> list[dict]:
    """Parse JSON trả về từ endpoint /products/<handle>.js hoặc /collections/all/products.json."""
    desc = p.get("description") or p.get("body_html") or p.get("content") or ""
    t = LexborHTMLParser(desc)
    spec: dict[str, str] = {}

    vendor = str(p.get("vendor") or "").strip()
    ptype = p.get("type") or p.get("product_type") or ""
    loai = str(ptype).strip()
    xuat_xu = ""
    bao_hanh = ""
    ma_vach = ""
    sku_from_tbl = ""
    vat_rate = ""
    raw_tags = p.get("tags") or []
    tags_list = [str(tg).strip() for tg in raw_tags] if isinstance(raw_tags, list) else [tg.strip() for tg in str(raw_tags).split(",") if tg.strip()]
    for tg in tags_list:
        tg_low = tg.lower().strip()
        if tg_low.startswith("bh_") and not bao_hanh:
            bao_hanh = tg[3:].strip()
        elif tg_low in ("vat8", "novat"):
            vat_rate = "8%"
        elif tg_low == "vat10":
            vat_rate = "10%"
    if not vat_rate:
        m_rate = re.search(r"VAT\s*(\d+%)", str(p.get("body_html") or ""), re.I)
        if m_rate:
            vat_rate = m_rate.group(1)
    if not vat_rate:
        for tg in tags_list:
            tg_low = tg.lower().strip()
            if tg_low == "vat":
                vat_rate = "VAT"
                break

    options = p.get("options") or []
    opt_names = {}
    for opt in options:
        if isinstance(opt, dict) and opt.get("position"):
            opt_names[opt["position"]] = str(opt.get("name") or "").strip()

    for tb in t.css("table"):
        for tr in tb.css("tr"):
            cells = [re.sub(r"\s+", " ", td.text(strip=True)) for td in tr.css("th, td")]
            if len(cells) >= 2 and cells[0] not in ("Thông số", "Đặc tính", "Đặc tính kỹ thuật", "Thông số kỹ thuật", "Thuộc tính", "Tên thuộc tính"):
                k, v = cells[0].strip(), cells[1].strip()
                kl = k.lower()
                if kl in ("thương hiệu", "thương hiệu :", "hãng sản xuất"):
                    if not vendor: vendor = v
                elif kl in ("xuất xứ", "xuất xứ :"):
                    if not xuat_xu: xuat_xu = v
                elif kl in ("bảo hành", "bảo hành :"):
                    if not bao_hanh: bao_hanh = v
                elif kl in ("mã vạch", "mã vạch (barcode)", "barcode"):
                    if not ma_vach: ma_vach = v
                elif kl in ("mã sản phẩm", "mã sản phẩm :"):
                    if not sku_from_tbl: sku_from_tbl = v
                elif kl in ("loại sản phẩm", "loại sản phẩm :"):
                    if not loai: loai = v
                else:
                    spec[k] = v

    # Fallback: chỉ quét li/p khi KHÔNG có table spec (vd 5SI-131-I dùng bullet list)
    # Nếu table đã có data → li/p toàn mô tả/hướng dẫn sử dụng → bỏ qua
    if not spec:
        for el in t.css("li, p"):
            txt = re.sub(r"\s+", " ", el.text(strip=True))
            if ":" in txt:
                parts = txt.split(":", 1)
                k, v = parts[0].strip(), parts[1].strip()
                if not v or len(k) > 60:
                    continue
                kl = k.lower()
                if any(w in kl for w in ("lưu ý", "quan trọng", "vệ sinh", "hướng dẫn", "bảo quản", "thay thế", "khi tắt", "khi hàn", "trước khi", "trong khi")):
                    continue
                if kl in ("thương hiệu", "thương hiệu :", "hãng sản xuất"):
                    if not vendor: vendor = v
                elif kl in ("xuất xứ", "xuất xứ :"):
                    if not xuat_xu: xuat_xu = v
                elif kl in ("bảo hành", "bảo hành :"):
                    if not bao_hanh: bao_hanh = v
                elif kl in ("mã vạch", "mã vạch (barcode)", "barcode"):
                    if not ma_vach: ma_vach = v
                elif kl in ("mã sản phẩm", "mã sản phẩm :"):
                    if not sku_from_tbl: sku_from_tbl = v
                elif kl in ("loại sản phẩm", "loại sản phẩm :"):
                    if not loai: loai = v
                else:
                    if k not in spec:
                        spec[k] = v

    cats = [c.strip() for c in ptype.split(":") if c.strip()]
    if cats and not loai:
        loai = cats[-1]
    images = []
    for im in (p.get("images") or []):
        u = im.get("src") or im.get("url") if isinstance(im, dict) else im
        if not u or not isinstance(u, str):
            continue
        u = u.strip()
        if u.startswith("//"):
            u = "https:" + u
        if u not in images:
            images.append(u)
    if not images and p.get("image"):
        im = p.get("image")
        u = im.get("src") or im.get("url") if isinstance(im, dict) else im
        if u and isinstance(u, str):
            u = u.strip()
            if u.startswith("//"):
                u = "https:" + u
            if u not in images:
                images.append(u)

    records = []
    variants = p.get("variants") or [{}]
    for var in variants:
        sku = var.get("sku") or var.get("barcode") or p.get("sku") or sku_from_tbl or ""
        vtitle = var.get("title") or ""
        name = p.get("title") or ""
        if len(variants) > 1 and vtitle and vtitle.lower() not in ("default title", "default", name.lower()):
            full_name = f"{name} ({vtitle})"
        else:
            full_name = name
        pr = var.get("price") if var.get("price") is not None else p.get("price")
        pr_base_val = 0.0
        if isinstance(pr, (int, float)) and pr > 0:
            pr_base_val = float(pr) / 100.0
        elif isinstance(pr, str) and pr.strip():
            try:
                raw_num = float(pr.strip().replace(",", "").replace(".", "").replace("₫", ""))
                if raw_num > 0:
                    pr_base_val = raw_num
            except Exception:
                pr_base_val = 0.0

        if pr_base_val > 0:
            pr_vnd = f"{int(round(pr_base_val)):,}".replace(",", ".") + "₫"
        else:
            pr_vnd = str(pr or "") if pr else ""

        avail = var.get("available") if var.get("available") is not None else p.get("available")
        tinh_trang = "Hàng có sẵn" if avail else "Hết hàng"
        inv = var.get("inventory_quantity") if var.get("inventory_quantity") is not None else p.get("inventory_quantity")
        if avail:
            stock = "Còn hàng"
            if isinstance(inv, (int, float)) and inv > 0:
                stock += f" (sl: {int(inv)})"
        else:
            stock = "Hết hàng"
            if isinstance(inv, (int, float)) and inv > 0:
                stock += f" (sl: {int(inv)})"

        var_ma_vach = str(var.get("barcode") or ma_vach or "").strip()
        v_xuat_xu = xuat_xu
        v_vendor = vendor
        v_bao_hanh = bao_hanh
        v_loai = loai

        for pos, opt_name in opt_names.items():
            val = str(var.get(f"option{pos}") or "").strip()
            if not val and pos - 1 < len(options) and isinstance(options[pos - 1], dict) and options[pos - 1].get("values"):
                val = str(options[pos - 1]["values"][0]).strip()
            if opt_name and val and val.lower() not in ("default title", "default"):
                kl = opt_name.lower()
                if kl in ("xuất xứ", "xuất xứ :", "xuat xu", "origin"):
                    v_xuat_xu = val
                elif kl in ("thương hiệu", "thương hiệu :", "hãng sản xuất", "vendor"):
                    v_vendor = val
                elif kl in ("bảo hành", "bảo hành :", "warranty"):
                    v_bao_hanh = val
                elif kl in ("loại sản phẩm", "loại sản phẩm :", "loại"):
                    v_loai = val

        if vat_rate == "8%" and pr_base_val > 0:
            pr_vat_val = round(pr_base_val * 1.08)
            pr_vat_vnd = f"{int(pr_vat_val):,}".replace(",", ".") + "₫"
        elif vat_rate == "10%" and pr_base_val > 0:
            pr_vat_val = round(pr_base_val * 1.10)
            pr_vat_vnd = f"{int(pr_vat_val):,}".replace(",", ".") + "₫"
        elif vat_rate == "VAT":
            pr_vat_vnd = pr_vnd
        else:
            pr_vat_vnd = pr_vnd if pr_vnd else ""

        records.append({
            "Input": "",
            "SKU": sku,
            "Ten": full_name,
            "Gia (VND)": pr_vnd,
            "Gia VAT": pr_vat_vnd,
            "Ton kho": stock,
            "Tinh trang": tinh_trang,
            "Status": "OK",
            "Thuong hieu": v_vendor,
            "Xuat xu": v_xuat_xu,
            "Bao hanh": v_bao_hanh,
            "Ma vach": var_ma_vach,
            "Loai san pham": v_loai,
            "VAT": vat_rate,
            "_cats": cats,
            "_spec": dict(spec),
            "Image": " | ".join(images[:5]),
            "URL": url
        })
    return records


def parse_record_from_html(html: str, url: str) -> list[dict] | None:
    """Fallback parse HTML từ trang /products/<handle>, chuẩn theo selectors."""
    t = LexborHTMLParser(html)
    h1 = t.css_first("h1.product-title, h1")
    if not h1:
        return None
    name = re.sub(r"\s+", " ", h1.text(strip=True))

    meta = extract_html_meta(html)
    sku = meta.get("SKU", "")
    stock_str = meta.get("Ton kho", "")
    tinh_trang = meta.get("Tinh trang", "")
    vendor = meta.get("Thuong hieu", "")
    bao_hanh = meta.get("Bao hanh", "")
    xuat_xu = meta.get("Xuat xu", "")
    ma_vach = meta.get("Ma vach", "")
    loai = meta.get("Loai san pham", "")
    price_str = ""
    images = []

    # Thử lấy thêm từ LD+JSON nếu thiếu SKU / images
    for sc in t.css('script[type="application/ld+json"]'):
        try:
            data = json.loads(sc.text())
        except Exception:
            continue
        items = data.get("@graph", []) if isinstance(data, dict) else (data if isinstance(data, list) else [data])
        for blk in items:
            if not isinstance(blk, dict) or blk.get("@type") not in ("Product", ["Product"]):
                continue
            sku = sku or blk.get("sku") or blk.get("mpn") or ""
            if not vendor:
                if isinstance(blk.get("brand"), dict):
                    vendor = blk["brand"].get("name", "")
                elif isinstance(blk.get("brand"), str):
                    vendor = blk.get("brand", "")
            img = blk.get("image")
            for u in (img if isinstance(img, list) else [img]):
                if isinstance(u, dict):
                    u = u.get("url") or u.get("src")
                if not u or not isinstance(u, str):
                    continue
                u = u.strip()
                if u.startswith("//"):
                    u = "https:" + u
                if u not in images:
                    images.append(u)
            offers = blk.get("offers")
            if isinstance(offers, dict):
                pr = offers.get("price")
                if pr and not price_str:
                    try:
                        price_str = f"{int(float(pr)):,}".replace(",", ".") + "₫"
                    except Exception:
                        price_str = str(pr)
                if offers.get("availability") and not stock_str:
                    stock_str = "Còn hàng" if "InStock" in str(offers["availability"]) else "Hết hàng"

    if not images:
        for im in t.css(".product-image-feature img, #sliderproduct img, .product-gallery img"):
            u = im.attributes.get("data-src") or im.attributes.get("src") or ""
            if u and u not in images:
                if str(u).startswith("//"):
                    u = "https:" + str(u)
                images.append(urljoin(url, u))

    # 2. Extract base price (Gia VND) cleanly without confusing with #price-with-vat
    price_div = t.css_first("#detail-product > div.wrapbox-detail > div.product-price, .product-price")
    if price_div:
        for sel in ("#price-not-vat", ".pro-price.main", "span.pro-price.main"):
            el = price_div.css_first(sel)
            if el and el.text(strip=True) and any(c.isdigit() for c in el.text()):
                price_str = re.sub(r"\s+", " ", el.text(strip=True))
                break
    if not price_str:
        meta_pr = t.css_first('meta[property="og:price:amount"], meta[itemprop="price"]')
        if meta_pr:
            pr_val = meta_pr.attributes.get("content")
            if pr_val and any(c.isdigit() for c in str(pr_val)):
                try:
                    val_int = int(float(pr_val)) if "." in str(pr_val) else int(re.sub(r"[^\d]", "", str(pr_val)))
                    if val_int > 0:
                        price_str = f"{val_int:,}".replace(",", ".") + "₫"
                except Exception:
                    pass
    if not price_str and price_div:
        for sel in ("#price-with-vat", ".pro-price.vat", "span.pro-price"):
            el = price_div.css_first(sel)
            if el and el.text(strip=True) and any(c.isdigit() for c in el.text()):
                price_str = re.sub(r"\s+", " ", el.text(strip=True))
                break
        if not price_str and any(c.isdigit() for c in price_div.text(strip=True)):
            price_str = re.sub(r"\s+", " ", price_div.text(strip=True))
    if not price_str:
        pr_el = t.css_first(".product-price span, #price-preview, .price, [data-price]")
        if pr_el:
            pr_val = pr_el.attributes.get("data-price") or pr_el.text(strip=True)
            if pr_val and any(c.isdigit() for c in str(pr_val)):
                try:
                    val_int = int(float(pr_val)) if "." in str(pr_val) else int(re.sub(r"[^\d]", "", str(pr_val)))
                    if val_int > 100000000:
                        val_int = int(val_int / 100)
                    if val_int > 0:
                        price_str = f"{val_int:,}".replace(",", ".") + "₫"
                except Exception:
                    pass

    # 4. spec thì CHỈ lấy bảng ở đây: #nav-desc > div > div > div.description-productdetail > table
    spec: dict[str, str] = {}
    desc_table_els = t.css("#nav-desc > div > div > div.description-productdetail > table, .description-productdetail table, #nav-desc table")
    for tb in desc_table_els:
        for tr in tb.css("tr"):
            cells = [re.sub(r"\s+", " ", td.text(strip=True)) for td in tr.css("th, td")]
            if len(cells) >= 2 and cells[0] not in ("Thông số", "Đặc tính", "Đặc tính kỹ thuật", "Thông số kỹ thuật", "Thuộc tính", "Tên thuộc tính"):
                k, v = cells[0].strip(), cells[1].strip()
                kl = k.lower()
                if kl in ("thương hiệu", "thương hiệu :", "hãng sản xuất"):
                    if not vendor: vendor = v
                elif kl in ("xuất xứ", "xuất xứ :"):
                    if not xuat_xu: xuat_xu = v
                elif kl in ("bảo hành", "bảo hành :"):
                    if not bao_hanh: bao_hanh = v
                elif kl in ("mã vạch", "mã vạch (barcode)", "barcode"):
                    if not ma_vach: ma_vach = v
                elif kl in ("mã sản phẩm", "mã sản phẩm :"):
                    if not sku: sku = v
                elif kl in ("loại sản phẩm", "loại sản phẩm :"):
                    if not loai: loai = v
                else:
                    spec[k] = v

    # Fallback: chỉ quét li/p khi KHÔNG có table spec
    if not spec:
        desc_container = t.css_first("#nav-desc > div > div > div.description-productdetail, .description-productdetail, #nav-desc")
        if desc_container:
            for el in desc_container.css("li, p"):
                txt = re.sub(r"\s+", " ", el.text(strip=True))
                if ":" in txt:
                    parts = txt.split(":", 1)
                    k, v = parts[0].strip(), parts[1].strip()
                    if not v or len(k) > 60:
                        continue
                    kl = k.lower()
                    if any(w in kl for w in ("lưu ý", "quan trọng", "vệ sinh", "hướng dẫn", "bảo quản", "thay thế", "khi tắt", "khi hàn", "trước khi", "trong khi")):
                        continue
                    if kl in ("thương hiệu", "thương hiệu :", "hãng sản xuất"):
                        if not vendor: vendor = v
                    elif kl in ("xuất xứ", "xuất xứ :"):
                        if not xuat_xu: xuat_xu = v
                    elif kl in ("bảo hành", "bảo hành :"):
                        if not bao_hanh: bao_hanh = v
                    elif kl in ("mã vạch", "mã vạch (barcode)", "barcode"):
                        if not ma_vach: ma_vach = v
                    elif kl in ("mã sản phẩm", "mã sản phẩm :"):
                        if not sku: sku = v
                    elif kl in ("loại sản phẩm", "loại sản phẩm :"):
                        if not loai: loai = v
                    else:
                        if k not in spec:
                            spec[k] = v

    cats = [re.sub(r"\s+", " ", a.text(strip=True)) for a in t.css(".breadcrumb a, nav.breadcrumb a")][1:]
    if cats and not loai:
        loai = cats[-1]

    # 3. #detail-product > div.wrapbox-detail > div.product-variants
    variants_div = t.css_first("#detail-product > div.wrapbox-detail > div.product-variants, .product-variants")
    options = variants_div.css("select#product-select option") if variants_div else []

    pw_el = t.css_first("#vat-price-wrapper .pro-title, .vat-price .pro-title, .vat-note")
    vp_txt = pw_el.text(strip=True) if pw_el else ""
    m_rate = re.search(r"(\d+%)", vp_txt)
    if not m_rate and vp_txt and "gồm hoá đơn vat" in vp_txt.lower():
        vat_rate = "VAT"
    elif m_rate:
        vat_rate = m_rate.group(1)
    else:
        desc_el = t.css_first("#detail-product .description-productdetail, #nav-desc")
        if desc_el:
            m_rate = re.search(r"VAT\s*(\d+%)", desc_el.text(), re.I)
        if m_rate:
            vat_rate = m_rate.group(1)
        else:
            p_tags = ""
            for s_sc in t.css("script"):
                st = s_sc.text() or ""
                if ("productJson" in st or "HaravanProduct" in st or "meta" in st) and (url.split("/")[-1] in st or "id" in st):
                    m_tg = re.search(r'["\']tags["\']\s*:\s*\[([^\]]*)\]', st, re.I)
                    if m_tg:
                        p_tags = m_tg.group(1)
                        break
            if re.search(r'["\'](vat8|novat)["\']', p_tags, re.I) or re.search(r'pro_vat\s*=\s*[\'"]novat[\'"]', html, re.I):
                vat_rate = "8%"
            elif re.search(r'["\']vat10["\']', p_tags, re.I):
                vat_rate = "10%"
            elif re.search(r'["\']VAT["\']', p_tags) or re.search(r'pro_vat\s*=\s*[\'"]vat[\'"]', html, re.I):
                vat_rate = "VAT"
            else:
                vat_rate = ""

    pw_price_el = t.css_first("#price-with-vat, #vat-price-wrapper .pro-price.vat, .vat-price .pro-price.vat")
    pw_price_str = pw_price_el.text(strip=True).replace(",", ".") if pw_price_el else ""
    base_num = float(re.sub(r"[^\d]", "", price_str)) if re.search(r"\d", price_str) else 0.0
    vat_num = float(re.sub(r"[^\d]", "", pw_price_str)) if re.search(r"\d", pw_price_str) else 0.0

    if not pw_price_str:
        if vat_rate == "8%" and base_num > 0:
            vat_num = round(base_num * 1.08)
            pw_price_str = f"{int(vat_num):,}".replace(",", ".") + "₫"
        elif vat_rate == "10%" and base_num > 0:
            vat_num = round(base_num * 1.10)
            pw_price_str = f"{int(vat_num):,}".replace(",", ".") + "₫"
        elif vat_rate == "VAT":
            pw_price_str = price_str
            vat_num = base_num

    if vat_num > base_num > 0 and (not vat_rate or vat_rate in ("VAT", "NOVAT")):
        calc_pct = round((vat_num - base_num) / base_num * 100)
        if calc_pct > 0:
            vat_rate = f"{calc_pct}%"

    if options and len(options) > 1:
        records = []
        for opt in options:
            vtitle = re.sub(r"\s+", " ", opt.text(strip=True)).split(" - ")[0].strip()
            if vtitle.lower() not in ("default title", "default", name.lower()):
                full_name = f"{name} ({vtitle})"
            else:
                full_name = name
            vprice = price_str
            dp = opt.attributes.get("data-price")
            if dp and dp.isdigit() and int(dp) > 0:
                vprice = f"{int(int(dp) / 100):,}".replace(",", ".") + "₫"
                base_var = int(dp) / 100
            else:
                base_var = base_num
            if vat_rate == "8%":
                vvat = f"{int(round(base_var * 1.08)):,}".replace(",", ".") + "₫"
            elif vat_rate == "10%":
                vvat = f"{int(round(base_var * 1.10)):,}".replace(",", ".") + "₫"
            elif vat_rate == "VAT":
                vvat = vprice
            else:
                vvat = pw_price_str or vprice
            records.append({
                "Input": "",
                "SKU": sku,
                "Ten": full_name,
                "Gia (VND)": vprice or price_str,
                "Gia VAT": vvat,
                "Ton kho": stock_str or "Còn hàng",
                "Tinh trang": tinh_trang or stock_str or "Hàng có sẵn",
                "Status": "OK",
                "Thuong hieu": vendor,
                "Xuat xu": xuat_xu,
                "Bao hanh": bao_hanh,
                "Ma vach": ma_vach,
                "Loai san pham": loai,
                "VAT": vat_rate,
                "_cats": cats,
                "_spec": dict(spec),
                "Image": " | ".join(images[:5]),
                "URL": url
            })
        return records

    return [{
        "Input": "",
        "SKU": sku,
        "Ten": name,
        "Gia (VND)": price_str,
        "Gia VAT": pw_price_str or (price_str if vat_rate == "VAT" else ""),
        "Ton kho": stock_str or "Còn hàng",
        "Tinh trang": tinh_trang or stock_str or "Hàng có sẵn",
        "Status": "OK",
        "Thuong hieu": vendor,
        "Xuat xu": xuat_xu,
        "Bao hanh": bao_hanh,
        "Ma vach": ma_vach,
        "Loai san pham": loai,
        "VAT": vat_rate,
        "_cats": cats,
        "_spec": spec,
        "Image": " | ".join(images[:5]),
        "URL": url
    }]


def export_excel(results: list[dict], path: pathlib.Path) -> pathlib.Path:
    """Spec gộp 1 CỘT 'Thong so' (bảng thông số kĩ thuật). Thương hiệu, xuất xứ, bảo hành, mã vạch, loại là cột riêng."""
    from openpyxl.utils import get_column_letter

    max_cat = max((len(r.get("_cats") or []) for r in results), default=0)
    rows = []
    for r in results:
        row = {
            "Input": r.get("Input", ""),
            "SKU": r.get("SKU", ""),
            "Ten": r.get("Ten", ""),
            "Gia (VND)": r.get("Gia (VND)", ""),
            "Gia VAT": r.get("Gia VAT", ""),
            "Ton kho": r.get("Ton kho", ""),
            "Tinh trang": r.get("Tinh trang", ""),
            "Status": r.get("Status", ""),
            "Thuong hieu": r.get("Thuong hieu", ""),
            "Xuat xu": r.get("Xuat xu", ""),
            "Bao hanh": r.get("Bao hanh", ""),
            "Ma vach": r.get("Ma vach", ""),
            "Loai san pham": r.get("Loai san pham", ""),
            "VAT": r.get("VAT", "")
        }
        cats = r.get("_cats") or []
        for i in range(max_cat):
            row[f"Category {i + 1}"] = cats[i] if i < len(cats) else ""
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
        from copy import copy
        for c in ws[col][1:]:
            na = copy(c.alignment)
            na.wrap_text = True
            na.vertical = "top"
            c.alignment = na
    return path


def _safe_log(msg: str) -> None:
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode("utf-8", errors="replace").decode("utf-8", errors="replace"))


class Engine:
    """Haravan .js API là chính; lỗi/chặn -> tự động rơi sang HTML."""

    def __init__(self, conc: int = 6, delay: float = 0.1, log=print, force_html: bool = False,
                 stop_evt: threading.Event | None = None,
                 pause_evt: threading.Event | None = None,
                 fresh: bool = True) -> None:
        self.conc, self.delay = max(1, conc), max(0.0, delay)
        if log is print:
            self.log = _safe_log
        else:
            self.log = log
        self.stop_evt = stop_evt or threading.Event()
        self.pause_evt = pause_evt or threading.Event()
        self.api_ok = not force_html
        self.force_html = force_html
        self.fresh = fresh
        self.fails = 0
        self.results: list[dict] = []
        self.seen_urls: set[str] = set()
        self.done = 0
        self.total = 0
        self.key = ""
        self._ckpt_loaded = False

    def add(self, recs: list[dict], key: str) -> list[dict]:
        """Thêm record mới (bỏ trùng theo URL + SKU) + ghi checkpoint."""
        fresh = []
        for r in recs:
            u = r.get("URL") or ""
            sku = r.get("SKU") or ""
            dedup_key = f"{u}|{sku}"
            if dedup_key in self.seen_urls:
                continue
            self.seen_urls.add(dedup_key)
            fresh.append(r)
        if fresh:
            self.results += fresh
            self._save(key, fresh)
        return fresh

    def ckpt(self, key: str) -> pathlib.Path:
        return HERE / f"_raw_{re.sub(r'[^A-Za-z0-9]+', '_', key)[:40]}.jsonl"

    def _load_ckpt(self, key: str) -> set[str]:
        p, seen = self.ckpt(key), set()
        if self._ckpt_loaded:
            return set(self.seen_urls)
        self._ckpt_loaded = True
        if not p.exists():
            return seen
        for line in p.read_text("utf-8", errors="ignore").splitlines():
            try:
                r = json.loads(line)
            except Exception:
                continue
            self.results.append(r)
            dedup_key = f"{r.get('URL', '')}|{r.get('SKU', '')}"
            seen.add(dedup_key)
            self.seen_urls.add(dedup_key)
        if seen:
            self.log(f"[SYS] Resume: {len(self.results)} record đã có -> bỏ qua")
        return seen

    def _save(self, key: str, recs: list[dict]) -> None:
        with self.ckpt(key).open("a", encoding="utf-8") as f:
            for r in recs:
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
            f.flush()

    def cleanup(self) -> None:
        p = self.ckpt(self.key) if self.key else None
        if p and p.exists() and not self.stop_evt.is_set():
            p.unlink()

    async def _gate(self) -> bool:
        while self.pause_evt.is_set() and not self.stop_evt.is_set():
            await asyncio.sleep(0.2)
        return not self.stop_evt.is_set()

    async def fetch(self, cli: httpx.AsyncClient, url: str, tries: int = 4) -> httpx.Response | None:
        for i in range(tries):
            if not await self._gate():
                return None
            try:
                r = await cli.get(url)
                if r.status_code in (403, 429, 503):
                    wait = min(2 ** i * 5, 60)
                    self.log(f"[WARN] {r.status_code} -> ngủ {wait}s | {url[-55:]}")
                    await asyncio.sleep(wait)
                    continue
                if r.status_code == 404:
                    return None
                if "/products/" in url and "/products/" not in str(r.url):
                    # Redirect ra khỏi product (ra homepage/search)
                    return None
                r.raise_for_status()
                self.fails = 0
                return r
            except Exception as e:
                if i < tries - 1:
                    await asyncio.sleep(min(2 ** i, 20))
                    continue
                self.fails += 1
                self.log(f"[ERR] {type(e).__name__} | {url[-55:]}")
                if self.fails >= CIRCUIT_BREAK:
                    self.log(f"[STOP] {CIRCUIT_BREAK} lỗi LIÊN TIẾP - nghi bị chặn IP. DỪNG.")
                    self.stop_evt.set()
                return None
        return None

    async def fetch_product(self, cli: httpx.AsyncClient, url: str) -> list[dict]:
        clean_url = url.split("?")[0].split("#")[0].rstrip("/")
        if not clean_url.startswith("http"):
            clean_url = urljoin(BASE, clean_url)
        if self.api_ok and not self.force_html:
            js_url = clean_url + ".js"
            resp = await self.fetch(cli, js_url)
            if resp and resp.status_code == 200:
                try:
                    data = resp.json()
                    if isinstance(data, dict) and (data.get("title") or data.get("handle")):
                        recs = parse_record_from_js(data, clean_url)
                        if recs:
                            return recs
                except Exception:
                    self.log(f"[WARN] Lỗi parse JSON .js -> chuyển HTML cho {clean_url[-40:]}")
        # fallback to html
        resp = await self.fetch(cli, clean_url)
        if not resp or resp.status_code != 200:
            return []
        recs = parse_record_from_html(resp.text, clean_url)
        return recs or []

    async def _one_code(self, cli: httpx.AsyncClient, code: str) -> dict | None:
        want = norm_code(code)
        # 1. Thử direct slug
        for u in slug_urls(code):
            recs = await self.fetch_product(cli, u)
            if recs:
                for r in recs:
                    if norm_code(r.get("SKU")) == want or norm_code(r.get("_spec", {}).get("Mã vạch (Barcode)")) == want:
                        r["Input"] = code
                        return r
                # Nếu không khớp chính xác nhưng có sản phẩm ở slug direct
                r = recs[0]
                r["Input"] = code
                if want not in norm_code(r.get("SKU")):
                    r["Status"] = f"OK (lệch mã: {r.get('SKU')})"
                return r

        # 2. Thử tìm kiếm Haravan HTML /search?q=...&type=product
        search_url = f"{BASE}/search?q={quote(code)}&type=product"
        resp = await self.fetch(cli, search_url)
        if resp and resp.status_code == 200:
            t = LexborHTMLParser(resp.text)
            links = []
            for a in t.css("a"):
                href = a.attributes.get("href") or ""
                if "/products/" in href:
                    clean = urljoin(BASE, href.split("?")[0].split("#")[0].rstrip("/"))
                    if clean not in links:
                        links.append(clean)
            for l in links[:5]:
                recs = await self.fetch_product(cli, l)
                for r in recs:
                    if norm_code(r.get("SKU")) == want or norm_code(r.get("_spec", {}).get("Mã vạch (Barcode)")) == want:
                        r["Input"] = code
                        return r
                for r in recs:
                    if want in norm_code(r.get("SKU")) or (r.get("SKU") and norm_code(r.get("SKU")) in want):
                        r["Input"] = code
                        r["Status"] = f"OK (lệch mã: {r.get('SKU')})"
                        return r
            if links:
                recs = await self.fetch_product(cli, links[0])
                if recs:
                    r = recs[0]
                    r["Input"] = code
                    r["Status"] = f"OK (lệch mã: {r.get('SKU')})"
                    return r
        return None

    async def _fetch_urls(self, cli: httpx.AsyncClient, urls: list[str], key: str, limit: int = 0) -> None:
        expanded: list[str] = []
        for u in urls:
            if "/products/" in u:
                expanded.append(u)
            elif "/collections/" in u or "/search" in u:
                self.log(f"[SYS] Khám phá danh mục/tìm kiếm: {u}")
                clean_u = re.sub(r"([?&])page=\d+", r"\1", u).rstrip("?&")
                sep = "&" if "?" in clean_u else "?"
                page = 1
                while page <= MAX_PAGES and await self._gate():
                    p_url = f"{clean_u}{sep}page={page}" if page > 1 else clean_u
                    resp = await self.fetch(cli, p_url)
                    if not resp or resp.status_code != 200:
                        break
                    t = LexborHTMLParser(resp.text)
                    page_urls = []
                    for a in t.css("a"):
                        href = a.attributes.get("href") or ""
                        if "/products/" in href:
                            clean_p = urljoin(BASE, href.split("?")[0].split("#")[0].rstrip("/"))
                            if clean_p not in page_urls and clean_p not in expanded:
                                page_urls.append(clean_p)
                    if not page_urls:
                        break
                    expanded.extend(page_urls)
                    self.log(f"[SYS] {clean_u.rsplit('/', 1)[-1][:35]} trang {page} -> thêm {len(page_urls)} SP (tổng {len(expanded)})")
                    if limit and len(expanded) >= limit:
                        break
                    page += 1
                    if self.delay:
                        await asyncio.sleep(self.delay)
            else:
                resp = await self.fetch(cli, u)
                if resp and resp.status_code == 200:
                    t = LexborHTMLParser(resp.text)
                    page_urls = [urljoin(BASE, a.attributes.get("href", "").split("?")[0].split("#")[0].rstrip("/"))
                                 for a in t.css("a") if a.attributes.get("href") and "/products/" in a.attributes.get("href")]
                    page_urls = list(dict.fromkeys(page_urls))
                    if page_urls:
                        expanded.extend(page_urls)
                        self.log(f"[SYS] Trang hub {u[:40]} -> tìm thấy {len(page_urls)} SP")
                    else:
                        expanded.append(u)
                else:
                    expanded.append(u)
            if limit and len(expanded) >= limit:
                break
        urls = expanded[:limit] if limit else expanded
        seen = self._load_ckpt(key)
        urls = [u for u in dict.fromkeys(urls) if u not in seen][:MAX_ITEMS]
        self.total, self.done = len(urls) + len(seen), len(seen)
        sem = asyncio.Semaphore(self.conc)

        async def one(u: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                if self.delay:
                    await asyncio.sleep(self.delay)
                recs = await self.fetch_product(cli, u)
                if recs:
                    for r in recs:
                        r["Input"] = r["Input"] or u
                    self.add(recs, key)
                self.done += 1
                name_log = recs[0]["Ten"][:40] if recs else ""
                self.log(f"[{'OK ' if recs else 'MISS'}] {self.done}/{self.total} | {name_log or u[-45:]}")
                if self.done % GC_EVERY == 0:
                    gc.collect()

        try:
            await asyncio.gather(*(one(u) for u in urls))
        except (KeyboardInterrupt, Exception) as e:
            self.log(f"[FATAL] {type(e).__name__}: {e}")

    async def _sitemap_product_urls(self, cli: httpx.AsyncClient) -> list[str]:
        sm = await self.fetch(cli, f"{BASE}/sitemap.xml")
        if sm is None:
            return []
        sms = re.findall(r"<loc>(https://dientudanxuan\.com/sitemap_products_\d+\.xml)</loc>", sm.text)
        sem = asyncio.Semaphore(self.conc)

        async def one(sub_sm: str) -> list[str]:
            async with sem:
                r = await self.fetch(cli, sub_sm)
                got = [u.split("?")[0].split("#")[0].rstrip("/")
                       for u in re.findall(r"<loc>(https://dientudanxuan\.com/products/[^<]+)</loc>", r.text)] if r else []
                self.log(f"[SYS] {sub_sm.rsplit('/', 1)[-1]} -> {len(got)} URL")
                return got

        res = [u for part in await asyncio.gather(*(one(s) for s in sms)) for u in part]
        return list(dict.fromkeys(res))

    async def run(self, mode: str, target: str = "", excel: str = "", sheet: str = "",
                  col: str = "", limit: int = 0, outdir: str = "out") -> list[dict]:
        key = self.key = f"{mode}_{target or excel}_{col}"
        if self.fresh:
            p = self.ckpt(key)
            if p.exists():
                try:
                    p.unlink()
                except Exception:
                    pass
        headers = {"User-Agent": UA, "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8"}
        async with httpx.AsyncClient(headers=headers, timeout=TIMEOUT, follow_redirects=True,
                                     max_redirects=MAX_REDIRECTS,
                                     limits=httpx.Limits(max_connections=self.conc + 4)) as cli:
            if mode == "all":
                self.log("[SYS] Crawl ALL qua bulk API /collections/all/products.json ...")
                page = 1
                bulk_ok = False
                seen_all = self._load_ckpt(key)
                self.total = 1263 + len(seen_all)
                while page <= MAX_PAGES and await self._gate():
                    p_url = f"{BASE}/collections/all/products.json?limit=250&page={page}"
                    resp = await self.fetch(cli, p_url)
                    if not resp or resp.status_code != 200:
                        break
                    try:
                        prods = resp.json().get("products", [])
                    except Exception:
                        break
                    if not prods:
                        break
                    bulk_ok = True
                    page_recs = []
                    for p in prods:
                        clean_p = urljoin(BASE, f"/products/{p.get('handle')}")
                        if clean_p in seen_all:
                            continue
                        seen_all.add(clean_p)
                        recs = parse_record_from_js(p, clean_p)
                        if recs:
                            for r in recs:
                                r["Input"] = r["Input"] or clean_p
                            page_recs.extend(recs)
                    if page_recs:
                        self.add(page_recs, key)
                    self.done += len(prods)
                    self.log(f"[SYS] Bulk API trang {page} -> lấy {len(prods)} SP (đã xử lý {self.done} SP)")
                    if limit and self.done >= limit:
                        break
                    page += 1
                    if self.delay:
                        await asyncio.sleep(self.delay)
                if not bulk_ok:
                    self.log("[WARN] Bulk API không phản hồi -> chuyển sang sitemap...")
                    urls = await self._sitemap_product_urls(cli)
                    if limit:
                        urls = urls[:limit]
                    self.log(f"[SYS] Crawl ALL: {len(urls)} sản phẩm từ sitemap")
                    await self._fetch_urls(cli, urls, key, limit=limit)

            elif mode in ("code", "excel"):
                raw = [norm_code(target)] if mode == "code" else read_excel_col(excel, sheet, col)
                raw = [c for c in dict.fromkeys(raw) if c][:MAX_ITEMS]
                links = [c for c in raw if c.lower().startswith("http")]
                codes = [c for c in raw if c not in links]
                if links:
                    await self._fetch_urls(cli, links, key, limit=limit)
                self._load_ckpt(key)
                self.total = len(codes) + self.done
                sem = asyncio.Semaphore(self.conc)

                async def one(c: str) -> None:
                    async with sem:
                        if not await self._gate():
                            return
                        rec = await self._one_code(cli, c)
                        if self.delay:
                            await asyncio.sleep(self.delay)
                        self.done += 1
                        if rec:
                            self.add([rec], key)
                            self.log(f"[OK ] {self.done}/{self.total} {c} -> {rec['Ten'][:45]}")
                        else:
                            self.results.append({"Input": c, "SKU": c, "Ten": "", "Gia (VND)": "", "Gia VAT": "",
                                                  "Ton kho": "", "Status": "Not Found", "VAT": "",
                                                  "_cats": [], "_spec": {}, "Image": "", "URL": ""})
                            self.log(f"[MISS] {self.done}/{self.total} {c}")

                await asyncio.gather(*(one(c) for c in codes
                                       if c not in {r.get("Input") for r in self.results}))

            elif mode == "keyword":
                urls: list[str] = []
                page = 1
                while page <= MAX_PAGES and await self._gate():
                    search_url = f"{BASE}/search?q={quote(target)}&type=product&page={page}"
                    resp = await self.fetch(cli, search_url)
                    if not resp or resp.status_code != 200:
                        break
                    t = LexborHTMLParser(resp.text)
                    page_urls = []
                    for a in t.css("a"):
                        href = a.attributes.get("href") or ""
                        if "/products/" in href:
                            clean = urljoin(BASE, href.split("?")[0].split("#")[0].rstrip("/"))
                            if clean not in page_urls and clean not in urls:
                                page_urls.append(clean)
                    if not page_urls:
                        break
                    urls.extend(page_urls)
                    self.log(f"[SYS] Tìm kiếm '{target}' trang {page} -> tìm thấy thêm {len(page_urls)} SP (tổng {len(urls)})")
                    if limit and len(urls) >= limit:
                        break
                    page += 1
                    if self.delay:
                        await asyncio.sleep(self.delay)

                if not urls:
                    self.log(f"[MISS] 0 kết quả tìm kiếm cho từ khoá '{target}'.")
                else:
                    if limit:
                        urls = urls[:limit]
                    await self._fetch_urls(cli, urls, key, limit=limit)
                    for r in self.results:
                        r["Input"] = target

            elif mode == "link":
                links = [u.strip() for u in target.splitlines() if u.strip().lower().startswith("http")]
                await self._fetch_urls(cli, links, key, limit=limit)

            elif mode == "images":
                await self._download_images(cli, excel, sheet, col, outdir)

            else:
                self.log(f"[ERR] Mode không hợp lệ: {mode}")
        return self.results

    async def _download_images(self, cli: httpx.AsyncClient, excel: str, sheet: str,
                               col: str, outdir: str) -> None:
        """Tải ảnh từ cột Excel (mã hoặc URL) về outdir/<SKU>/<SKU>_i.jpg."""
        root = pathlib.Path(outdir) if os.path.isabs(outdir) else HERE / outdir
        root.mkdir(parents=True, exist_ok=True)
        keys = read_excel_col(excel, sheet, col)[:MAX_ITEMS]
        self.total = len(keys)
        sem = asyncio.Semaphore(self.conc)

        async def one(k: str) -> None:
            async with sem:
                if not await self._gate():
                    return
                recs = None
                if k.lower().startswith("http"):
                    recs = await self.fetch_product(cli, k)
                else:
                    r = await self._one_code(cli, k)
                    recs = [r] if r else None
                self.done += 1
                if not recs or not recs[0].get("Image"):
                    self.log(f"[MISS] {self.done}/{self.total} không có ảnh cho: {k}")
                    return
                rec = recs[0]
                name = re.sub(r'[\\/*?:"<>|]', "_", (rec.get("SKU") or k))[:60]
                folder = root / name
                folder.mkdir(parents=True, exist_ok=True)
                ok = 0
                for i, iu in enumerate(rec["Image"].split(" | "), 1):
                    try:
                        rr = await cli.get(iu, headers={"Referer": rec.get("URL") or BASE})
                        rr.raise_for_status()
                        ext = pathlib.Path(urlparse(iu).path).suffix or ".jpg"
                        (folder / f"{name}_{i}{ext}").write_bytes(rr.content)
                        ok += 1
                    except Exception as e:
                        self.log(f"[ERR] tải ảnh {iu[-45:]}: {type(e).__name__}")
                self.log(f"[OK ] {self.done}/{self.total} {name} -> {ok} ảnh")

        await asyncio.gather(*(one(k) for k in keys))


def read_excel_col(path: str, sheet: str, col: str) -> list[str]:
    if not path:
        raise ValueError("Chưa chọn file Excel")
    df = pd.read_excel(path, sheet_name=sheet or 0, dtype=str)
    df.columns = df.columns.str.strip()
    if col not in df.columns:
        raise ValueError(f"Không thấy cột '{col}'. Cột có sẵn: {list(df.columns)}")
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
    root.title("Điện Tử Đan Xuân Crawler — dientudanxuan.com (Haravan .js API + HTML fallback)")
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
    for val, txt in [("code", "1. Theo MÃ"), ("keyword", "2. Theo TỪ KHÓA"),
                     ("link", "3. Theo LINK (mỗi dòng 1 link)"), ("excel", "4. Theo cột EXCEL"),
                     ("all", "5. Crawl ALL (~1.26k SP)"), ("images", "6. Tải ẢNH")]:
        ttk.Radiobutton(left, text=txt, value=val, variable=mode).pack(anchor="w")

    ttk.Label(left, text="\nMã / Từ khoá:").pack(anchor="w")
    ent = ttk.Entry(left, width=36)
    ent.pack(anchor="w")
    ttk.Label(left, text="Link (mỗi dòng 1 link):").pack(anchor="w")
    box = tk.Text(left, width=36, height=5)
    box.pack(anchor="w")

    xls, sheet, col = tk.StringVar(), tk.StringVar(), tk.StringVar()
    outdir = tk.StringVar(value="out")
    attach_excel_picker(left, xls, sheet, col, outdir)

    conc, delay, limit = tk.IntVar(value=6), tk.DoubleVar(value=0.1), tk.IntVar(value=0)
    for lbl, var in [("Concurrency:", conc), ("Delay (s):", delay), ("Limit (0=all):", limit)]:
        f3 = ttk.Frame(left)
        f3.pack(anchor="w", pady=1)
        ttk.Label(f3, text=lbl, width=13).pack(side="left")
        ttk.Entry(f3, textvariable=var, width=8).pack(side="left")
    force_html = tk.BooleanVar(value=False)
    ttk.Checkbutton(left, text="Ép đi đường HTML (bỏ Haravan .js API)", variable=force_html).pack(anchor="w")
    fresh = tk.BooleanVar(value=True)
    ttk.Checkbutton(left, text="Cào mới từ đầu (Tự xóa cache/ckpt cũ)", variable=fresh).pack(anchor="w")

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
        eng = Engine(conc.get(), delay.get(), q.put, force_html.get(), stop_evt, pause_evt, fresh.get())
        state["eng"] = eng
        try:
            asyncio.run(eng.run(m, target, xls.get(), sheet.get(), col.get(),
                                limit.get(), outdir.get()))
            eng.cleanup()
        except BaseException as e:
            q.put(f"[FATAL] {type(e).__name__}: {e}")
        finally:
            state["results"] = eng.results
            state["running"] = False
            q.put(f"[SYS] XONG — {len(eng.results)} record. Bấm EXPORT để xuất Excel.")

    def start() -> None:
        if state["running"]:
            return
        stop_evt.clear()
        pause_evt.clear()
        state["running"] = True
        console.delete("1.0", "end")
        threading.Thread(target=worker, daemon=True).start()

    def export() -> None:
        if not state["results"]:
            messagebox.showinfo("Export", "Chưa có dữ liệu")
            return
        p = ask_export_path(xls.get(), HERE, "dientudanxuan.xlsx")
        if p:
            export_excel(state["results"], pathlib.Path(p))
            messagebox.showinfo("Export", f"Đã lưu:\n{p}")

    fb = ttk.Frame(left)
    fb.pack(anchor="w", pady=8)
    ttk.Button(fb, text="START", command=start).pack(side="left")
    ttk.Button(fb, text="Pause", command=lambda: (pause_evt.clear() if pause_evt.is_set()
                                                  else pause_evt.set())).pack(side="left")
    ttk.Button(fb, text="Stop", command=stop_evt.set).pack(side="left")
    ttk.Button(fb, text="EXPORT", command=export).pack(side="left")

    def on_close() -> None:
        if state["running"] and not messagebox.askyesno("Thoát", "Đang crawl. Thoát và tắt luôn?"):
            return
        stop_evt.set()
        try:
            root.destroy()
        finally:
            os._exit(0)

    root.protocol("WM_DELETE_WINDOW", on_close)
    drain()
    root.mainloop()


# ================================================================ CLI ENTRY
def main() -> None:
    import sys
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    ap = argparse.ArgumentParser(description="dientudanxuan.vn Crawler (GUI/CLI)")
    ap.add_argument("--mode", choices=["code", "keyword", "link", "excel", "all", "images"],
                    help="Pipeline mode (mặc định mở GUI)")
    ap.add_argument("--code", default="", help="Mã SP cho mode code")
    ap.add_argument("--keyword", default="", help="Từ khoá cho mode keyword")
    ap.add_argument("--link", default="", help="URL sản phẩm cho mode link")
    ap.add_argument("--excel", default="", help="Path file Excel đầu vào")
    ap.add_argument("--sheet", default="", help="Sheet name")
    ap.add_argument("--col", default="Mã SP", help="Tên cột mã / URL")
    ap.add_argument("--conc", type=int, default=6, help="Số luồng đồng thời")
    ap.add_argument("--delay", type=float, default=0.1, help="Nghỉ giữa các request")
    ap.add_argument("--limit", type=int, default=0, help="Giới hạn số SP (0=all)")
    ap.add_argument("--outdir", default="out", help="Thư mục tải ảnh")
    ap.add_argument("--html", action="store_true", help="Ép parse HTML")
    ap.add_argument("--resume", action="store_true", help="Dùng lại checkpoint cũ (mặc định cào mới từ đầu)")
    ap.add_argument("--export", default="dientudanxuan.xlsx", help="File Excel đầu ra")
    a = ap.parse_args()

    if not a.mode:
        run_gui()
        return

    target = a.code or a.keyword or a.link
    eng = Engine(conc=a.conc, delay=a.delay, force_html=a.html, fresh=not a.resume)
    results = asyncio.run(eng.run(a.mode, target, a.excel, a.sheet, a.col, a.limit, a.outdir))
    eng.cleanup()

    if a.mode != "images" and results:
        export_excel(results, pathlib.Path(a.export))
        _safe_log(f"\n[XONG] Đã xuất {len(results)} dòng ra {a.export}")
    elif a.mode == "images":
        _safe_log(f"\n[XONG] Hoàn tất tải ảnh vào {a.outdir}")
    else:
        _safe_log("\n[MISS] Không có kết quả nào.")


if __name__ == "__main__":
    main()
