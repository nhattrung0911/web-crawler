# ==========================================================
# HYBRID CRAWLER FULL (PRODUCTION)
# Stage 1: HTTP async (fast) + Playwright fallback (accurate)
# Stage 2: Download images async from Stage 1 report
#
# PIPELINE MENU:
#   1 - Chỉ crawl dữ liệu (Stage 1)
#   2 - Chỉ tải ảnh từ report có sẵn (Stage 2)
#   3 - Crawl dữ liệu + tải ảnh
#   0 - Thoát
#
# NOTE:
# - GIỮ NGUYÊN LOGIC & KIẾN TRÚC PHIÊN BẢN B
# - CHỈ THÊM MENU + CONTROL FLOW
# ==========================================================

import asyncio
import os
import re
import time
from tkinter import Tk, filedialog
from urllib.parse import urlparse

import aiohttp
import pandas as pd
from lxml import html as LH
from playwright.async_api import async_playwright

# =========================
# CONFIG
# =========================
OUTPUT_DIR = "output3"
REPORT_XLSX = os.path.join(OUTPUT_DIR, "report_stage1.xlsx")
TMP_CSV = os.path.join(OUTPUT_DIR, "_stage1_tmp.csv")

IMG_PRODUCT_DIR = os.path.join(OUTPUT_DIR, "images", "product")
IMG_DRAWING_DIR = os.path.join(OUTPUT_DIR, "images", "drawing")

# Stage 1 - HTTP
HTTP_CONCURRENCY = 40
HTTP_BATCH_SIZE = 3000
HTTP_TIMEOUT = aiohttp.ClientTimeout(total=30)

# Stage 1 - Playwright
PW_CONCURRENCY = 12
PW_TIMEOUT_MS = 25_000

# Stage 2 - Download
DL_CONCURRENCY = 60
DL_TIMEOUT = aiohttp.ClientTimeout(total=60)
DL_RETRIES = 2

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
}

# =========================
# UTIL
# =========================

def ensure_dirs():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(IMG_PRODUCT_DIR, exist_ok=True)
    os.makedirs(IMG_DRAWING_DIR, exist_ok=True)


def choose_excel_file():
    root = Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="Chọn file Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )
    root.destroy()
    return path


def detect_column(df, keywords):
    for col in df.columns:
        cl = str(col).lower()
        for k in keywords:
            if k.lower() in cl:
                return col
    return None


def safe_text(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


def safe_part_id(s: str) -> str:
    s = safe_text(s)
    s = re.sub(r"[^\w\-.]+", "_", s, flags=re.UNICODE)
    return s[:150]


def is_valid_http_url(u: str) -> bool:
    u = safe_text(u)
    return u.startswith("http://") or u.startswith("https://")


def ext_from_url(u: str) -> str:
    try:
        path = urlparse(u).path
        _, ext = os.path.splitext(path)
        ext = ext.lower().strip()

        if ext in {
            ".jpg",
            ".jpeg",
            ".png",
            ".webp",
            ".gif",
            ".svg",   # 👈 BỔ SUNG SVG
        }:
            return ext
    except Exception:
        pass

    # fallback an toàn
    return ".png"



def merge_missing(base: dict, fill: dict) -> dict:
    out = dict(base)
    for k, v in fill.items():
        if k not in out or out[k] in ("", None):
            out[k] = v
    return out


def export_report_if_exists():
    if os.path.exists(TMP_CSV):
        print("\n🧾 Đang xuất report hiện tại...")
        pd.read_csv(TMP_CSV).to_excel(REPORT_XLSX, index=False)
        print(f"✅ Đã xuất: {REPORT_XLSX}")
    else:
        print("⚠️ Không có TMP_CSV để xuất")

# =========================
# PROGRESS
# =========================

def progress_line(p: dict):
    done, total = p["done"], p["total"]
    elapsed = max(1e-6, time.monotonic() - p["start"])
    pct = done / total * 100
    spm = done / elapsed * 60
    eta = (total - done) / (done / elapsed) if done else 0
    print(
        f"\r🚀 {done}/{total} | {pct:5.1f}% | {spm:7.1f} url/phút | ETA {eta/60:6.1f} phút",
        end="",
        flush=True,
    )

# =========================
# STAGE 1 - HTTP PARSER
# =========================

def xpath_first(tree, xp: str) -> str:
    v = tree.xpath(xp)
    if not v:
        return ""
    if isinstance(v[0], str):
        return safe_text(v[0])
    return safe_text(v[0].text_content())


def parse_html_lxml(doc: str, part_id: str, url: str) -> dict:
    data = {
        "part_id": part_id,
        "url_product": url,
        "product_name": "",
        "sku": "",
        "stock": "",
        "price": "",
        "img_product_url": "",
        "img_drawing_url": "",
    }

    tree = LH.fromstring(doc)

    cats = tree.xpath('//nav[contains(@class,"woocommerce-breadcrumb")]//a')
    for i, a in enumerate(cats, 1):
        data[f"cate_{i}"] = safe_text(a.text_content())

    data["product_name"] = xpath_first(tree, '//h1[contains(@class,"product_title")][1]')

    sku_text = xpath_first(tree, '//*[contains(@class,"product-sku-display")][1]')
    if sku_text:
        data["sku"] = re.sub(r"^SKU:\s*", "", sku_text, flags=re.I)
    else:
        data["sku"] = xpath_first(tree, '//*[contains(@class,"sku")][1]')

    data["stock"] = xpath_first(tree, '(//p[contains(@class,"stock")])[1]')

    price = tree.xpath('(//p[contains(@class,"price")]//bdi)[1]')
    if price:
        data["price"] = safe_text(price[0].text_content())
    else:
        data["price"] = xpath_first(tree, '(//bdi)[1]')

    zoom = tree.xpath('(//img[contains(@class,"zoomImg")]/@src)[1]')
    if zoom:
        data["img_product_url"] = zoom[0]
    else:
        og = tree.xpath('(//meta[@property="og:image"]/@content)[1]')
        if og:
            data["img_product_url"] = og[0]

    imgs = tree.xpath('//img/@data-lazy-src | //img/@data-src | //img/@src')
    for s in imgs:
        if re.search(r"banve|ban-ve|ban_ve", s or "", re.I):
            data["img_drawing_url"] = s
            break

    return data


def is_complete_http(d: dict) -> bool:
    must = [d.get("product_name"), d.get("sku"), d.get("price"), d.get("img_product_url")]
    cats = [v for k, v in d.items() if k.startswith("cate_") and v]
    return all(safe_text(x) for x in must) and len(cats) >= 2


async def http_fetch_one(sem, session, part_id, url, prog):
    async with sem:
        out = {"part_id": part_id, "url_product": url}
        try:
            async with session.get(url, headers=HEADERS) as r:
                r.raise_for_status()
                doc = await r.text()
            out.update(parse_html_lxml(doc, part_id, url))
        except Exception as e:
            out["error"] = str(e)
        prog["done"] += 1
        progress_line(prog)
        return out

# =========================
# STAGE 1 - PLAYWRIGHT
# =========================

async def pw_extract_one(page, part_id, url) -> dict:
    data = {"part_id": part_id, "url_product": url}
    try:
        await page.goto(url, wait_until="commit", timeout=PW_TIMEOUT_MS)
        result = await page.evaluate(
            """
            () => {
                const out = {categories: []};
                document.querySelectorAll("nav.woocommerce-breadcrumb a").forEach(a => out.categories.push(a.innerText||""));
                out.product_name = document.querySelector("h1.product_title")?.innerText || "";
                out.sku = document.querySelector(".product-sku-display")?.innerText || document.querySelector(".sku")?.innerText || "";
                out.stock = document.querySelector("p.stock")?.innerText || "";
                out.price = document.querySelector("p.price bdi")?.innerText || document.querySelector("bdi")?.innerText || "";
                out.img_product_url = document.querySelector("img.zoomImg")?.src || "";
                for (const im of document.images) {
                    const s = im.dataset?.lazySrc || im.dataset?.src || im.src || "";
                    if (/banve|ban-ve|ban_ve/i.test(s)) { out.img_drawing_url = s; break; }
                }
                return out;
            }
            """
        )
        for i, c in enumerate(result.get("categories", []), 1):
            if c:
                data[f"cate_{i}"] = c.strip()
        for k in ("product_name", "sku", "stock", "price", "img_product_url", "img_drawing_url"):
            data[k] = safe_text(result.get(k, ""))
    except Exception as e:
        data["error"] = str(e)
    return data


async def pw_fallback_run(items, domain_hint: str):
    out = []
    prog = {"done": 0, "total": len(items), "start": time.monotonic()}

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()

        async def route_handler(route, request):
            if request.resource_type in ("image", "font", "media", "stylesheet"):
                await route.abort()
                return
            if request.resource_type == "script" and domain_hint and domain_hint not in request.url:
                await route.abort()
                return
            await route.continue_()

        await context.route("**/*", route_handler)

        sem = asyncio.Semaphore(PW_CONCURRENCY)

        async def run(item):
            part_id, url, base_http = item
            async with sem:
                page = await context.new_page()
                pw = await pw_extract_one(page, part_id, url)
                await page.close()
                prog["done"] += 1
                progress_line(prog)
                return merge_missing(base_http, pw)

        tasks = [run(i) for i in items]
        for coro in asyncio.as_completed(tasks):
            out.append(await coro)

        await context.close()
        await browser.close()

    print()
    return out

# =========================
# STAGE 1 - HYBRID
# =========================

async def stage1_hybrid(excel_path: str):
    df = pd.read_excel(excel_path)

    col_part = detect_column(df, ["part", "part_id", "code", "mã", "sku"])
    col_url = detect_column(df, ["url", "link"])
    if not col_part or not col_url:
        raise ValueError("Không detect được cột part/url")

    df = df.rename(columns={col_part: "part_id", col_url: "url"})
    df["part_id"] = df["part_id"].map(safe_part_id)
    df["url"] = df["url"].map(safe_text)
    df = df[df["url"].map(is_valid_http_url)].reset_index(drop=True)

    if os.path.exists(TMP_CSV):
        os.remove(TMP_CSV)

    base_domain = urlparse(df["url"].iloc[0]).netloc if len(df) else ""

    sem = asyncio.Semaphore(HTTP_CONCURRENCY)

    async with aiohttp.ClientSession(timeout=HTTP_TIMEOUT) as session:
        for i in range(0, len(df), HTTP_BATCH_SIZE):
            batch = df.iloc[i : i + HTTP_BATCH_SIZE]
            prog = {"done": 0, "total": len(batch), "start": time.monotonic()}

            http_tasks = [
                http_fetch_one(sem, session, r["part_id"], r["url"], prog)
                for _, r in batch.iterrows()
            ]
            http_results = await asyncio.gather(*http_tasks)
            print()

            ok, fallback = [], []
            for d in http_results:
                if d.get("error") or not is_complete_http(d):
                    fallback.append((d.get("part_id"), d.get("url_product"), d))
                else:
                    ok.append(d)

            if fallback:
                fixed = await pw_fallback_run(fallback, base_domain)
                merged = ok + fixed
            else:
                merged = ok

            pd.DataFrame(merged).to_csv(
                TMP_CSV,
                mode="a",
                header=not os.path.exists(TMP_CSV),
                index=False,
                encoding="utf-8-sig",
            )

    export_report_if_exists()

# =========================
# STAGE 2 - DOWNLOAD IMAGES
# =========================

async def download_one(sem, session, url, path):
    async with sem:
        for i in range(DL_RETRIES + 1):
            try:
                async with session.get(url, headers=HEADERS) as r:
                    r.raise_for_status()
                    with open(path, "wb") as f:
                        async for c in r.content.iter_chunked(65536):
                            f.write(c)
                return ""
            except Exception as e:
                if i >= DL_RETRIES:
                    return str(e)
                await asyncio.sleep(0.6 * (i + 1))


async def stage2_download_images(report_path: str):
    df = pd.read_excel(report_path)
    sem = asyncio.Semaphore(DL_CONCURRENCY)

    tasks = []
    for _, r in df.iterrows():
        part_id = safe_part_id(r.get("part_id", ""))
        if not part_id:
            continue

        img1 = safe_text(r.get("img_product_url", ""))
        img2 = safe_text(r.get("img_drawing_url", ""))

        if img1:
            path = os.path.join(IMG_PRODUCT_DIR, f"{part_id}{ext_from_url(img1)}")
            if not os.path.exists(path):
                tasks.append((part_id, img1, path))

        if img2:
            path = os.path.join(IMG_DRAWING_DIR, f"{part_id}{ext_from_url(img2)}")
            if not os.path.exists(path):
                tasks.append((part_id, img2, path))

    async with aiohttp.ClientSession(timeout=DL_TIMEOUT) as session:
        runners = [download_one(sem, session, u, p) for _, u, p in tasks]
        await asyncio.gather(*runners)

# =========================
# PIPELINE MENU
# =========================

def choose_pipeline_mode() -> str:
    print("\n================ PIPELINE MODE ================")
    print("1 - Chỉ crawl dữ liệu (Stage 1)")
    print("2 - Chỉ tải ảnh từ report có sẵn (Stage 2)")
    print("3 - Crawl dữ liệu + tải ảnh")
    print("0 - Thoát")
    print("================================================")

    while True:
        c = input("Chọn chế độ (0 / 1 / 2 / 3): ").strip()
        if c in {"0", "1", "2", "3"}:
            return c
        print("❌ Lựa chọn không hợp lệ")

# =========================
# MAIN
# =========================

async def main():
    ensure_dirs()

    mode = choose_pipeline_mode()
    if mode == "0":
        print("👋 Thoát chương trình.")
        return

    # ==================================================
    # MODE 1 – CHỈ CRAWL (STAGE 1)
    # ==================================================
    if mode == "1":
        excel_path = choose_excel_file()
        if not excel_path:
            print("❌ Không chọn file Excel.")
            return
        try:
            await stage1_hybrid(excel_path)
        except KeyboardInterrupt:
            print("\n🛑 CTRL+C – xuất report hiện tại...")
            export_report_if_exists()
        except Exception as e:
            print(f"\n❌ Lỗi không mong muốn: {e}")
            export_report_if_exists()
            raise
        return

    # ==================================================
    # MODE 2 – CHỈ TẢI ẢNH (STAGE 2)
    # ==================================================
    if mode == "2":
        if not os.path.exists(REPORT_XLSX):
            print(f"❌ Không tìm thấy {REPORT_XLSX}")
            print("👉 Hãy crawl Stage 1 trước.")
            return
        await stage2_download_images(REPORT_XLSX)
        return

    # ==================================================
    # MODE 3 – CRAWL → HỎI → TẢI ẢNH (GIỮ CẤU TRÚC CŨ)
    # ==================================================
    if mode == "3":
        excel_path = choose_excel_file()
        if not excel_path:
            print("❌ Không chọn file Excel.")
            return

        try:
            await stage1_hybrid(excel_path)
        except KeyboardInterrupt:
            print("\n🛑 CTRL+C – xuất report hiện tại...")
            export_report_if_exists()
            return
        except Exception as e:
            print(f"\n❌ Lỗi không mong muốn: {e}")
            export_report_if_exists()
            raise

        # 👇 GIỮ ĐÚNG HÀNH VI CŨ: HỎI SAU KHI CRAWL
        ans = input("\n📥 Bạn có muốn tải ảnh ngay không? (y/n): ").strip().lower()
        if ans == "y":
            await stage2_download_images(REPORT_XLSX)
        else:
            print("ℹ️ Bỏ qua Stage 2. Bạn có thể tải ảnh sau.")



if __name__ == "__main__":
    asyncio.run(main())
