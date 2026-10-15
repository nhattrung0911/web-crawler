import asyncio
import json
import re
import pandas as pd
from tkinter import Tk, filedialog
from playwright.async_api import async_playwright


# =====================================================================
# CHỌN FILE EXCEL
# =====================================================================
def choose_excel_file():
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )


# =====================================================================
# LÀM SẠCH LABEL
# =====================================================================
def normalize_label(label: str) -> str:
    label = label.strip().replace("\n", " ")
    if label.endswith(":"):
        label = label[:-1].strip()
    return " ".join(label.split())


# =====================================================================
# MAP LABEL -> KEY CHUẨN HÓA CHO TECH (OPTION A)
# =====================================================================
def map_spec_label_to_key(label: str) -> str | None:
    label_lower = label.lower()

    mapping_patterns = [
        (["mô-men siết", "mô men siết"], "torque_nm"),
        (["tốc độ không tải"], "rpm"),
        (["điện thế pin", "điện thế (v)", "điện thế"], "voltage_v"),
        (["dung lượng"], "battery_capacity_ah"),
        (["trọng lượng kèm pin", "trọng lượng"], "weight_kg"),
        (["chiều dài"], "length_mm"),
        (["xuất xứ"], "origin"),
        (["đèn led"], "led"),
    ]

    for keywords, key in mapping_patterns:
        for kw in keywords:
            if kw in label_lower:
                return key

    return None


# =====================================================================
# TÌM SẢN PHẨM + XÁC ĐỊNH STATUS
# =====================================================================
async def search_product(page, part_code: str):
    """
    Trả về:
        url (str | None),
        status (str):
            - "done"
            - "no_product"
            - "<number>"
    """
    await page.goto("https://milwaukeevietnam.com.vn/", timeout=60000)

    # Nhập mã vào ô tìm kiếm
    await page.fill("form.searchform input[name='s']", part_code)
    await page.press("form.searchform input[name='s']", "Enter")

    await page.wait_for_timeout(1500)

    url_now = page.url

    # 1) TH1 — vào thẳng trang sản phẩm
    if "/san-pham/" in url_now:
        return url_now, "done"

    # 2) TH2 — hiện danh sách sản phẩm
    list_selector = "div.products div.product-grid-item a.product-image-link"

    try:
        await page.wait_for_selector(list_selector, timeout=3000)
    except:
        return None, "no_product"

    items = page.locator(list_selector)
    count = await items.count()

    if count == 0:
        return None, "no_product"

    # Lấy sản phẩm đầu tiên
    url_first = await items.first.get_attribute("href")
    return url_first, str(count)


# =====================================================================
# CRAWL BẢNG THÔNG SỐ KỸ THUẬT
# =====================================================================
async def extract_spec_table(page, result: dict):
    rows = page.locator("table tbody tr")

    try:
        count = await rows.count()
    except:
        count = 0

    qa_pairs = []
    other_specs = {}

    if count == 0:
        result.setdefault("spec_qa", "")
        result.setdefault("other_specs", "")
        return

    for i in range(count):
        try:
            tds = rows.nth(i).locator("td")
            if await tds.count() < 2:
                continue

            raw_label = await tds.nth(0).inner_text()
            raw_value = await tds.nth(1).inner_text()

            label = normalize_label(raw_label)
            value = raw_value.replace("\n", " ").strip()

            if not label and not value:
                continue

            # QA format
            qa_pairs.append(f"{label}: {value}")

            # Map TECH
            key = map_spec_label_to_key(label)
            if key:
                if key not in result or not result[key]:
                    result[key] = value
            else:
                other_specs[label] = value

        except:
            continue

    result["spec_qa"] = ", ".join(qa_pairs) if qa_pairs else ""

    if other_specs:
        result["other_specs"] = json.dumps(other_specs, ensure_ascii=False)
    else:
        result.setdefault("other_specs", "")


# =====================================================================
# CRAWL THÔNG SỐ DẠNG LIST <ul>
# =====================================================================
async def extract_spec_list(page, result: dict):
    try:
        items = await page.locator("#thongsokythuat_chitiet li").all_inner_texts()
    except:
        items = []

    if items:
        cleaned = [normalize_label(x) for x in items]
        result["spec_list"] = ", ".join(cleaned)
    else:
        result.setdefault("spec_list", "")


# =====================================================================
# CRAWL TRANG CHI TIẾT
# =====================================================================
async def extract_detail(page, url: str | None) -> dict:
    if not url:
        return {"status": "no_product"}

    await page.goto(url, timeout=60000)

    result = {"url": url}

    # NAME
    try:
        name = await page.inner_text("h1.product_title")
        result["name"] = name.strip()
    except:
        result["name"] = ""

    # PRICE
    try:
        price = await page.inner_text("p.price .woocommerce-Price-amount")
        result["price"] = price.strip()
    except:
        result["price"] = ""

    # IMAGE
    try:
        result["image"] = await page.get_attribute("img.zoomImg", "src") or ""
    except:
        result["image"] = ""

    # SKU
    try:
        sku = await page.inner_text(".product_meta .sku")
        result["sku"] = sku.strip()
    except:
        result["sku"] = ""

    # CATEGORY
    try:
        cats = await page.locator(".product_meta .posted_in a").all_inner_texts()
        result["category"] = ", ".join(cats)
    except:
        result["category"] = ""

    # TAGS
    try:
        tags = await page.locator(".product_meta .tagged_as a").all_inner_texts()
        result["tags"] = ", ".join(tags)
    except:
        result["tags"] = ""

    # SPEC TABLE
    await extract_spec_table(page, result)

    # SPEC LIST
    await extract_spec_list(page, result)

    return result


# =====================================================================
# WORKER
# =====================================================================
async def worker(browser, part: str, worker_id: int, results_list: list):
    context = await browser.new_context()
    page = await context.new_page()

    print(f"[Worker {worker_id}] Đang xử lý: {part}")

    try:
        url, status = await search_product(page, part)
        detail = await extract_detail(page, url)

        detail["part"] = part
        detail["status"] = status

        results_list.append(detail)

        print(f"[Worker {worker_id}] ✔ Xong: {part} ({status})")

    except Exception as e:
        results_list.append({
            "part": part,
            "status": f"error: {e}"
        })
        print(f"[Worker {worker_id}] ❌ Lỗi: {e}")

    finally:
        await context.close()


# =====================================================================
# RUN BATCH
# =====================================================================
async def run_batch(browser, batch: list, results):
    tasks = [
        worker(browser, part, i + 1, results)
        for i, part in enumerate(batch)
    ]
    await asyncio.gather(*tasks)


# =====================================================================
# LƯU EXCEL
# =====================================================================
def save_excel(results):
    fixed_cols = [
        "part",
        "status",
        "name",
        "price",
        "url",
        "image",
        "sku",
        "category",
        "tags",
        "spec_qa",
        "spec_list",
        "other_specs",
        "torque_nm",
        "rpm",
        "voltage_v",
        "battery_capacity_ah",
        "weight_kg",
        "length_mm",
        "origin",
        "led",
    ]

    all_keys = set()
    for r in results:
        all_keys.update(r.keys())

    dynamic_cols = sorted([k for k in all_keys if k not in fixed_cols])
    final_cols = fixed_cols + dynamic_cols

    df = pd.DataFrame(results)

    for col in final_cols:
        if col not in df.columns:
            df[col] = ""

    df = df[final_cols]
    df.to_excel("Milwaukee_Report_V3.xlsx", index=False)

    print("\n📁 Đã xuất file: Milwaukee_Report_V3.xlsx\n")


# =====================================================================
# MAIN
# =====================================================================
async def main():
    excel_file = choose_excel_file()
    if not excel_file:
        print("❌ Không chọn file Excel.")
        return

    df = pd.read_excel(excel_file)

    if "part" not in df.columns:
        print("❌ File Excel phải có cột 'part'")
        return

    parts = df["part"].astype(str).tolist()
    results = []

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False)

        for i in range(0, len(parts), 10):
            batch = parts[i:i + 10]
            print(f"\n=== Batch {i//10 + 1} ({len(batch)} mã) ===")
            await run_batch(browser, batch, results)

        await browser.close()

    save_excel(results)


if __name__ == "__main__":
    asyncio.run(main())
