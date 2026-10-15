import streamlit as st
import asyncio
import aiohttp
import pandas as pd
from bs4 import BeautifulSoup
import os
import re
import io
import sys
import time

# --- CẤU HÌNH TỐI ƯU ---
st.set_page_config(
    page_title="Kawasaki Tools Pro Crawler",
    page_icon="⚡",
    layout="wide"
)

# Cấu hình Crawler
CONCURRENCY_LIMIT = 50  # Max Speed
MAX_IMAGES_COLUMNS = 10 # Số cột ảnh

# File paths
MASTER_DB_FILE = "master_url_db.csv"
RESULT_FILE = "kawasaki_product_details.csv"

# Fix lỗi asyncio trên Windows
if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# Style
st.markdown("""
    <style>
    :root { --primary: #0055AA; --secondary: #FFCC00; }
    .stApp { background-color: #F0F2F6; }
    .header-box { background: var(--primary); padding: 15px; border-radius: 8px; color: white; border-bottom: 4px solid var(--secondary); margin-bottom: 15px; }
    .stButton>button { background-color: var(--primary); color: white; border-radius: 6px; font-weight: bold; width: 100%; }
    .stButton>button:hover { background-color: #004488; color: var(--secondary); }
    .success-msg { color: #155724; background-color: #d4edda; padding: 10px; border-radius: 5px; margin: 10px 0; }
    </style>
""", unsafe_allow_html=True)

# --- UTILS ---

def clean_text(text):
    if text: return re.sub(r'\s+', ' ', text).strip()
    return ""

def normalize_code(code):
    """Chuẩn hóa mã: KPT-123 -> kpt123 để so sánh tuyệt đối"""
    if not code: return ""
    return re.sub(r'[\s\-\.,_]', '', str(code)).lower()

# --- ENGINE ---

class FastCrawler:
    def __init__(self):
        # Không tạo connector ở đây
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        self.semaphore = asyncio.Semaphore(CONCURRENCY_LIMIT)

    async def fetch(self, session, url):
        async with self.semaphore:
            try:
                async with session.get(url, headers=self.headers, timeout=15) as response:
                    if response.status == 200:
                        return await response.text()
            except: pass
            return None

    # --- PIPE 1: QUÉT LINK & KHỬ TRÙNG (LOGIC MỚI) ---
    async def get_all_products(self):
        connector = aiohttp.TCPConnector(limit=100, ttl_dns_cache=300)
        async with aiohttp.ClientSession(connector=connector) as session:
            # 1. Lấy Categories
            url = "https://tools.kawasaki-mac.co.jp/products/category/"
            html = await self.fetch(session, url)
            if not html: return []

            soup = BeautifulSoup(html, 'lxml')
            
            # Dùng Dictionary để khử trùng URL Category ngay lập tức
            # Key: URL -> Value: Name
            unique_cats = {} 
            
            # Select rộng hơn rồi lọc bằng code
            slides = soup.select('a[href*="/products/category/?id="]')
            for slide in slides:
                href = slide.get('href')
                # Tìm tên category gần nhất
                name_tag = slide.find('h3')
                if not name_tag:
                    name_tag = slide.find_parent().find('h3') if slide.find_parent() else None
                
                cat_name = clean_text(name_tag.text) if name_tag else "General"
                
                if href:
                    full_url = href if href.startswith('http') else f"https://tools.kawasaki-mac.co.jp{href}"
                    # Chỉ thêm nếu URL chưa có trong dict
                    if full_url not in unique_cats:
                        unique_cats[full_url] = cat_name

            # Chuyển về list dict
            cat_list = [{'url': k, 'name': v} for k, v in unique_cats.items()]
            
            # 2. Quét sản phẩm song song
            tasks = [self.parse_category(session, c) for c in cat_list]
            results = await asyncio.gather(*tasks)
            
            # GỘP VÀ KHỬ TRÙNG TUYỆT ĐỐI THEO PRODUCT CODE
            # Dùng Dictionary: Key = Normalized Code -> Value = Product Data
            # Logic này đảm bảo 1 mã KPT-123 chỉ xuất hiện 1 lần duy nhất dù HTML có lặp 100 lần
            unique_products_map = {}
            
            for group in results:
                for p in group:
                    norm_code = p['Normalized']
                    if norm_code and norm_code not in unique_products_map:
                        unique_products_map[norm_code] = p
            
            return list(unique_products_map.values())

    async def parse_category(self, session, cat):
        html = await self.fetch(session, cat['url'])
        products = []
        if html:
            soup = BeautifulSoup(html, 'lxml')
            units = soup.select('.fixnav__unit a.fixnav__unit__title')
            for unit in units:
                name = clean_text(unit.text)
                href = unit.get('href')
                if href and name:
                    full_url = href if href.startswith('http') else f"https://tools.kawasaki-mac.co.jp{href}"
                    products.append({
                        'Category': cat['name'],
                        'Product Code': name,
                        'Normalized': normalize_code(name),
                        'URL': full_url
                    })
        return products

    # --- PIPE 2: CRAWL DETAIL ---
    async def crawl_details_batch(self, product_list, progress_bar, status_text):
        connector = aiohttp.TCPConnector(limit=100, ttl_dns_cache=300)
        async with aiohttp.ClientSession(connector=connector) as session:
            total = len(product_list)
            chunk_size = 20 
            
            for i in range(0, total, chunk_size):
                batch = product_list[i : i + chunk_size]
                tasks = [self.parse_detail(session, item) for item in batch]
                results = await asyncio.gather(*tasks)
                
                valid_data = [r for r in results if r is not None]
                
                if valid_data:
                    df_chunk = pd.DataFrame(valid_data)
                    
                    # Tách cột ảnh
                    for idx in range(1, MAX_IMAGES_COLUMNS + 1):
                        col = f"Image {idx}"
                        if col not in df_chunk.columns: df_chunk[col] = ""
                    
                    # Reorder
                    cols = ['Category', 'Product Code', 'Product Type', 'Specs', 'PDF URL', 'Product URL'] + \
                           [f"Image {k}" for k in range(1, MAX_IMAGES_COLUMNS + 1)]
                    df_chunk = df_chunk.reindex(columns=cols).fillna("")

                    # Check Header để không ghi lại nhiều lần
                    hdr = not os.path.isfile(RESULT_FILE)
                    df_chunk.to_csv(RESULT_FILE, mode='a', header=hdr, index=False, encoding='utf-8-sig')

                # Update UI
                progress = min((i + chunk_size) / total, 1.0)
                progress_bar.progress(progress)
                status_text.text(f"🚀 Processing: {min(i + chunk_size, total)}/{total}")

    async def parse_detail(self, session, item):
        html = await self.fetch(session, item['URL'])
        if not html: return None
        soup = BeautifulSoup(html, 'lxml')
        
        # Tìm container chứa
        target_id = item['URL'].split('#')[-1] if '#' in item['URL'] else None
        container = soup.find(id=target_id) if target_id else soup
        
        if not container or container == soup:
             # Fallback tìm theo text title
             titles = soup.find_all('h1', class_='item__info__title')
             found = False
             for t in titles:
                 # So sánh tương đối để tìm đúng block
                 if normalize_code(item['Product Code']) in normalize_code(t.text):
                     container = t.parent
                     found = True
                     break
             if not found: return None 

        type_tag = container.find(class_='item__info__type')
        p_type = clean_text(type_tag.text) if type_tag else ""

        specs = []
        rows = container.select('.item__info__table__row')
        for row in rows:
            h = row.find(class_='item__info__table__head')
            b = row.find(class_='item__info__table__body')
            if h and b:
                specs.append(f"{clean_text(h.text)}: {clean_text(b.text)}")
        spec_str = " | ".join(specs)

        img_dict = {}
        imgs = container.select('.item__image__nav__unit img')
        for idx, img in enumerate(imgs):
            if idx >= MAX_IMAGES_COLUMNS: break
            src = img.get('src')
            if src:
                full = src if src.startswith('http') else f"https://tools.kawasaki-mac.co.jp{src}"
                img_dict[f"Image {idx+1}"] = full

        pdf_tag = container.find('a', class_='item__info__print')
        pdf_url = ""
        if pdf_tag:
            href = pdf_tag.get('href')
            pdf_url = href if href.startswith('http') else f"https://tools.kawasaki-mac.co.jp{href}"

        return {
            'Category': item['Category'],
            'Product Code': item['Product Code'],
            'Product Type': p_type,
            'Specs': spec_str,
            'PDF URL': pdf_url,
            'Product URL': item['URL'],
            **img_dict
        }

# --- MAIN ---

def main():
    st.markdown('<div class="header-box"><h2>🚀 KAWASAKI PRO: NO DUPLICATE VERSION</h2></div>', unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["1. Quét Master (Strict Dedup)", "2. Crawl Detail"])

    with tab1:
        col1, col2 = st.columns([1, 2])
        with col1:
            if st.button("START PIPE 1 (Clean Scan)"):
                with st.status("Đang quét & Khử trùng...", expanded=True):
                    # Xóa file cũ để tránh nối đuôi
                    if os.path.exists(MASTER_DB_FILE): os.remove(MASTER_DB_FILE)
                    
                    crawler = FastCrawler()
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    data = loop.run_until_complete(crawler.get_all_products())
                    loop.close()
                    
                    if data:
                        df = pd.DataFrame(data)
                        # KHỬ TRÙNG LẦN CUỐI BẰNG PANDAS CHO CHẮC CHẮN
                        df.drop_duplicates(subset=['Normalized'], inplace=True)
                        df.to_csv(MASTER_DB_FILE, index=False, encoding='utf-8-sig')
                        st.success(f"✅ Đã lưu {len(df)} sản phẩm DUY NHẤT vào Master DB.")
                    else:
                        st.error("Lỗi: Không lấy được dữ liệu.")

        if os.path.exists(MASTER_DB_FILE):
            df = pd.read_csv(MASTER_DB_FILE)
            st.write(f"📂 Master DB: **{len(df)}** Unique Products")
            st.dataframe(df, height=250)

    with tab2:
        if not os.path.exists(MASTER_DB_FILE):
            st.warning("Vui lòng chạy Pipe 1 trước.")
        else:
            df_master = pd.read_csv(MASTER_DB_FILE)
            
            # Kiểm tra file kết quả để resume
            done_codes = []
            if os.path.exists(RESULT_FILE):
                try:
                    df_done = pd.read_csv(RESULT_FILE)
                    if 'Product Code' in df_done.columns:
                        # Chuẩn hóa code đã làm để so sánh
                        done_codes = df_done['Product Code'].astype(str).apply(normalize_code).tolist()
                except: pass
            
            # Lọc danh sách cần chạy
            # Logic: Chỉ lấy những mã có trong Master mà CHƯA có trong Result
            df_master['Norm_Check'] = df_master['Product Code'].astype(str).apply(normalize_code)
            queue_df = df_master[~df_master['Norm_Check'].isin(done_codes)]
            queue = queue_df.to_dict('records')

            st.write(f"Tổng: {len(df_master)} | Đã xong: {len(done_codes)} | **Cần chạy: {len(queue)}**")
            
            if st.button("🚀 RUN CRAWLER", type="primary"):
                if not queue:
                    st.success("Tất cả dữ liệu đã được crawl!")
                else:
                    bar = st.progress(0)
                    txt = st.empty()
                    crawler = FastCrawler()
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    try:
                        loop.run_until_complete(crawler.crawl_details_batch(queue, bar, txt))
                        st.success("✅ Hoàn tất!")
                        time.sleep(1)
                        st.rerun()
                    finally:
                        loop.close()
            
            # Download & View
            if os.path.exists(RESULT_FILE):
                st.divider()
                st.subheader("Kết quả")
                df_res = pd.read_csv(RESULT_FILE)
                # Hiển thị khử trùng (đề phòng file cũ bẩn)
                df_view = df_res.drop_duplicates(subset=['Product Code'])
                st.write(f"Hiển thị {len(df_view)} dòng (Đã lọc trùng hiển thị)")
                st.dataframe(df_view, height=300)
                
                buf = io.BytesIO()
                with pd.ExcelWriter(buf) as writer:
                    df_view.to_excel(writer, index=False)
                st.download_button("Download Excel Clean", buf.getvalue(), "kawasaki_clean.xlsx")

if __name__ == "__main__":
    main()