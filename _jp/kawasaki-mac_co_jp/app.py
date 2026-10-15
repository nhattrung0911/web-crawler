import streamlit as st
import asyncio
import aiohttp
import pandas as pd
from bs4 import BeautifulSoup
import os
import time
import zipfile
import io
import re
from datetime import datetime

# --- CONFIGURATION & THEME ---
st.set_page_config(
    page_title="Kawasaki Tools Crawler Pro",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Enterprise Theme: Blue & Yellow
st.markdown("""
    <style>
    :root {
        --primary-color: #0055AA;
        --secondary-color: #FFCC00;
        --text-color: #333333;
    }
    .stApp {
        background-color: #F5F7FA;
    }
    .main-header {
        background-color: var(--primary-color);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 20px;
        border-bottom: 5px solid var(--secondary-color);
    }
    .stButton>button {
        background-color: var(--primary-color);
        color: white;
        border: none;
        border-radius: 5px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #004488;
        color: var(--secondary-color);
    }
    .stDownloadButton>button {
        background-color: var(--secondary-color);
        color: black;
        font-weight: bold;
    }
    h1, h2, h3 {
        color: var(--primary-color);
    }
    .success-box {
        padding: 10px;
        background-color: #d4edda;
        color: #155724;
        border-radius: 5px;
        border-left: 5px solid #28a745;
    }
    </style>
""", unsafe_allow_html=True)

# Global Constants
BASE_URL = "https://tools.kawasaki-mac.co.jp"
MASTER_DB_FILE = "master_url_db.csv"
RESULT_FILE = "kawasaki_product_data.csv"
ASSET_DIR = "downloaded_assets"
CONCURRENCY_LIMIT = 10  # Số lượng request đồng thời (Batch size)

# --- UTILS ---

def clean_text(text):
    if text:
        return text.strip().replace('\n', ' ').replace('\r', '')
    return ""

def normalize_part_number(pn):
    """Chuẩn hóa mã sản phẩm để mapping chính xác (bỏ dấu, space, uppercase)"""
    if pd.isna(pn) or pn == "":
        return ""
    # Bỏ dấu chấm, phẩy, gạch ngang, khoảng trắng, đưa về chữ thường
    return re.sub(r'[\s\-\.,_]', '', str(pn)).lower()

# --- ASYNC CRAWLER ENGINE ---

class KawasakiCrawler:
    def __init__(self):
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        self.semaphore = asyncio.Semaphore(CONCURRENCY_LIMIT)

    async def fetch(self, session, url):
        """Hàm fetch cơ bản với retry và timeout"""
        async with self.semaphore:
            try:
                async with session.get(url, headers=self.headers, timeout=30) as response:
                    if response.status == 200:
                        return await response.text()
                    else:
                        return None
            except Exception as e:
                return None

    async def fetch_bytes(self, session, url):
        """Hàm tải file binary (ảnh, pdf)"""
        async with self.semaphore:
            try:
                async with session.get(url, headers=self.headers, timeout=60) as response:
                    if response.status == 200:
                        return await response.read()
                    return None
            except Exception:
                return None

    # --- PIPELINE 1: DISCOVERY ---
    async def get_all_product_urls(self):
        async with aiohttp.ClientSession() as session:
            # 1. Get Categories
            start_url = f"{BASE_URL}/products/category/"
            html = await self.fetch(session, start_url)
            if not html:
                return []
            
            soup = BeautifulSoup(html, 'lxml')
            categories = []
            
            # Selector swiper-slide--category-products
            slides = soup.select('.swiper-slide--category-products a')
            for slide in slides:
                href = slide.get('href')
                name = slide.find('h3', class_='category-product-slider__info__title')
                cat_name = clean_text(name.text) if name else "Unknown"
                if href:
                    full_cat_url = href if href.startswith('http') else f"{BASE_URL}{href}"
                    categories.append({'cat_name': cat_name, 'url': full_cat_url})
            
            # 2. Visit each Category to get Product URLs
            all_products = []
            
            # Tạo task cho mỗi category
            tasks = [self.parse_category_page(session, cat) for cat in categories]
            results = await asyncio.gather(*tasks)
            
            for res in results:
                all_products.extend(res)
                
            return all_products

    async def parse_category_page(self, session, category):
        html = await self.fetch(session, category['url'])
        products = []
        if html:
            soup = BeautifulSoup(html, 'lxml')
            # Selector fixnav__unit
            units = soup.select('.fixnav__unit a.fixnav__unit__title')
            for unit in units:
                p_name = clean_text(unit.text)
                href = unit.get('href')
                if href:
                    # Xử lý link dạng anchor #sec_... -> cần link gốc của sản phẩm nếu có, 
                    # Nhưng theo đề bài "truy cập vào mỗi category lấy all product url"
                    # Link trong fixnav thường trỏ đến section trong trang category hoặc trang detail.
                    # Kiểm tra xem href có phải là query param ?id=... hay không
                    full_url = href if href.startswith('http') else f"{BASE_URL}{href}"
                    
                    # Chuẩn hóa để lưu
                    products.append({
                        'Category': category['cat_name'],
                        'Product Code': p_name,
                        'Normalized Code': normalize_part_number(p_name),
                        'URL': full_url
                    })
        return products

    # --- PIPELINE 2: DETAIL EXTRACTION ---
    async def crawl_details(self, product_list, status_placeholder):
        results = []
        async with aiohttp.ClientSession() as session:
            total = len(product_list)
            completed = 0
            
            # Chia nhỏ thành các batch để tránh overload và dễ save
            batch_size = 10
            for i in range(0, total, batch_size):
                batch = product_list[i:i+batch_size]
                tasks = [self.parse_product_detail(session, item) for item in batch]
                batch_results = await asyncio.gather(*tasks)
                
                valid_results = [r for r in batch_results if r is not None]
                results.extend(valid_results)
                
                # Auto-save logic (Append mode)
                df_batch = pd.DataFrame(valid_results)
                if not os.path.isfile(RESULT_FILE):
                    df_batch.to_csv(RESULT_FILE, index=False, encoding='utf-8-sig')
                else:
                    df_batch.to_csv(RESULT_FILE, mode='a', header=False, index=False, encoding='utf-8-sig')

                completed += len(batch)
                status_placeholder.progress(min(completed / total, 1.0), text=f"Đang xử lý: {completed}/{total} sản phẩm...")
                
        return results

    async def parse_product_detail(self, session, item):
        html = await self.fetch(session, item['URL'])
        if not html:
            return None
        
        soup = BeautifulSoup(html, 'lxml')
        
        # 1. Basic Info
        # Tìm section ID tương ứng nếu URL có anchor (#sec_123)
        # Tuy nhiên, aiohttp lấy toàn bộ source, ta cần tìm đúng block sản phẩm nếu trang category list hết sp.
        # Logic: Nếu URL dẫn về trang Category có parameter ?id=77#sec_1181
        # Thì ta phải tìm element có id="sec_1181" HOẶC crawl theo cấu trúc detail page riêng lẻ.
        # Dựa vào đề bài, ta giả định mỗi URL dẫn đến nội dung cụ thể.
        
        # Để chính xác nhất với Kawasaki tools (thường là Single Page cho Category chứa nhiều Product):
        # Ta cần parse HTML và tìm đúng khối dữ liệu dựa trên Product Code hoặc ID.
        # Nhưng để đơn giản và hiệu quả, ta sẽ tìm khối chứa title khớp với Product Code.
        
        # Thử tìm theo cấu trúc mẫu user cung cấp
        # Tìm thẻ h1 chứa tên sản phẩm
        
        # Logic tìm kiếm scope:
        # Nếu URL có #sec_..., dùng BeautifulSoup tìm ID đó.
        target_id = None
        if '#' in item['URL']:
            target_id = item['URL'].split('#')[-1]
        
        container = soup
        if target_id:
            found_container = soup.find(id=target_id)
            if found_container:
                container = found_container
            else:
                # Fallback: Tìm h1 text khớp product code
                titles = soup.find_all('h1', class_='item__info__title')
                for t in titles:
                    if item['Product Code'] in t.text:
                        container = t.find_parent() # Lên cấp cha để lấy context
                        break

        # Extract Data
        # Type
        type_tag = container.find(class_='item__info__type')
        p_type = clean_text(type_tag.text) if type_tag else ""
        
        # Name (Code) - Lấy lại để confirm
        name_tag = container.find(class_='item__info__title')
        p_name = clean_text(name_tag.text) if name_tag else item['Product Code']
        
        # Specs (Table)
        specs_list = []
        rows = container.select('.item__info__table__row')
        for row in rows:
            head = row.find(class_='item__info__table__head')
            body = row.find(class_='item__info__table__body')
            if head and body:
                k = clean_text(head.text)
                v = clean_text(body.text)
                specs_list.append(f"{k}: {v}")
        spec_str = "\n".join(specs_list)
        
        # Images
        img_urls = []
        # Thumbnails nav
        imgs = container.select('.item__image__nav__unit img')
        for img in imgs:
            src = img.get('src')
            if src:
                # Xử lý để lấy ảnh gốc (nếu có logic resize trong url, thường WP thì bỏ kích thước cuối file)
                # VD: image-150x150.png -> image.png (Logic tùy chọn, ở đây lấy raw trước)
                full_src = src if src.startswith('http') else f"{BASE_URL}{src}"
                img_urls.append(full_src)
        img_str = " | ".join(img_urls)

        # PDF / Print URL
        pdf_tag = container.find('a', class_='item__info__print')
        pdf_url = ""
        if pdf_tag:
            href = pdf_tag.get('href')
            pdf_url = href if href.startswith('http') else f"{BASE_URL}{href}"

        return {
            'Category': item['Category'],
            'Product Code': p_name,
            'Product Type': p_type,
            'Specs': spec_str,
            'Image URLs': img_str,
            'PDF URL': pdf_url,
            'Product URL': item['URL']
        }

    # --- PIPELINE 3: ASSET DOWNLOADER ---
    async def download_assets(self, data_df, status_placeholder):
        if not os.path.exists(ASSET_DIR):
            os.makedirs(ASSET_DIR)
            
        async with aiohttp.ClientSession() as session:
            total_items = len(data_df)
            processed = 0
            
            for index, row in data_df.iterrows():
                p_code = row['Product Code']
                # Tạo folder cho sản phẩm (sanitize tên folder)
                safe_name = "".join([c for c in p_code if c.isalnum() or c in (' ', '-', '_')]).strip()
                item_dir = os.path.join(ASSET_DIR, safe_name)
                if not os.path.exists(item_dir):
                    os.makedirs(item_dir)
                
                tasks = []
                
                # 1. Download Images
                img_urls = str(row['Image URLs']).split(' | ')
                for i, url in enumerate(img_urls):
                    if url and url.strip():
                        ext = url.split('.')[-1]
                        filename = f"{safe_name}_img_{i+1}.{ext}"
                        filepath = os.path.join(item_dir, filename)
                        tasks.append(self._save_file(session, url, filepath))
                
                # 2. Download PDF
                pdf_url = row['PDF URL']
                if pdf_url and str(pdf_url) != 'nan':
                    filename = f"{safe_name}_spec.pdf" # Link thường là print page, có thể save HTML hoặc PDF nếu link là .pdf
                    # Nếu link là /products/print/?id=... thì nó là trang HTML view print.
                    # Crawler này sẽ tải nội dung về. Nếu cần PDF thật sự phải có link .pdf
                    filepath = os.path.join(item_dir, filename)
                    tasks.append(self._save_file(session, pdf_url, filepath))
                
                if tasks:
                    await asyncio.gather(*tasks)
                
                processed += 1
                status_placeholder.progress(min(processed / total_items, 1.0), 
                                          text=f"Đang tải Assets: {p_code} ({processed}/{total_items})")

    async def _save_file(self, session, url, filepath):
        content = await self.fetch_bytes(session, url)
        if content:
            with open(filepath, 'wb') as f:
                f.write(content)

# --- STREAMLIT UI LOGIC ---

def run_pipeline_1():
    st.subheader("🛠 Pipeline 1: Khởi tạo & Quét danh mục")
    st.info("Bước này sẽ quét toàn bộ website để lấy danh sách mã sản phẩm và lưu vào Master DB.")
    
    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("🚀 Bắt đầu Quét Master Data"):
            crawler = KawasakiCrawler()
            with st.status("Đang khởi tạo crawler...", expanded=True) as status:
                st.write("Đang truy cập trang danh mục...")
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                all_products = loop.run_until_complete(crawler.get_all_product_urls())
                
                if all_products:
                    df = pd.DataFrame(all_products)
                    df.to_csv(MASTER_DB_FILE, index=False, encoding='utf-8-sig')
                    status.update(label="✅ Quét hoàn tất!", state="complete", expanded=False)
                    st.success(f"Đã tìm thấy {len(df)} sản phẩm. Data đã lưu vào '{MASTER_DB_FILE}'")
                else:
                    status.update(label="❌ Lỗi kết nối", state="error")

    if os.path.exists(MASTER_DB_FILE):
        df_master = pd.read_csv(MASTER_DB_FILE)
        st.dataframe(df_master, use_container_width=True, height=300)
    else:
        st.warning("Chưa có dữ liệu Master. Vui lòng chạy quét trước.")

def run_pipeline_2():
    st.subheader("📊 Pipeline 2: Chi tiết Sản phẩm & Specs")
    
    if not os.path.exists(MASTER_DB_FILE):
        st.error("⚠️ Vui lòng chạy Pipeline 1 trước để có danh sách sản phẩm.")
        return

    df_master = pd.read_csv(MASTER_DB_FILE)
    
    # Selection Mode
    option = st.radio("Chọn chế độ Crawler:", ["Quét Toàn Bộ (All Data)", "Quét theo danh sách Part Number (Input)"])
    
    target_products = []
    
    if option == "Quét theo danh sách Part Number (Input)":
        user_input = st.text_area("Nhập danh sách mã sản phẩm (cách nhau bởi dấu phẩy, khoảng trắng hoặc xuống dòng):", height=150)
        if user_input:
            # Normalize Input
            raw_parts = re.split(r'[,\s\n]+', user_input)
            input_parts_norm = [normalize_part_number(p) for p in raw_parts if p]
            
            # Map input to Master DB
            df_master['Normalized Code'] = df_master['Product Code'].apply(normalize_part_number)
            
            # Filter
            matched_df = df_master[df_master['Normalized Code'].isin(input_parts_norm)]
            target_products = matched_df.to_dict('records')
            
            st.write(f"🔍 Đã tìm thấy {len(target_products)} / {len(input_parts_norm)} mã sản phẩm hợp lệ trong DB.")
            if len(target_products) < len(input_parts_norm):
                st.warning("Một số mã không tìm thấy, vui lòng kiểm tra lại chính tả.")
    else:
        target_products = df_master.to_dict('records')
        st.write(f"📂 Sẵn sàng quét toàn bộ {len(target_products)} sản phẩm.")

    # Crawler Action
    if st.button("▶️ Bắt đầu lấy chi tiết (Resume Supported)", type="primary"):
        if not target_products:
            st.error("Không có sản phẩm nào để xử lý.")
        else:
            # Resume Logic: Check existing results
            existing_codes = []
            if os.path.exists(RESULT_FILE):
                try:
                    df_exist = pd.read_csv(RESULT_FILE)
                    existing_codes = df_exist['Product Code'].astype(str).tolist()
                    st.info(f"Phát hiện file kết quả cũ. Đã có {len(existing_codes)} sản phẩm. Đang bỏ qua các mã đã có...")
                except:
                    pass
            
            # Filter out existing
            final_queue = [p for p in target_products if str(p['Product Code']) not in existing_codes]
            
            if not final_queue:
                st.success("Tất cả sản phẩm đã được quét xong trước đó!")
            else:
                progress_bar = st.progress(0, text="Chuẩn bị...")
                crawler = KawasakiCrawler()
                
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                loop.run_until_complete(crawler.crawl_details(final_queue, progress_bar))
                
                st.success("✅ Hoàn tất quá trình lấy dữ liệu chi tiết!")
    
    # Display Result
    if os.path.exists(RESULT_FILE):
        st.divider()
        st.markdown("### 📥 Kết quả")
        df_result = pd.read_csv(RESULT_FILE)
        st.dataframe(df_result, use_container_width=True)
        
        # Download Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_result.to_excel(writer, index=False, sheet_name='Sheet1')
        
        st.download_button(
            label="Tải xuống file Excel (.xlsx)",
            data=buffer.getvalue(),
            file_name="kawasaki_tools_data.xlsx",
            mime="application/vnd.ms-excel"
        )

def run_pipeline_3():
    st.subheader("💾 Pipeline 3: Tải Assets (Images & PDFs)")
    st.info("Pipe này hoạt động độc lập. Bạn có thể tải Assets dựa trên file kết quả từ Pipe 2.")
    
    if not os.path.exists(RESULT_FILE):
        uploaded_file = st.file_uploader("Hoặc tải lên file CSV kết quả (nếu chưa chạy Pipe 2)", type=['csv'])
        if uploaded_file:
            df = pd.read_csv(uploaded_file)
        else:
            st.warning("Vui lòng chạy Pipe 2 hoặc upload file CSV.")
            return
    else:
        df = pd.read_csv(RESULT_FILE)
    
    st.write(f"Đang có dữ liệu của {len(df)} sản phẩm.")
    
    # Filter selection for pipe 3
    option = st.radio("Phạm vi tải:", ["Tải toàn bộ", "Nhập mã cụ thể để tải"])
    target_df = df
    
    if option == "Nhập mã cụ thể để tải":
        user_input = st.text_area("Nhập mã (Pipe 3):", height=100)
        if user_input:
            raw_parts = re.split(r'[,\s\n]+', user_input)
            # Simple contains filter
            target_df = df[df['Product Code'].astype(str).apply(lambda x: any(p.lower() in x.lower() for p in raw_parts if p))]
            st.write(f"Lọc được {len(target_df)} sản phẩm.")

    if st.button("📥 Bắt đầu Tải File & Nén ZIP"):
        if target_df.empty:
            st.error("Không có sản phẩm để tải.")
        else:
            progress_bar = st.progress(0, text="Đang khởi tạo...")
            crawler = KawasakiCrawler()
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            loop.run_until_complete(crawler.download_assets(target_df, progress_bar))
            
            # Zip logic
            st.write("🗜️ Đang nén file ZIP...")
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(ASSET_DIR):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, ASSET_DIR)
                        zf.write(file_path, arcname)
            
            st.success("✅ Đã tạo file ZIP thành công!")
            st.download_button(
                label="📦 Tải xuống All Assets (.zip)",
                data=zip_buffer.getvalue(),
                file_name="kawasaki_assets.zip",
                mime="application/zip"
            )

# --- MAIN APP LAYOUT ---

def main():
    st.markdown('<div class="main-header"><h1>🏭 KAWASAKI TOOLS AUTOMATION PIPELINE</h1></div>', unsafe_allow_html=True)
    
    tabs = st.tabs(["⚙️ Pipe 1: Master Data", "📊 Pipe 2: Detail Crawler", "💾 Pipe 3: Asset Downloader"])
    
    with tabs[0]:
        run_pipeline_1()
    
    with tabs[1]:
        run_pipeline_2()
        
    with tabs[2]:
        run_pipeline_3()

if __name__ == "__main__":
    main()