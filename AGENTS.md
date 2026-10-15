# AGENTS.md — WORX GREEN crawler workshop

> Đọc file này TRƯỚC KHI build bất kỳ crawler mới nào trong thư mục này.
> Mục tiêu: mọi agent (Claude / Codex / Gemini) làm ra crawler **đúng pattern đã
> kiểm chứng**, chất lượng đồng đều, an toàn máy, test xong mới báo done.
>
> Recipe này chắt ra từ các crawler thật trong folder. Ví dụ nhắc tên folder =
> mở file đó đọc code thật.

---

## BẮT ĐẦU NHANH — agent mới đọc mục này trước (cập nhật 2026-10-15)

**Sơ đồ tech stack + cách dùng (1 trang, dễ nhìn):** https://web-crawler-diagram.pages.dev

**Môi trường:** Windows, Python 3.13. Mọi lệnh chạy từ thư mục `WORX GREEN/`.
```bash
pip install httpx selectolax openpyxl pandas tenacity pytest   # + playwright khi BUỘC phải dùng browser
```

**Làm 1 site mới = 6 bước, không bỏ bước nào:**
1. **Tra registry:** `python sites.py https://<site>/` → đã có thì sửa folder cũ, đừng làm trùng.
2. **Recon (§3):** curl/httpx probe site thật, săn API TRƯỚC. Đo thời gian 1 request + thử song song (§5.1).
3. **Clone archetype SẠCH gần nhất** (bảng §1.2, cột "doctor" = SACH). KHÔNG clone crawler còn `[!]`.
4. **Code** vào `_<vùng>/<domain_underscore>/<name>.py` + `<name>.bat` (§1).
   - Output/log/checkpoint ghi vào `out_<name>/` TRONG folder site (đã gitignore).
   - Test pytest: `tests/<name>/test_<name>.py` + `tests/<name>/fixtures/*.html` = HTML THẬT tải từ site.
5. **Kiểm chứng:**
   - Chạy live bằng mã/URL thật (§9.5).
   - Chạy test: `python -B -m pytest tests/<name> -q -p no:cacheprovider`.
   - Chạy doctor: `python -B sites.py doctor <name>` → phải ra `SACH`.
   - Dùng `-B` để không đẻ `__pycache__` (doctor coi là rác).
6. **Registry:** sửa `sites.json` (status + note recon thật) → `python sites.py` sinh lại `SITES.md`.

**Bản đồ root:**

| File / thư mục | Vai trò |
|---|---|
| `AGENTS.md` | Luật (file này) |
| `sites.json` / `sites.py` / `SITES.md` | Registry + `doctor` (§1.1) |
| `gui_common.py` | Ô chọn Excel dùng chung cho GUI (§4.4) |
| `crawl_store.py`, `schema.sql`, `product_crawl_schema.json` | Schema/store SQLite dùng chung |
| `tests/<name>/` | Test pytest + fixture HTML thật theo site |
| `_vn/ _cn/ _jp/ _global/` | Crawler theo vùng |

**CẤM:**
- Xoá/ghi đè `out_*`, `_raw_*.jsonl`, `_*_index.json`, `products.db` của site KHÁC (data/checkpoint thật của user).
- `taskkill /IM python.exe` (giết luôn process của user) — chỉ kill PID mình tạo.
- Commit/push khi user chưa bảo.

---

## 0. GATE BẮT BUỘC — đọc trước, không được bỏ qua (ĐÂY LÀ MECHANICAL GATE, không tin trí nhớ)

**KHÔNG ĐƯỢC nói "xong"/"done" cho một site nếu chưa làm đủ 3 việc dưới đây VÀ dán
bằng chứng (output thật, không tóm tắt) vào câu trả lời cuối:**

1. **Test headless bằng mã/URL THẬT** của chính site đó (không phải mã bịa) —
   dán bảng kết quả OK/Not Found + lý do (§9.5, §11).
2. **`python sites.py doctor <ten-site>`** chạy từ thư mục `WORX GREEN/` → PHẢI
   ra `SACH, khong pham luat nao`. Còn báo `[!]` là CHƯA xong, sửa rồi chạy lại,
   không được báo done trước.
3. **Cập nhật registry**: sửa `sites.json` (status + note recon thật, không để
   trống/`done?`) → chạy `python sites.py` để sinh lại `SITES.md`.

**Đây không phải gợi ý — có hook cơ khí:**
- Cài `pre-commit` hook gọi `python sites.py doctor --staged` (exit 1 khi phạm luật) để chặn commit site bẩn.
- Chưa cài hook thì **tự chạy `doctor`** trước khi commit.

Hook cũng không kiểm được "có test live thật hay không" (cần internet). Phần đó
**agent tự chịu trách nhiệm dán bằng chứng test thật ở bước 1**, đừng bịa kết quả.

Mọi agent (Claude / Codex / Gemini) đều phải qua gate này như nhau — không có
ngoại lệ theo agent.

---

## 1. Thư mục này là gì

Bộ **crawler cho từng website nhà cung cấp**. Chuẩn: **1 site = 1 folder (tên =
domain, dấu chấm→gạch dưới) = 1 file `.py` + 1 file `.bat`**. Folder gom theo
**vùng/quốc gia** cho gọn khi upload:

```
WORX GREEN/
  _jp/  _cn/  _vn/  _global/          <- nhóm theo vùng (Nhật/Trung/Việt/US-EU-khác)
    <domain_underscore>/              <- vd vn_bosch-pt_com, digikey_com
      <name>.py                       <- crawler + GUI + CLI trong 1 file
      <name>.bat                      <- @echo off / cd /d "%~dp0" / python <name>.py %* / if errorlevel 1 pause
  crawl_store.py schema.sql products.db product_crawl_schema.json   <- HẠ TẦNG CHUNG ở root
  sites.json  sites.py  SITES.md      <- REGISTRY: site nào đã làm / chưa (xem §1.1)
```
- Mở bat = chạy (mặc định GUI). File .py = crawler-core + GUI + CLI gộp 1 (mẫu:
  `_cn/bosi_cn/bosi.py`, `_cn/zcwz_com/zcwz.py`,
  `_global/digikey_com/digikey.py`).
- Crawler cần schema chung (`crawl_store.py`) thì tự tìm nó bằng cách đi ngược lên
  cây thư mục (`next(p for p in (HERE,*HERE.parents) if (p/'crawl_store.py').exists())`)
  → sống sót khi folder bị đổi nhóm.
- Crawler lớn/nhiều pha CÓ THỂ giữ nhiều file module (GPYH client+crawl, ZCWZ, Kawasaki
  streamlit) — không ép 1 file nếu là project module hợp lệ.
- Chỉ giữ code + bat + `.env`/`requirements.txt` trong folder site; data/output/log
  KHÔNG commit chung (dọn trước khi giao).

## 1.1 REGISTRY — site nào đã làm, site nào chưa

**TRƯỚC KHI làm site mới, TRA REGISTRY. XONG SITE MỚI, CẬP NHẬT REGISTRY.** Đừng làm trùng.

| File | Vai trò |
|---|---|
| `sites.json` | **Nguồn sự thật** — sửa TAY: `kind` / `status` / `note` từng site |
| `sites.py` | Quét folder → đồng bộ json → sinh lại `SITES.md`; kiêm lệnh tra cứu |
| `SITES.md` | Bảng đọc người (tự sinh, ĐỪNG sửa tay) — nhóm theo VN/CN/JP/Global + mục "Còn việc" |

```bash
python sites.py https://www.toptul.com/   # tra 1 site (dán nguyên URL cũng được)
python sites.py                           # quét lại folder + sinh lại SITES.md
python sites.py doctor                    # SOI TẤT CẢ crawler xem có phạm luật file này không
python sites.py doctor toptul             # soi 1 site
```

### `doctor` — ĐỪNG tin trí nhớ, để MÁY kiểm
File md dài, người (và agent) sẽ quên. Nên các luật quan trọng của AGENTS.md được **mã hoá
thành test chạy được** trong `sites.py` (hàm `doctor()`, danh sách `RULES`):

| Luật bị soi | Nguồn |
|---|---|
| Còn `import selenium` | §2 |
| `export_excel` chưa gộp spec vào 1 cột `Thong so` | §4.6 |
| GUI chưa dùng `gui_common` (còn tự chế ô chọn Excel) | §4.4 |
| `norm_code()` viết hoa cả URL | §10 bẫy #1 |
| Thiếu `CIRCUIT_BREAK` | §4.5 |
| Thiếu resume/checkpoint | §5 |
| Thiếu file `.bat` | §1 |
| Còn `_raw_*.jsonl` / `__pycache__` trong folder site | §1 |
| `raise SystemExit` trong file có GUI | §10 bẫy #10 |
| `def worker` GUI không bắt `BaseException` | §10 bẫy #10 |

**Thêm luật mới = thêm 1 dòng vào `RULES`**, KHÔNG phải viết thêm md. Luật nào không kiểm được
bằng máy thì mới ghi md. `doctor` trả exit code 1 nếu có vi phạm → cắm được vào CI/hook.

> **Hiện trạng 2026-10-15:** 57 site, **24 SẠCH / 33 còn vi phạm**. Vi phạm là nợ code cũ, gồm:
> - 27 site thiếu circuit breaker.
> - 23 site thiếu resume.
> - 6 site có worker GUI không bắt `BaseException`.
> - 5 site export chưa gộp `Thong so`.
> - 1 site còn Selenium.
>
> Dọn dần KHI ĐỤNG TỚI site đó. Không refactor hàng loạt nếu user không yêu cầu.
> Xem danh sách: `python -B sites.py doctor`.
- Folder site mới → `sites.py` tự thêm vào json với `status: "done?"` → **sửa note cho thật** rồi chạy lại.
- Site trong json mà mất folder → tự đánh `todo`.
- `status`: `done` (list/detail + spec → Excel) · `partial` (chạy được nhưng THIẾU: chỉ ảnh /
  chỉ tên / chỉ link) · `todo` (chưa có code).
- `note` phải ghi **fact recon** đắt giá: có/không API + endpoint, phân trang, quirk, giới hạn.
  Đây là thứ đọc lại 6 tháng sau vẫn hiểu.

## 1.2 Archetype tham chiếu (clone cái gần site mới nhất)

**Ưu tiên clone crawler SẠCH (doctor `SACH`, có GUI chuẩn + resume + circuit breaker):**

| Folder | Nền tảng | Điểm mấu chốt |
|---|---|---|
| `_vn/vxanh_com` | SSR HTML tự viết | list + sitemap UNION, verify mã exact 2 loại mã |
| `_vn/insumart_vn` | Sapo | `/products.json` CAP offset 30k → chia theo collection + đối chiếu sitemap |
| `_vn/sieuthicongcu_com` | Sapo/Haravan | `products.json` + `/products/<alias>.js`, 1 dòng/variant |
| `_vn/dungcuvang_com`, `_vn/thietbidungcukythuat_com` | WooCommerce | REST/Store API, `sku=` exact |
| `_vn/namtoanviet_vn`, `_global/toptul_com` | PHP / catalog | **CODE-INDEX cache có TTL**, bài học tốc độ §5.1 |
| `_vn/fact-link_com_vn` | PHP SSR, danh bạ DN | **PHP session lock → pool N client cookie riêng**; nguồn search có state; union 4 nguồn; mô tả → `.md` riêng + field → Excel; test fixture HTML thật |

**Archetype cũ** (pattern hay nhưng code còn vi phạm doctor — đọc ý tưởng, ĐỪNG copy nguyên):

| Folder | Nền tảng | Điểm mấu chốt |
|---|---|---|
| `_jp/anextool_co_jp` | WordPress | URL=slug, sitemap XML, spec `<dl><dt><dd>`, direct lookup |
| `_jp/armsangyo_co_jp` | HTML tĩnh | không sitemap → BFS, spec `<table>`, circuit breaker |
| `_jp/tsunoda-japan_com` | HTML song ngữ | sitemap cũ → BFS hub, table header gộp (rowspan/colspan), 1 page nhiều model, direct+BFS-filter |
| `_vn/nguyenkhue_com` | Haravan | API `.js` cho detail + HTML pagination discovery, 429-hibernate |
| `_vn/storelammoc_vn`, `_vn/xsafe_vn` | Sapo | `?view=json` search API + slug-direct theo SKU, **JSON session resume**, bs4/lxml |
| `_cn/zcwz_com` | catalog lớn | **SQLite work-queue**, tenacity, **WAF pause-event**, rate-limiter, 2-pha pipe, master-code-list |
| `_cn/bosi_cn` | HTML tĩnh | brute-force ID enum (⚠ giải pháp CUỐI, xem anti-pattern) |
| `_global/digikey_com` | Next.js + Cloudflare | data nằm trong `__NEXT_DATA__` SSR JSON (khỏi parse HTML); Playwright **warmup homepage** qua Cloudflare; exact-part → `exactMatch`→detail; 5 mode keyword/code/excel/link/all |
| `_vn/trangvangvietnam_com` | danh bạ DN | P0 danh mục → P1 listing → P2 detail, SQLAlchemy |

---

## 2. Stack

### Chuẩn (mặc định dùng)
| Thành phần | Dùng | Lý do |
|---|---|---|
| HTTP async | `httpx.AsyncClient` (pooled, retries, follow_redirects, timeout) | nhanh, giữ IP an toàn |
| Parse HTML | `selectolax.lexbor.LexborHTMLParser` | nhanh hơn bs4 nhiều lần |
| GUI | `tkinter`+`ttk` (theme `clam`) | không cần cài thêm |
| Excel | `pandas`+`openpyxl` | export cột động |
| Concurrency | `asyncio.Semaphore` | giới hạn kết nối |

> **CẤM Selenium.** Chậm, ngốn RAM, hay để process ma. Site tĩnh/có API → `httpx`+`selectolax`.
> Buộc phải chạy browser (JS-heavy/anti-bot) → **Playwright/Patchright**, KHÔNG Selenium.
> Crawler cũ dùng Selenium đã viết lại hết (toptul, idmarket) — không tạo mới.

### Alternatives CHẤP NHẬN (khi có lý do)
- `aiohttp` thay httpx (ZCWZ) — async, ok.
- `requests`+`urllib3.Retry`+`HTTPAdapter` (XSafe/SLM) — bản sync của retry/backoff; hợp khi dùng thread-pool thay asyncio.
- **`BeautifulSoup(lxml)`** khi phải scrape **full HTML page** (breadcrumb, summary, nhiều selector DOM) — selectolax vẫn ưu tiên cho parse gọn, nhưng bs4+lxml hợp cho full-page fallback (XSafe/SLM).
- `SQLAlchemy`+SQLite khi dataset LỚN / cần resume bền (ZCWZ).
- `customtkinter` thay ttk cho GUI đẹp hơn (ZCWZ GUI).
- `tenacity` cho retry có cấu trúc.

Chỉ dùng browser (Playwright/Patchright) khi site **JS-heavy / anti-bot mạnh**.
Site tĩnh hay có API → KHÔNG vác browser vào.

### Bổ sung 2026-10 (đã kiểm phiên bản + license trên PyPI ngày 2026-10-15)
Áp dụng cho crawler **MỚI** và khi đụng tới crawler cũ. KHÔNG viết lại crawler đang chạy tốt chỉ để đổi thư viện.

| Thư viện | Bản (ngày) | Dùng khi | Ghi chú |
|---|---|---|---|
| `curl_cffi` | 0.16.3 (09/2026), MIT | httpx bị 403 / trang challenge do **TLS fingerprint** | Có `AsyncSession` kiểu requests, giả lập TLS/HTTP2 của Chrome. Rẻ hơn mở browser nhiều. |
| `aiolimiter` | 1.3.0 (09/2026), MIT | **Mọi crawler mới**: giới hạn req/giây theo host | Semaphore chỉ giới hạn số request ĐỒNG THỜI, không giới hạn TỐC ĐỘ. Dùng cả hai. |
| `protego` | 0.7.0 (09/2026), BSD | Đọc `robots.txt` + `crawl-delay` | Parser của Scrapy. Đọc 1 lần lúc khởi động, kiểm `can_fetch`. |
| `orjson` | 3.13.0 (10/2026) | Ghi/đọc JSONL checkpoint, API JSON lớn | `orjson.dumps` trả **bytes** → mở file chế độ `ab`. |
| `xlsxwriter` | 3.2.9 (09/2025), BSD | Xuất Excel **lớn** (> ~50k dòng) | `constant_memory` không ngốn RAM. `openpyxl` giữ cho đọc/sửa file có sẵn. |
| `pydantic` v2 | 2.14 (10/2026), MIT | Chuẩn hoá record (SKU, giá, URL) | 1 model / loại record; chuẩn hoá trong model, KHÔNG trong exporter. Tuỳ chọn. |
| `ruff` | 0.17 (10/2026), MIT | Lint + format trước khi giao | `ruff check <file>` · `ruff format <file>`. |
| `uv` | 0.13 (10/2026) | Quản lý môi trường, chạy script 1 file | Header PEP 723 (`# /// script`) khai báo thư viện ngay trong `<name>.py`. Tuỳ chọn. |

**Thang leo khi bị chặn — thử cái RẺ trước, ghi lý do leo vào note `sites.json`:**
JSON/XHR API → sitemap → `httpx` → `curl_cffi` (giả lập TLS) → Playwright → Patchright.

**Luật mới:**
- Công cụ giả lập/stealth CHỈ dùng cho site chặn client thường. Vẫn tôn trọng `robots.txt`, `crawl-delay`, điều khoản site; giữ tốc độ thấp.
- Có `Retry-After` trong phản hồi 429/503 → chờ đúng số giây đó trước khi retry.
- Dữ liệu có cấu trúc trước, CSS selector sau: `__NEXT_DATA__` / JSON-LD (`<script type="application/ld+json">` → `json.loads`) / API.
- **Kiểm license trước khi thêm thư viện:** tránh GPL/AGPL (`ultimate-sitemap-parser` GPL-3, `nodriver` AGPL). Sitemap tự viết ~40 dòng (xử lý index lồng nhau + `.gz`).
- `selectolax` vừa lên **1.0.0 (03/10/2026)**; máy đang dùng 0.4.x. Nâng cấp thì chạy lại toàn bộ test fixture trước.
- KHÔNG dùng làm mặc định: Crawlee, Scrapling (chưa 1.0), rnet (GPL), nodriver (AGPL), Camoufox (chưa 1.0) — chỉ cân nhắc từng ca, ghi lý do.

---

## 3. Quy trình build (BẮT BUỘC theo thứ tự)

### Bước 1 — RECON trước, code sau

#### 1a. SĂN API TRƯỚC TIÊN (ưu tiên số 1)
Việc ĐẦU TIÊN: **săn cho ra API/endpoint JSON**. Có API = crawl nhanh, sạch, ít
vỡ (JSON thẳng). CHỈ khi chắc không có API mới xuống parse HTML/BFS/Playwright.

**Cách săn:** DevTools → Network → **Fetch/XHR**, thao tác search/lọc/phân
trang/mở sản phẩm → request nào trả JSON = API.

**Cheatsheet endpoint theo nền tảng (thử NGAY bằng curl):**
```
Shopify        /products.json?limit=250 , /collections/<c>/products.json , /products/<h>.js
Haravan (VN)   /products/<alias>.js  (detail JSON)   + HTML /collections/all?page=N để list
Sapo (VN)      /search?q=<sku>&view=json  (search JSON) , /products/<alias>.js
WordPress      /wp-json/wp/v2/product , /wp-json/wp/v2/<cpt> , ?rest_route=
WooCommerce    /wp-json/wc/store/v1/products?search=<kw>  (mpemall) , /wp-json/wc/store/products
Magento        /rest/V1/products
Catalog TQ     thường có /search/getmodel.html?keyword= kiểu AJAX trả {code,data}
Next.js (SSR)  data nhúng sẵn trong <script id="__NEXT_DATA__"> -> props.pageProps.envelope.data
               (DigiKey: result-page=topResults, filter=products[], detail=productOverview,
               category=exactMatch). Đọc JSON này = khỏi parse HTML. Phân trang: _next/data/<buildId>/*.json
Khác           /api/search?q= , GraphQL /graphql , Algolia, /api/hybris/* (Grainger)
```
- **API chính thức cần ĐĂNG KÝ (client_id/secret, OAuth) → tránh** trừ khi user có sẵn
  creds. Ưu tiên **internal/embedded API khỏi đăng ký** (Store API, `__NEXT_DATA__`,
  `.js`/`?view=json`). (DigiKey official API cần đăng ký → dùng `__NEXT_DATA__` thay.)
- **Cloudflare/anti-bot mạnh** (DigiKey 403): Playwright load **homepage trước (warmup)**
  cho set cookie clearance rồi mới vào trang sâu; retry chờ `__NEXT_DATA__` xuất hiện;
  `--headed` khi headless bị chặn.
```bash
curl -s -A "Mozilla/5.0" "https://SITE/products/ITEM.js" | head
curl -s -A "Mozilla/5.0" "https://SITE/search?q=SKU&view=json" | head
curl -s -A "Mozilla/5.0" "https://SITE/wp-json/wp/v2/product?per_page=1" | head
```

**Khi săn API, chú ý:**
- **Signed/guest auth**: body request có `sign`/`token` mã hoá? (GPYH: AES-CBC +
  MD5 sign, payload-independent, guest — KHÔNG cần login). Kiểm tra trước khi kết
  luận "phải login". Reverse header/body trong Network.
- **Server result cap**: API chặn N kết quả/query (GPYH cap 4000 nhưng DIN933 có
  16k). → **facet-partition đệ quy**: chia query theo trục (đường kính→vật
  liệu→dài→hãng→bề mặt→cấp) tới khi mỗi nhánh < cap. Dedup theo id.
- **Phân trang**: `per_page`/`page`/cursor; **stuck-page detection** (dừng khi
  trang không thêm id mới) + trần `max_pages`.
- Guest có thể thiếu giá/tồn (login-only) — ghi rõ hạn chế vào docstring.

#### 1b. Câu hỏi recon còn lại (probe site thật bằng curl/httpx)
```
CÓ API/JSON không? (săn ở 1a) -> nếu có, ưu tiên tuyệt đối. Nền tảng gì?
robots.txt / sitemap.xml có không? MỚI hay CŨ/thiếu? (Tsunoda: sitemap 2017 vô dụng)
CMS gì? WordPress/Haravan/Sapo/Shopify/Magento/tĩnh?
URL trang sản phẩm dạng gì? slug có = mã sản phẩm không?
Spec ở đâu? <table> / <dl> / <div> / trong JSON? Header table có gộp rowspan/colspan?
1 trang = 1 model hay NHIỀU model (nhiều dòng bảng)?  (quyết định cần crawl+filter)
Ảnh URL dạng gì? suy từ mã được không? có //cdn thiếu https: không?
Song ngữ? charset? (thiếu charset header -> force utf-8)
Có WAF/rate-limit/429 không? (thử fetch nhanh vài chục request xem có bị chặn)
```

### Bước 2 — Chọn chiến lược khám phá (discovery)
| Tình huống | Chiến lược |
|---|---|
| **CÓ API/JSON** | **Crawl thẳng API (phân trang/cursor) — ưu tiên số 1** |
| API có cap kết quả | Facet-partition đệ quy theo trục tới khi < cap (GPYH) |
| Có sitemap ĐẦY ĐỦ & MỚI | Load sitemap → list URL (Anex) |
| Sitemap thiếu/cũ / không có | **BFS** hub/category pages gom link sản phẩm (ARM/Tsunoda) |
| Slug = mã sản phẩm | Direct lookup `build_url(code)` (Anex/Tsunoda) |
| Mã sống dạng row trong trang khác | Crawl + **filter theo mã** (ARM/Tsunoda) |
| Có sẵn/ sinh được list mã | **Master-code-list** → search API → detail (ZCWZ) |
| List phân trang, không biết tổng | Pagination **dừng khi count chững** (NguyenKhue) |
| Nhập SKU, cần resolve URL | Search API resolve + **khớp mã chính xác** (XSafe/SLM) |
| KHÔNG còn cách nào | Brute-force ID enum — ⚠ CUỐI CÙNG, chậm/bẩn (Bosi) |

Kết hợp được: direct trước (nhanh) → mã miss thì BFS/crawl filter (Tsunoda `_run_model`).

### Bước 3 — Implement theo template (mục 4–7)
### Bước 4 — TEST end-to-end headless (mục 9) — CHƯA test CHƯA done
### Bước 5 — Báo done kèm **bảng kết quả test thật**

---

## 4. Template code

### 4.1 Tách CORE khỏi GUI (để test headless)
Hàm thuần không đụng GUI:
- `norm_code(s)` — chuẩn hoá mã so khớp (bỏ trắng, upper).
- `build_url(code)` / `url_candidates(code, lang)` — mã → URL.
- `parse_page(html|json, entry, url) -> list[dict]` — 1 record/model.
- `parse_spec_table`/`build_grid` — nếu spec là table (xử lý rowspan/colspan).
- Engine async chạy độc lập (class `Engine`, xem Tsunoda `TsunodaEngine`).
- `export_excel(results, path)` — cột động.

### 4.2 Pipeline chuẩn (GUI radio; thực tế 3–4 tuỳ site)
1. **Theo MÃ** — nhập 1 mã → direct/crawl-filter (hoặc exact-match→detail).
2. **Theo keyword** — search từ khoá → list sản phẩm.
3. **Theo slug/URL / link** — dán 1..n link (ô multi-line, mỗi dòng 1 link) → detail.
4. **Theo cột Excel** — chọn file → sheet → cột (mã hoặc URL).
5. **Crawl ALL** — quét sạch (sitemap/BFS/API/pagination/category-tree) + `--limit`.
(NguyenKhue: ALL/KEYWORD/EXCEL; ZCWZ: all/brand-deep/excel; DigiKey:
keyword/code/excel/link/all. Số pipeline tuỳ site — miễn phủ được các cách nhập.
Mode `link` để **ô textbox nhiều dòng**, hint rõ "mỗi dòng 1 link".)

### 4.3 Khớp mã CHÍNH XÁC khi resolve (đừng tin blind kết quả đầu)
Search API trả nhiều kết quả → **khớp mã word-boundary**, đừng lấy đại kết quả đầu:
- Tier: exact-match (biên từ, WU380S ≠ WU380S.9) → substring → mới fallback first.
- Heuristic SKU từ title/URL nếu API không có (NguyenKhue).

### 4.4 GUI skeleton — ô chọn Excel PHẢI dùng `gui_common.py` (hạ tầng chung ở root)
KHÔNG bắt user gõ tay tên sheet/cột. KHÔNG copy-paste khối chọn Excel vào từng crawler.
```python
root_dir = next(p for p in (HERE, *HERE.parents) if (p / "gui_common.py").exists())
sys.path.insert(0, str(root_dir))
from gui_common import ask_export_path, attach_excel_picker

attach_excel_picker(left, xls, sheet, col, outdir)   # Sheet+Cot = dropdown đọc từ file Excel
p = ask_export_path(xls.get(), HERE, "<site>.xlsx")  # lưu CẠNH file Excel đầu vào
```
- Chọn Excel → tự nạp sheet_names → tự nạp cột của sheet → **tự đoán cột mã** (`part|mã|code|sku|item`).
- Thư mục ảnh tự đặt cạnh file Excel (`anh_<tên file>/`); EXPORT mặc định `<tên file>_ketqua.xlsx`.

### 4.4b GUI phần còn lại (copy ARM/Tsunoda, đổi màu brand)
- Trái: pipeline + input + Concurrency + Delay + START/Pause/Stop/EXPORT.
- Phải: console tối + progressbar + stats `x/y | z model`.
- Log qua `queue.Queue`, drain `root.after(100,...)` (thread-safe).
- Worker `threading.Thread(daemon=True)` → `asyncio.run(...)`.
- GUI có thể redirect `sys.stdout/stderr` vào console textbox (ZCWZ) để bắt log async.

### 4.5 Safety consts (BẮT BUỘC)
```python
MAX_REDIRECTS = 5
CIRCUIT_BREAK = 25       # dừng sau N lỗi LIÊN TIẾP (nghi block IP)
MAX_ITEMS     = 5000     # chặn work-list phình
MAX_BFS_PAGES = 1500     # chặn BFS vô tận
MAX_PAGES     = 600      # trần phân trang API
```

### 4.6 Export Excel
- Cột cố định đầu: `Input, Page/Name, Item No, Status`.
- **Spec gộp 1 CỘT `Thong so`** — mỗi dòng `Tên: Giá trị`, ngăn `\n`, bật `wrap_text` + width 60.
  KHÔNG trải union key thành cột động: hãng đặt tên spec tuỳ hứng → nổ ra hàng trăm cột rỗng
  (totalbusiness.com: 543 cột, 207 key chỉ xuất hiện 1 lần, có cả typo `Impcat rate`). Dòng
  không có dấu `:` (bullet mô tả) nối luôn vào cột này. Category vẫn tách `Category N`.
- Cột cuối: `Image, URL`. Mã miss vẫn có dòng `Not Found`.
- Ảnh `//cdn...` → tự thêm `https:`.
- **BẮT BUỘC phẳng & người-thường-đọc-được**: mỗi ô là scalar (string/số). TUYỆT ĐỐI
  không đổ nguyên dict/list JSON vào ô (vd DigiKey bọc `{type:'stringList',value:[{label}]}`,
  price `{qtyAvailable,pricing:[...]}`). Phải unwrap: stringList→join label; price→tách
  `UnitPrice`/`PriceBreaks("1:$0.32; 10:$0.22")`/`Stock`/`Packaging`; attributes→mỗi
  thông số 1 cột. Nested list→`", ".join(...)`. Viết hàm `_clean()`/`flatten()` cho field bọc.
- Test nhanh: mở file, không được thấy `{'...'}`/`[{...}]` trong bất kỳ ô nào.

---

## 5. Resume / checkpoint — BẮT BUỘC cho crawl vừa→lớn

Crawl trăm–nghìn item mà crash/Stop là mất hết → PHẢI có checkpoint để **chạy lại
là tiếp tục, không làm lại từ đầu**. Chọn 1 trong 3 flavor theo quy mô:

| Flavor | Khi nào | Cơ chế |
|---|---|---|
| **JSONL sidecar** (GPYH) | crawl API/lớn, CLI | mỗi item mới → append+flush `_raw_<key>.jsonl`; start reload → rebuild `seen`; key = sanitize(mode+target) để đúng lệnh chạy tiếp; xoá khi xong sạch, GIỮ khi crash |
| **JSON session** (XSafe/SLM) | GUI, list SKU | `save_session()` sau MỖI item vào `.<site>_session.json`; start so fingerprint list input → hỏi "Resume?" → skip mã đã done; clear khi xong |
| **SQLite work-queue** (ZCWZ) | rất lớn, bền nhất | bảng `urls(url UNIQUE, crawled BOOL)` = queue + dedup; `products(url UNIQUE,...)` detail; `session.merge()` upsert; Pipe2 chỉ lấy `crawled=False`, `mark_crawled()` sau mỗi cái. Idempotent xuyên run. PRAGMA `journal_mode=WAL, synchronous=OFF` cho ghi nhanh |

Kèm **crash-proof**: bọc crawl trong `try/except (KeyboardInterrupt, Exception)/finally`,
`finally` LUÔN ghi data đã gom. **Autosave từng phần** khi đang chạy (GPYH flush
xlsx mỗi 2000; NguyenKhue append CSV mỗi 50).

⚠ CSV `mode='a'` KHÔNG phải resume (Bosi) — restart vẫn crawl lại từ đầu. Đừng nhầm.

---

## 5.1 TỐC ĐỘ — ĐO trước, đừng đoán (bài học toptul 2026-07-13)

Crawl chậm thì **đo từng chặng** (search / detail / parse), đừng đoán rồi tối ưu nhầm chỗ:

| Chặng | toptul (đo thật) |
|---|---|
| **Search 1 mã** | **3,4 s** ← 99% thời gian nằm ở đây |
| Tải trang detail | 0,04–0,2 s |
| Parse HTML 700 KB (selectolax) | 0,002 s |

→ Parse **không bao giờ** là nút cổ chai. Nút cổ chai là **số round-trip mạng**. 3 chiêu:

1. **CODE-INDEX (mã → URL), cache ra đĩa + TTL.** Site kiểu catalog: **1 trang chứa cả series**
   (toptul: 1 trang = 26 mã BAEA1608..1634). Tải 1 trang → ghi nhớ MỌI mã trong đó → các mã cùng
   series **khỏi search**. Kết quả thật: 20 mã cùng series **24,9 s → 5,2 s**; chạy lại **1,3 s**.
2. **Mẫu rải đều rồi bung song song — TỐI ĐA 2 VÒNG CHỜ.** ĐỪNG chia đợt tuần tự để "học dần"
   (thử rồi: 24,9 s, tệ hơn cả song song thuần). Vòng 1 search một mẫu rải đều → index mở khoá
   cả cụm; vòng 2 bung song song phần còn lại.
3. **Quét sitemap dựng index toàn catalog** khi list dài (`BUILD_INDEX_FROM = 40`): trang detail
   rẻ (0,04 s) nên 2,5k trang ≈ 100 s, đổi lại tra mã **tức thì mãi về sau** — rẻ hơn search
   40 mã × 3,4 s. Có cờ `--build-index` chạy tay.

**Concurrency phải ĐO, không phóng đại**: toptul `conc=16` → **rớt 10/40 request** → retry backoff
làm crawl **chậm gấp 5** (60 mã: 88 s). `conc=8` → 40/40 OK. Chậm bất thường mà không thấy lỗi →
nghi retry đang nuốt thời gian, hạ `conc` xuống rồi đo lại.

**Song song mà KHÔNG nhanh hơn → nghi SESSION LOCK** (bài học fact-link 2026-10-01). Site PHP
dùng `session_start()` sẽ XẾP HÀNG mọi request cùng `PHPSESSID`.
- Đo thật: 8 GET song song qua 1 client = **48 s**; qua 8 client cookie riêng = **8 s**.
- Cách đo nhanh: tách TTFB file tĩnh (`robots.txt`) với trang động bằng `curl -w "%{time_starttransfer}"`.
  Rồi so N request trên 1 client với N client riêng.
- Gặp session lock thì dùng **pool N `httpx.AsyncClient`, mỗi client 1 cookie jar**. Mỗi request mượn 1 client.
  Tham khảo `Net` trong `_vn/fact-link_com_vn`.

## 6. Retry / rate-limit / 429-WAF

### Retry + backoff (transient error: Timeout, TransportError, 5xx, 429)
- tenacity: `@retry(stop_after_attempt(5), wait_exponential(multiplier=2,min=5,max=60), retry_if_exception_type((ClientError,TimeoutError)))` (ZCWZ).
- Hand-rolled: `sleep(min(2**i, 20))`, ~6 lần (GPYH).
- Sync: `urllib3.Retry(total=3, backoff_factor=0.3, status_forcelist=[500,502,503,504,429])` + HTTPAdapter (XSafe/SLM).
  **⚠ mặc định urllib3 KHÔNG có 429 trong forcelist — PHẢI thêm tay.**

### 429 / WAF — chống bị block IP
- Đếm **429 liên tiếp** → backoff leo thang; ngưỡng cao → **hibernate 60–120s**
  (NguyenKhue: ≥10×429 → ngủ 120s).
- **Global pause-event** (`asyncio.Event` chia sẻ) — gặp 403/429/503/520/521/522
  thì `clear()` event, TẤT CẢ coroutine chờ; task nền `_resume_after(60s)` set lại
  → rồi để retry backoff xử tiếp (ZCWZ). Mạnh hơn breaker per-task.
- Circuit breaker: N lỗi liên tiếp → `stop_evt.set()`.

### Rate limit
- **Token-bucket** `rate_limit_per_second` + async lock (ZCWZ) — mịn hơn semaphore.
- Hoặc `delay` per-request cấu hình (mặc định 0.05–0.2s) + semaphore/thread-pool.
- Crawler MỚI: `aiolimiter.AsyncLimiter(<req>, 1)` theo host + Semaphore (xem §2 "Bổ sung 2026-10").

---

## 7. Kiến trúc lớn (khi site rất lớn)

### 2-pha pipeline (ZCWZ)
- **Pipe1 discover**: quét list/API → nhét URL vào queue (DB/JSONL), dedup unique.
- **Pipe2 scrape**: lấy URL chưa crawl từ queue → parse detail → lưu → mark done.
- Tách qua queue → resume, chạy lại pha nào cũng được, không giữ hết trong RAM.

### CLI mode (không chỉ GUI)
- `argparse` (`--mode --keyword --url --excel --sheet --col ...`) + fallback menu
  `input()` khi không truyền mode (GPYH). Cùng 1 script chạy `.bat` lẫn scripted.
- Crawler lớn nên có CORE chạy headless (đã yêu cầu ở 4.1) → CLI dễ.

### File logging
- `logging` + `FileHandler(<name>.log)` + `StreamHandler`; GUI redirect stdout vào
  cùng stream (ZCWZ). NguyenKhue dùng `RotatingFileHandler`. Lưu lỗi để debug sau.

---

## 8. Ngôn ngữ & nội dung
- Docstring đầu file: **liệt kê fact recon site** (có/không API + endpoint, CMS,
  URL pattern, spec ở đâu, quirk). Đây là tài liệu quan trọng nhất của crawler.
- UI text tiếng Việt. Log prefix: `[SYS] [OK] [HIT] [MISS] [WARN] [ERR] [STOP] [FATAL]`.
- Site song ngữ: cho chọn EN/JP, ưu tiên ngôn ngữ chọn, thiếu thì fallback.
- Catalog nước ngoài: có thể xuất **đa sheet đa ngữ** (zh/en/vi) dùng
  `deep_translator`, **cache dịch ra đĩa** (batch + fallback từng item), export
  từng ngôn ngữ sau khi dịch xong để dừng giữa chừng không mất (GPYH).

---

## 9. An toàn runtime & tài nguyên (end-to-end) — BẮT BUỘC

Crawler chạy trên **máy user**: không treo máy / tràn RAM / để process ma.

### 9.1 Tắt GUI = tắt SẠCH
```python
def _on_close(self):
    if self.is_running and not messagebox.askyesno("Thoat", "Dang crawl. Thoat va tat luon?"):
        return
    self.stop_evt.set()
    try: self.root.destroy()
    finally: os._exit(0)      # kill cứng: không để thread asyncio/console treo
self.root.protocol("WM_DELETE_WINDOW", self._on_close)
```
- Worker `daemon=True` → chết theo main. `os._exit(0)` (KHÔNG `sys.exit`) để đóng
  luôn terminal .bat, không để cửa sổ đen + thread nền + socket pool ăn RAM.

### 9.2 Chống loop vô tận — mọi vòng phải có phanh
- **Redirect guard**: fetch xong check URL cuối còn trong vùng sản phẩm
  (`"/itempage/" in final`); trượt ra homepage/search → "Not Found", KHÔNG parse.
- `follow_redirects` + `max_redirects` + bắt `TooManyRedirects`. Pagination có thể
  `follow_redirects=False`, 301/302 = hết trang → dừng (NguyenKhue).
- **BFS `visited` set + MAX_BFS_PAGES**; frontier chỉ thêm URL mới.
- **Pagination**: `max_pages` + **stuck-page** (dừng khi không thêm id mới).
- **Circuit breaker** + dedup work-list + `timeout` mọi request.

### 9.3 Không tràn RAM
- Giới hạn cứng `MAX_ITEMS/MAX_BFS_PAGES/MAX_PAGES`; vượt → cắt + log, không ôm hết.
- `asyncio.Semaphore(conc)` / thread-pool giới hạn — không mở nghìn socket.
- **Vứt HTML thô sau parse**, chỉ giữ dict record gọn.
- Crawl rất lớn: **flush ra đĩa mỗi N item + `results.clear()` + `gc.collect()`**
  (NguyenKhue) hoặc dùng DB queue (§7) thay vì giữ hết trong list.
- BFS bằng vòng lặp + frontier, KHÔNG đệ quy sâu vô hạn.

### 9.4 Tôn trọng pause/stop trong MỌI vòng lặp (kể cả BFS)
```python
while self.pause_evt.is_set() and not self.stop_evt.is_set(): await asyncio.sleep(0.2)
if self.stop_evt.is_set(): return
# re-check stop SAU khi giành semaphore slot
```

### 9.5 Test các điểm này (không chỉ happy path)
- Stop giữa chừng → dừng ~1s, giữ kết quả, không treo.
- Pause/Resume → đúng chỗ.
- Đóng cửa sổ khi chạy → Task Manager không còn python.exe ma.
- Mã sai/redirect → ra `Not Found`, không loop, không nổ circuit oan.
- Crawl ALL site lớn → RAM ổn định, không leo vô hạn.
- Chạy lại sau khi Stop → **resume tiếp**, không làm lại từ đầu.

---

## 10. Chất lượng & anti-pattern

### Quality bar
1. **API-first**: săn API trước, docstring ghi rõ có/không + endpoint.
2. **Recon-driven**: docstring phản ánh đúng site thật, không đoán.
3. **Simple**: copy archetype gần nhất, sửa recon; không abstraction thừa.
4. **Robust**: redirect/404/table-header-gộp/thiếu-charset/mã-là-row/khớp-mã-chính-xác.
5. **An toàn IP**: retry backoff + rate-limit + 429/WAF handling + circuit breaker.
6. **Bền**: resume checkpoint + crash-proof finally + autosave.
7. **An toàn máy**: os._exit close, daemon, semaphore, giới hạn RAM, gc/flush.
8. **Testable**: core tách GUI, chạy headless/CLI được.
9. **Coverage tối đa**: kết hợp direct + crawl-filter + fallback ngôn ngữ.

### ĐỪNG làm
- Viết code trước khi recon / **bỏ săn API lao vào parse HTML ngay**.
- **Selenium** (bất kỳ lý do gì) — dùng httpx/selectolax; buộc browser thì Playwright.
- Playwright cho site tĩnh hoặc có API.
- Tin sitemap mù mà không kiểm nó cũ/thiếu.
- **Brute-force ID enum** khi còn cách khác (chậm, bẩn — chỉ dùng cuối).
- **`verify=False`** tắt TLS (Bosi) — không làm trừ khi bất khả kháng.
- **CSV append tưởng là resume** — không phải, restart vẫn làm lại.
- **urllib3 Retry quên 429** trong status_forcelist.
- Lấy blind kết quả search đầu tiên mà không khớp mã chính xác.
- Báo done mà chưa chạy thử mã thật.
- Bỏ circuit breaker/rate-limit → crawl mù khi bị block.
- Parse table bằng regex ngây thơ, bỏ qua rowspan/colspan.

### BẪY ĐÃ DÍNH THẬT (2026-07-13, review 4 crawler) — kiểm đúng mấy chỗ này
1. **`norm_code()` viết hoa CẢ URL** → cột link trong Excel thành `HTTPS://SITE/EN/PRODUCT-...`
   → 404 (path phân biệt hoa/thường). Luôn: `return v if v.lower().startswith("http") else v.upper()`.
2. **Trộn LINK chung với CODE rồi lọc theo `codes`** → record của link bị xoá sạch (Item No của
   nó đâu nằm trong danh sách mã). Crawl link RIÊNG, không áp bộ lọc mã.
3. **Nạp checkpoint 2 lần** (gọi `_load_ckpt` ở nhánh link rồi nhánh code) → record nhân đôi.
   Đặt cờ `self._ckpt_loaded`.
4. **Điều kiện fallback dùng `not self.results`** → khi RESUME (results đã có từ checkpoint),
   fallback bị nuốt → kết thúc lặng lẽ với data thiếu mà vẫn báo "XONG". Guard bằng cờ đường
   dẫn (`api_ok`/`search_ok`), KHÔNG bằng `results` rỗng.
5. **Lấy mù kết quả search đầu** (§4.3) → ghi Status OK cho SẢN PHẨM SAI. Phải mở hit, so mã
   thật (site VN thường có 2 mã: **mã hệ thống** + **mã hãng/model** — so CẢ HAI); lệch thì vẫn
   lấy nhưng ĐÁNH DẤU trong `Status`.
6. **Báo "Not Found" oan**: SP đã vào kết quả qua dòng LINK (dedup theo URL) → tập "đã có" chỉ
   so `Input` sẽ không thấy mã của nó. So cả mã hệ thống + mã hãng.
7. **API phân trang trả TRÙNG** (site đang sync giá/tồn → đổi thứ tự giữa các trang): idmarket
   trùng 101/3000 ngay cả khi chạy sạch. LUÔN dedup theo URL/id ở tầng gom.
8. **Cache sitemap không TTL** → SP mới của site không bao giờ thấy. Đặt TTL (vd 7 ngày).
9. **Sitemap `_group_1` không tồn tại** (toptul): file gốc tên khác — đừng đoán dãy số, kiểm 404.
10. **`raise SystemExit` trong code thư viện + worker GUI chỉ bắt `Exception`** → `SystemExit` là
    `BaseException`, lọt qua → thread chết câm, nút START vô hiệu vĩnh viễn. Thư viện `raise ValueError`;
    worker `except BaseException` + `finally` trả lại state `running=False`. (doctor có soi.)
11. **HTML lồng sai chuẩn** (fact-link: `<p class="desc"><p>…</p></p>`) → lexbor đóng thẻ ngoài ngay →
    text RỖNG mà không báo lỗi. Fixture một ngôn ngữ (EN) không lồng nên test vẫn xanh. Test fixture
    MỌI ngôn ngữ/template mà crawler dùng. Field rỗng 100% trên mẫu live = nghi parser, ĐỪNG kết luận
    "site không có dữ liệu".
12. **Trang danh sách có STATE trong session** (fact-link `search_post.php` → `search_result.php`): gọi thẳng
    trang kết quả = rỗng. Mỗi nguồn giữ 1 client riêng suốt phân trang. Trang lỗi tạm thời thì thử lại.
    Chỉ dừng sau 3 trang xấu LIÊN TIẾP, không dừng ở trang ngắn đầu tiên.
13. **Báo cáo "đã sửa" ≠ đã sửa.** Reviewer phải tự chạy lại test/doctor và mở dữ liệu output thật.
    Fact-link: worker báo đã fix nhãn `::` nhưng code chưa đổi.

---

## 11. Checklist trước khi báo done
- [ ] **Đã tra REGISTRY (§1.1)** trước khi làm — không làm trùng site đã có.
- [ ] **Đã săn API trước tiên** (DevTools/Network + cheatsheet + curl); có thì dùng.
- [ ] Recon xong, docstring phản ánh đúng site (ghi rõ có/không API + endpoint).
- [ ] Discovery đúng chiến lược (API/sitemap/BFS/slug/master-code/pagination).
- [ ] Pipeline nhập liệu đủ (model/codes/all/excel hoặc tương đương site).
- [ ] Khớp mã chính xác khi resolve (không blind first-result).
- [ ] Safety consts đủ (circuit break, max items/bfs/pages, redirect guard, timeout).
- [ ] Retry + backoff (429 CÓ trong danh sách) + rate-limit + 429/WAF handling.
- [ ] Anti-loop: BFS `visited`, stuck-page, redirect guard, dedup.
- [ ] **Resume/checkpoint** (JSONL / JSON-session / SQLite) + crash-proof finally + autosave.
- [ ] Tắt GUI = `os._exit(0)`, worker `daemon=True` → không process ma, không treo terminal.
- [ ] RAM: semaphore/thread-pool giới hạn, vứt HTML thô, gc/flush hoặc DB queue cho crawl lớn.
- [ ] Pause/Stop tôn trọng mọi vòng (kể cả BFS/pagination).
- [ ] Core tách khỏi GUI (chạy headless/CLI được).
- [ ] File logging (nếu crawl lớn).
- [ ] Test headless với **mã thật** → bảng kết quả (OK/miss + lý do).
- [ ] Test Stop / Pause / đóng-cửa-sổ / **resume-sau-Stop** → không process ma, RAM ổn.
- [ ] Export Excel mở được (cột spec đúng, mã miss vẫn có dòng).
- [ ] `<name>.bat` tạo sẵn cạnh `<name>.py` (`@echo off` / `cd /d "%~dp0"` / `python <name>.py %*` / `if errorlevel 1 pause`).
- [ ] Crawler mới: có giới hạn req/giây (`aiolimiter`) + đọc `robots.txt` (`protego`); `ruff check` sạch; thư viện mới không GPL/AGPL (§2 Bổ sung 2026-10).
- [ ] Test pytest `tests/<name>/` với fixture HTML THẬT (đủ mọi template/ngôn ngữ crawler dùng) → pass.
- [ ] **Đi hết 13 BẪY ĐÃ DÍNH THẬT ở §10**:
      - URL bị upper, link trộn code, checkpoint nạp 2 lần.
      - Fallback guard sai, lấy mù kết quả search, Not Found oan.
      - API phân trang trùng, cache không TTL, đoán dãy sitemap.
      - SystemExit, HTML lồng sai chuẩn, nguồn có session state, báo cáo chưa kiểm.
- [ ] Không để rác trong folder site: chỉ `.py` / `.bat` / `.env` / `requirements.txt` + `out_<name>/` (gitignore).
- [ ] **Test Excel có LẪN cả mã lẫn LINK** (không chỉ toàn mã) — đây là chỗ hay vỡ nhất.
- [ ] **Cập nhật REGISTRY** (§1.1): sửa `sites.json` (status + note recon thật) → `python sites.py`.
- [ ] **`python sites.py doctor <site>` → phải ra "SACH"** (máy tự soi các luật §2/§4.4/§4.5/§4.6/§5/§10).
- [ ] Báo done kèm bảng test + đường dẫn file.
