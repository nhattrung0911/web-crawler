# web-crawler

Bộ crawler dữ liệu sản phẩm/danh bạ cho các website nhà cung cấp (VN / CN / JP / Global).
Mỗi site = 1 folder `_<vùng>/<domain>/` gồm `<name>.py` (core + GUI + CLI) và `<name>.bat`.

- Đọc **`AGENTS.md`** trước khi sửa hoặc thêm crawler — quy chuẩn, archetype, bẫy đã gặp.
- **`SITES.md`**: danh sách site + ghi chú recon. `python sites.py doctor` soi code theo luật.
- Cài: `pip install httpx selectolax openpyxl pandas tenacity pytest` (Python 3.13).
- Chạy: mở `<name>.bat` (GUI) hoặc `python <name>.py --mode ...` (CLI).
- Dữ liệu crawl ghi vào `out_<name>/` (không commit).

Dùng có trách nhiệm: tôn trọng robots.txt và điều khoản của từng website, giữ tốc độ hợp lý.
