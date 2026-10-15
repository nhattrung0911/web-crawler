# web-crawler

Bộ crawler dữ liệu sản phẩm/danh bạ cho các website nhà cung cấp (VN / CN / JP / Global).
Mỗi site là một folder độc lập `_<vùng>/<domain>/` gồm `<name>.py` (core + GUI + CLI) và `<name>.bat`.

**Sơ đồ tech stack + cách dùng:** https://web-crawler-diagram.pages.dev

## Cài đặt
```bash
pip install httpx selectolax openpyxl pandas tenacity pytest   # Python 3.13
pip install playwright && playwright install chromium          # chỉ khi crawler cần browser
```

## Chạy một crawler
```bash
cd _vn/vxanh_com
vxanh.bat                       # mở GUI (Windows)
python vxanh.py --help          # CLI: xem các mode của từng crawler
```
Kết quả (Excel, checkpoint, log) ghi vào `out_<name>/` trong folder site, không commit.
Dừng giữa chừng rồi chạy lại = chạy tiếp, không làm lại từ đầu.

## Cấu trúc
| Đường dẫn | Nội dung |
|---|---|
| `_vn/ _cn/ _jp/ _global/` | Crawler theo vùng |
| `AGENTS.md` | Quy chuẩn viết crawler: quy trình, archetype, bẫy đã gặp — đọc trước khi sửa/thêm |
| `SITES.md`, `sites.json` | Danh sách site + ghi chú recon (sinh bởi `python sites.py`) |
| `sites.py doctor` | Soi code theo luật (circuit breaker, resume, export...) |
| `gui_common.py` | Thành phần GUI dùng chung |
| `tests/` | pytest với HTML thật đã lưu |
| `site/` | Trang sơ đồ (Cloudflare Pages) |

## Kiểm tra
```bash
python -B -m pytest tests -q -p no:cacheprovider
python -B sites.py doctor
```

## Dùng có trách nhiệm
Tôn trọng robots.txt và điều khoản của từng website, giữ tốc độ vừa phải. Crawler có sẵn
giới hạn đồng thời, backoff khi gặp 429 và tự dừng khi lỗi liên tiếp.
