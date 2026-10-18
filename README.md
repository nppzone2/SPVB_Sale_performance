# SPVB Sale Performance

Dashboard Sale Performance tháng 10/2026, đưa lên GitHub Pages.

## Cập nhật dữ liệu
1. Đặt ba file vào thư mục `input/` (tên file có chứa: `San_luong`/`SKU`, `Danh_sach`, `Loyalty`).
2. Đẩy lên nhánh `main`. Workflow sẽ tự tạo lại trang và đưa lên Pages.
3. Target: vào tab **Nhập Target** trên trang, nhập số, bấm **Tải file target**, rồi đưa file `Target_*.xlsx` vào `input/` và đẩy lên.

## Cấu trúc
- `input/` — file Excel đầu vào
- `src/template.html` — giao diện và cách tính (chạy trên trình duyệt)
- `engine/build.py` — đọc `input/`, nhúng dữ liệu vào trang, tạo `site/index.html`
- `.github/workflows/build-deploy.yml` — tự động build và deploy
