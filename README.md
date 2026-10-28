# SPVB Sale Performance

Dashboard Sale Performance tháng 10/2026, đưa lên GitHub Pages. Dữ liệu được mã hoá bằng mật khẩu nhóm.

## Cài đặt một lần
1. Vào **Settings → Secrets and variables → Actions → New repository secret**, tạo secret tên `SPVB_PASSWORD` với mật khẩu nhóm.
2. Mật khẩu này cũng là mật khẩu mở dữ liệu trên trang. Chia sẻ cho người được phép xem.

## Cập nhật dữ liệu
1. Mở trang dashboard → tab **Dữ liệu nguồn** → **Chọn file và gửi lên repo**.
2. Chọn file Excel (BC Sản lượng, DS Loyalty hoặc Danh sách ASO). Trang mã hoá file ngay trên máy rồi gửi lên `input/`.
3. Đợi khoảng 2–3 phút để workflow build lại, sau đó làm mới trang.

## Target
Tab **Nhập Target** → **Lưu target lên repo** (cần token GitHub đã lưu trên máy). Target được mã hoá và lưu vào `targets/targets.enc`.

## Cấu trúc
- `input/*.xlsx.enc` — dữ liệu đã mã hoá
- `targets/targets.enc` — target đã mã hoá
- `src/template.html` — giao diện và cách tính (chạy trên trình duyệt)
- `engine/build.py` — giải mã input, tạo `site/data.enc` và `site/index.html`
- `.github/workflows/build-deploy.yml` — tự động build và deploy
