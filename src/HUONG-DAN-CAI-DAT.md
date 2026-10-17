# Vật Lý Cô Lâm Cường, bản 7: máy chủ mới, nhanh gấp 20–50 lần

## Vì sao đổi
Đo thật ngày 13/10: lệnh rỗng (`ping`) trên Apps Script mất 1,4–2,2 giây; tra điểm lần đầu 8,5 giây. Đó là phí cố định
của Google Apps Script, không giảm được. Bản 7 chuyển máy chủ sang **Cloudflare Worker + D1** (miễn phí, máy chủ tại
Việt Nam): đo cục bộ mỗi lượt 15–130 ms. Toàn bộ nghiệp vụ là code cũ đã kiểm thử, chạy nguyên vẹn trên lớp tương thích.

## Dữ liệu
- `du-lieu/xuat-tu-sheet-2026-10-17/`: bản xuất đầy đủ 14 bảng hệ thống + 7 tab danh sách lớp (JSON).
- `du-lieu/migrate.sql`: dữ liệu đó ở dạng nạp vào D1.
- File Google Sheet cũ **không bị xóa**: giữ làm sao lưu, và máy chủ mới tự chép lại toàn bộ dữ liệu về đó mỗi đêm 2h.
- File "DANH SÁCH LỚP" vẫn là nơi mẹ sửa tay. Trên web bấm **Đồng bộ nhóm** (trang Lớp) để kéo về; sửa trên web thì tự đẩy lên Sheet.

## Vai trò còn lại của Apps Script (cầu nối)
Drive (ảnh phiếu, file đáp án), gửi email, đọc/ghi danh sách lớp, nhận sao lưu. Vẫn dùng project cũ, chỉ thêm file.

## Cài đặt
### A. Apps Script (cầu nối)
1. Thêm file **Bridge**; dán đè **Bai, Code, Config, Setup** (4 file). Ctrl+S.
2. Chạy hàm `taoKhoaCauNoi` (chọn trong ô hàm, bấm Run) → xem Execution log, chép chuỗi `BRIDGE_SECRET`.
3. Deploy → Manage deployments → bút chì → New version → Deploy. URL web app **không đổi**.

### B. Cloudflare (máy chủ mới)
Cần tài khoản Cloudflare (miễn phí). Hai cách:
- **Cách 1, t làm thay qua trình duyệt trong app Claude** khi máy m đang kết nối và đã đăng nhập Cloudflare.
- **Cách 2, m tự chạy** trong thư mục `worker/` theo `worker/README.md` (7 lệnh, có Node là chạy được).

Kết quả là một địa chỉ dạng `https://vatlycocuong-api.<tên>.workers.dev`.

### C. Web (GitHub, repo vatlycocuong)
1. Tải cả thư mục `web/` lên (đè). Thư mục này không có `config.js`.
2. Sửa `config.js` trên GitHub: `API_URL` = địa chỉ Worker ở bước B.
3. Ctrl+Shift+R. Menu tài khoản → **Đo tốc độ** để thấy số mới.

### D. Quay lại bản cũ nếu cần
Chỉ cần đổi `API_URL` trong `config.js` về URL Apps Script. Dữ liệu Sheet cũ vẫn còn nguyên tới lúc chuyển
(và được cập nhật lại mỗi đêm từ máy chủ mới), nên quay lại không mất gì ngoài thao tác sau đêm sao lưu gần nhất.
