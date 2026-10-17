# Máy chủ mới: Cloudflare Worker + D1

- `src/index.js`: nhận lượt gọi, khóa khi ghi, nạp dữ liệu, chạy nghiệp vụ, ghi D1, việc phụ (mail, Drive, đồng bộ Sheet).
- `src/shim.js`: lớp tương thích thay SpreadsheetApp/CacheService/PropertiesService/LockService.
- `src/app.generated.js`: SINH TỰ ĐỘNG từ `../apps-script/*.gs` bằng `node build.cjs` (đừng sửa tay).
- `src/bridge.js`: gọi về Apps Script cũ (Drive, email, danh sách lớp, sao lưu).
- `schema.sql`: 3 bảng D1. `migrate.cjs`: chuyển JSON xuất từ Sheet thành SQL.
- `dist/index.js`: bản gộp 1 file (sinh bởi `wrangler deploy --dry-run --outdir dist`) để dán vào bảng điều khiển nếu không dùng wrangler.

Lệnh (trong thư mục này, cần Node 18+):
    npm install
    npx wrangler login
    npx wrangler d1 create vatlycocuong            → dán database_id vào wrangler.toml
    npx wrangler d1 execute vatlycocuong --remote --file=schema.sql
    npx wrangler d1 execute vatlycocuong --remote --file=../du-lieu/migrate.sql
    npx wrangler secret put BRIDGE_SECRET          → dán khóa từ hàm taoKhoaCauNoi() trong Apps Script
    (sửa APPS_SCRIPT_URL trong wrangler.toml = URL web app Apps Script hiện tại)
    npm run deploy                                 → in ra địa chỉ https://vatlycocuong-api.<tên>.workers.dev
