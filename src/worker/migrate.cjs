// Chuyển dữ liệu đã xuất từ Google Sheet (JSON) thành file SQL nạp vào D1.
// Dùng: node migrate.js <thư mục export> <file sql ra>
const fs = require('fs'), path = require('path');
const [dir, out] = process.argv.slice(2);
const SYS_TABS = ['NguoiDung', 'Nhom', 'Phien', 'BaiKiemTra', 'DapAn', 'KetQua', 'NhatKySua', 'GopY', 'NhatKyLoi', 'BuoiHoc', 'DiemDanhLog', 'YeuCauSua', 'ChoXuLy', 'CaiDat'];
const q = s => "'" + String(s).replace(/'/g, "''") + "'";
const lines = ["DELETE FROM rows; DELETE FROM sheets;"];
let total = 0;
function addTab(ss, name, pos, values) {
  // bỏ các dòng trống ở cuối (Sheet xuất kèm), giữ dòng trống ở giữa để không lệch số dòng
  let n = values.length; while (n > 0 && (values[n - 1] || []).every(v => v === '' || v === null || v === undefined)) n--;
  values = values.slice(0, n).map(r => (r || []).map(v => (v === null || v === undefined ? '' : v)));
  lines.push(`INSERT INTO sheets(ss,tab,pos) VALUES (${q(ss)},${q(name)},${pos});`);
  values.forEach((r, i) => lines.push(`INSERT INTO rows(ss,tab,rn,data) VALUES (${q(ss)},${q(name)},${i + 1},${q(JSON.stringify(r))});`));
  total += values.length;
  return values.length;
}
const report = [];
SYS_TABS.forEach((t, i) => {
  const f = path.join(dir, t + '.json');
  const v = fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, 'utf8')).values || [] : [];
  if (t === 'Phien') { report.push([t, addTab('SYS', t, i, v.slice(0, 1)) - 1, '(phiên cũ không chuyển)']); return; }
  report.push([t, addTab('SYS', t, i, v) - 1]);
});
const ro = JSON.parse(fs.readFileSync(path.join(dir, 'roster.json'), 'utf8'));
ro.tabs.forEach((t, i) => report.push(['DS ' + t.title, addTab(ro.spreadsheetId, t.title, i, t.values) - 2]));
lines.push(`INSERT OR REPLACE INTO kv(k,v) VALUES ('SYSTEM_SHEET_ID','SYS');`);
lines.push(`INSERT OR IGNORE INTO kv(k,v) VALUES ('__lock','');`);
lines.push(`INSERT OR REPLACE INTO kv(k,v) VALUES ('OWNER_EMAIL','dinhmanh3802@gmail.com');`);
lines.push(`INSERT OR REPLACE INTO kv(k,v) VALUES ('DA_CHUYEN_LUC',${q(new Date().toISOString())});`);
fs.writeFileSync(out, lines.join('\n') + '\n');
console.log(report.map(r => r.join(': ')).join(' | '));
console.log('Tổng', total, 'dòng →', out);
