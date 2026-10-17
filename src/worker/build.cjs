// Gộp toàn bộ apps-script/*.gs thành một module: createApp(G) → { handle, fn }.
// Mọi khai báo cấp cao nhất của code cũ nằm trong phạm vi hàm, nên mỗi lượt gọi là một trạng thái mới tinh (giống Apps Script).
const fs = require('fs'), path = require('path');
const SRC = path.join(__dirname, '..', 'apps-script'), OUT = path.join(__dirname, 'src', 'app.generated.js');
const ORDER = ['Config.gs', 'Util.gs', 'Setup.gs', 'Score.gs', 'Drive.gs', 'Auth.gs', 'Students.gs', 'Bai.gs', 'Diem.gs', 'DiemNgoai.gs', 'Public.gs', 'ThongKe.gs', 'DiemDanh.gs', 'Anh.gs', 'CaiDat.gs', 'Bridge.gs', 'Code.gs'];
const files = fs.readdirSync(SRC).filter(f => f.endsWith('.gs'));
const missing = files.filter(f => !ORDER.includes(f));
if (missing.length) throw new Error('Chưa xếp thứ tự cho: ' + missing.join(', '));
const GLOBALS = ['SpreadsheetApp', 'PropertiesService', 'CacheService', 'LockService', 'Utilities', 'ContentService', 'DriveApp', 'MailApp', 'GmailApp', 'ScriptApp', 'Session', 'UrlFetchApp', 'HtmlService', 'Logger'];
const EXPORT = ['handle', 'route', 'currentUser', 'requireAuth', 'requireStaff', 'findBai', 'requireBaiAccess', 'sysRows', 'sysTouch', 'getSetting', 'getBai', 'parseJsonSafe', 'safeName', 'ngayStr', 'parseNhomList', 'donDepHangNgay', 'cleanupOldImages', 'pubTouch', 'dataTouch', 'ensureSchema', 'setTextCell', 'idx', 'nowIso', 'getSystemSheet', 'getNhomMap', 'rosterTouch', 'listGroups', 'syncTeacher', 'logSua'];
let body = '';
for (const f of ORDER) { if (!files.includes(f)) continue; body += `\n/* ===================== ${f} ===================== */\n` + fs.readFileSync(path.join(SRC, f), 'utf8') + '\n'; }
const out = `/* FILE SINH TỰ ĐỘNG bởi build.js từ apps-script/*.gs. Đừng sửa tay. */
export function createApp(G) {
  const { ${GLOBALS.join(', ')} } = G;
  const WORKER_MODE = !!G.WORKER_MODE;
  const console = G.console;
${body}
  const _fn = {};
  ${EXPORT.map(n => `try { _fn.${n} = ${n}; } catch (e) {}`).join('\n  ')}
  return _fn;
}
`;
fs.writeFileSync(OUT, out);
console.log('Đã sinh', OUT, Math.round(out.length / 1024) + ' KB,', ORDER.filter(f => files.includes(f)).length, 'file .gs');
