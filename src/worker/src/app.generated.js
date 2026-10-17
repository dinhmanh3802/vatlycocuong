/* FILE SINH TỰ ĐỘNG bởi build.js từ apps-script/*.gs. Đừng sửa tay. */
export function createApp(G) {
  const { SpreadsheetApp, PropertiesService, CacheService, LockService, Utilities, ContentService, DriveApp, MailApp, GmailApp, ScriptApp, Session, UrlFetchApp, HtmlService, Logger } = G;
  const WORKER_MODE = !!G.WORKER_MODE;
  const console = G.console;

/* ===================== Config.gs ===================== */
/**
 * CẤU HÌNH CHUNG
 * ROSTER_SHEET_ID chỉ là file danh sách của giáo viên ĐẦU TIÊN (mẹ/cô Cường).
 * Mỗi giáo viên về sau có file danh sách riêng, lưu trong bảng NguoiDung.
 */

const ROSTER_SHEET_ID = '11S9VfB_83K74Bv9m36QCF8FHTvqXg0s8A_tZWXRG48c';

const DEFAULT_TEMP_PW = 'Matkhau@123';
const SESSION_TTL_HOURS = 12;
const APP_VERSION = '7.0.0';
/** Địa chỉ máy chủ mới (Cloudflare Worker) — dùng để xác nhận lời gọi cầu nối. Để trống nếu chưa chuyển. */
var WORKER_URL = '';
/* Đang chạy trên Cloudflare Worker (lớp tương thích đặt G.WORKER_MODE = true) hay trên Apps Script */
function isWorker() { try { return typeof WORKER_MODE !== 'undefined' && !!WORKER_MODE; } catch (e) { return false; } }
const TEN_HE_THONG = 'Vật Lý Cô Lâm Cường';

function _props() { return PropertiesService.getScriptProperties(); }

/*
 * Mở một file Google Sheet tốn khoảng 0,2–0,5 giây. Trong MỘT lượt gọi, mỗi file chỉ mở MỘT lần
 * (trước đây một thao tác có thể mở lại cùng file 5–10 lần).
 */
var _SS_MEMO = {};
var _SYS_ID = null;

function getSystemSheetId() {
  if (_SYS_ID) return _SYS_ID;
  var id = _props().getProperty('SYSTEM_SHEET_ID');
  if (!id) throw new Error('Chưa chạy setup(). Hãy chạy hàm setup() một lần trong trình soạn Apps Script.');
  _SYS_ID = id;
  return id;
}
function openSpreadsheetMemo(id) {
  if (!_SS_MEMO[id]) _SS_MEMO[id] = SpreadsheetApp.openById(id);
  return _SS_MEMO[id];
}
function getSystemSheet() { return openSpreadsheetMemo(getSystemSheetId()); }

/** Mở file danh sách của một giáo viên theo ID */
function openRoster(rosterSheetId) {
  if (!rosterSheetId) throw new Error('Giáo viên chưa được gán file danh sách');
  return openSpreadsheetMemo(rosterSheetId);
}


/* ===================== Util.gs ===================== */
/** TIỆN ÍCH CHUNG */

function uuid() { return Utilities.getUuid(); }
function nowIso() { return new Date().toISOString(); }

function sha256Hex(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function hashPassword(pw, salt) {
  var h = salt + '|' + pw;
  for (var i = 0; i < 1000; i++) h = sha256Hex(h);
  return h;
}
function randomSalt() { return uuid().replace(/-/g, ''); }

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function idx(headerRow) {
  var m = {};
  headerRow.forEach(function (h, i) { m[String(h).trim()] = i; });
  return m;
}

function toBool(v) {
  if (v === true) return true;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'x' || s === 'có';
}

/** SBD luôn 6 chữ số, giữ số 0 ở đầu */
function normSbd(v) {
  if (v === null || v === undefined) return '';
  var s = String(v).trim();
  if (s === '') return '';
  if (/^\d+(\.0+)?$/.test(s)) s = String(parseInt(s, 10));
  s = s.replace(/\D/g, '');
  if (s === '') return '';
  while (s.length < 6) s = '0' + s;
  return s;
}

/** Mã nhóm luôn 3 chữ số (091, 104...) */
function normMa(v) {
  if (v === null || v === undefined || String(v).trim() === '') return '';
  var n = parseInt(String(v).replace(/\D/g, ''), 10);
  if (isNaN(n)) return '';
  return ('00' + n).slice(-3);
}
function pad3(n) { return ('00' + parseInt(n, 10)).slice(-3); }

function setTextCell(sheet, row, col1, val) {
  var rg = sheet.getRange(row, col1);
  rg.setNumberFormat('@');
  rg.setValue(val);
  sysTouch();
}

/*
 * Đọc một tab của Sheet hệ thống thành mảng object (kèm _row là dòng tuyệt đối).
 * Trong một lượt gọi, mỗi tab chỉ đọc một lần (sysTouch() xóa bộ nhớ này sau khi ghi).
 */
var _ROWS_MEMO = {};
function sysTouch(tab) { if (tab) delete _ROWS_MEMO[tab]; else _ROWS_MEMO = {}; }
function sysRows(tab) {
  if (_ROWS_MEMO[tab]) return _ROWS_MEMO[tab].slice();
  var sh = getSystemSheet().getSheetByName(tab);
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  if (!vals.length) return [];
  var H = idx(vals[0]);
  var keys = Object.keys(H);
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    var o = {};
    keys.forEach(function (k) { o[k] = vals[i][H[k]]; });
    o._row = i + 1;
    out.push(o);
  }
  _ROWS_MEMO[tab] = out;
  return out.slice();
}

/*
 * Bộ nhớ đệm dùng chung giữa các lượt gọi (CacheService, tối đa 6 giờ).
 * Mỗi mục bị giới hạn 100 KB nên dữ liệu lớn được cắt thành nhiều mảnh.
 */
var CACHE_CHUNK = 30000;   // ký tự; chữ Việt có dấu chiếm tới 3 byte nên để dư
function cachePutBig(key, obj, ttl) {
  var s = JSON.stringify(obj), c = CacheService.getScriptCache();
  var n = Math.max(1, Math.ceil(s.length / CACHE_CHUNK)), m = {};
  for (var i = 0; i < n; i++) m[key + '_' + i] = s.substr(i * CACHE_CHUNK, CACHE_CHUNK);
  m[key + '_n'] = String(n);
  try { c.putAll(m, ttl || 600); } catch (e) { /* quá lớn hoặc cache bận: bỏ qua, lần sau đọc thẳng */ }
}
function cacheGetBig(key) {
  var c = CacheService.getScriptCache(), n = Number(c.get(key + '_n'));
  if (!n) return null;
  var ks = [];
  for (var i = 0; i < n; i++) ks.push(key + '_' + i);
  var got = c.getAll(ks), s = '';
  for (var j = 0; j < n; j++) { if (got[ks[j]] === undefined || got[ks[j]] === null) return null; s += got[ks[j]]; }
  try { return JSON.parse(s); } catch (e) { return null; }
}
/* ---------- phiên bản dữ liệu: mọi thao tác GHI đổi phiên bản, nên kết quả nhớ tạm không bao giờ cũ ---------- */
var _DVER = null;
var _KHONG_DOI = false;      // lệnh ghi nhưng lần này không đổi gì (vd mở buổi đã có sẵn)
function dataVer() {
  if (_DVER) return _DVER;
  var c = CacheService.getScriptCache(), v = c.get('DVER');
  if (!v) { v = Date.now().toString(36); c.put('DVER', v, 21600); }
  return (_DVER = v);
}
function dataTouch() { _DVER = Date.now().toString(36) + Math.random().toString(36).slice(2, 5); CacheService.getScriptCache().put('DVER', _DVER, 21600); }
/** Trả kết quả đã tính nếu dữ liệu chưa đổi; p.fresh = true thì tính lại */
function cachedResult(key, ttl, fn, fresh) {
  var k = 'cr_' + key + '_' + dataVer();
  if (!fresh) { var hit = cacheGetBig(k); if (hit) return hit; }
  var res = fn();
  try { cachePutBig(k, res, ttl); } catch (e) {}
  return res;
}
function cacheDel(key) { try { CacheService.getScriptCache().remove(key + '_n'); } catch (e) {} }

/**
 * Danh sách mã nhóm của một bài, ví dụ ['103','104'].
 * Mới lưu dạng chữ "103,104"; nhưng dữ liệu cũ có thể đã bị Google Sheets tự đổi thành số
 * (103,104 -> 103104, hoặc 103.104 nếu Sheets đặt vùng Việt Nam) nên đọc có sửa lại.
 */
function parseNhomList(v) {
  if (v === null || v === undefined || v === '') return [];
  var raw = [], i;
  if (typeof v === 'number') {
    var s = String(v);
    if (s.indexOf('.') !== -1) {                       // 103.104: phần thập phân bị cắt số 0 cuối, đệm lại cho đủ 3 chữ số
      s.split('.').forEach(function (part, k) { raw.push(k === 0 ? part : (part + '00').slice(0, 3)); });
    } else {                                            // 103104: mất số 0 đầu của nhóm đầu tiên, đệm lại
      while (s.length % 3) s = '0' + s;
      for (i = 0; i < s.length; i += 3) raw.push(s.substr(i, 3));
    }
  } else {
    String(v).split(/[,;\s]+/).forEach(function (x) { if (x) raw.push(x); });
  }
  var seen = {}, out = [];
  raw.forEach(function (x) { var m = normMa(x); if (m && !seen[m]) { seen[m] = true; out.push(m); } });
  return out;
}


/* ===================== Setup.gs ===================== */
/** THIẾT LẬP BAN ĐẦU — chạy hàm setup() một lần (chạy lại được, an toàn) */

var SYS_TABS = {
  NguoiDung:  ['username', 'role', 'fullName', 'passHash', 'salt', 'mustChange', 'active', 'rosterSheetId', 'rangeStart', 'rangeEnd', 'createdAt', 'email', 'maNhom', 'quanLy'],
  Nhom:       ['maNhom', 'tenNhom', 'sheetName', 'teacher', 'rosterSheetId', 'total', 'active'],
  Phien:      ['token', 'username', 'role', 'fullName', 'expiry'],
  BaiKiemTra: ['maBai', 'tenBai', 'ngay', 'teacher', 'cacNhom', 'cauHinh', 'trangThai', 'driveFolderId', 'ngayCongBo', 'createdAt', 'luuTru', 'soAnh', 'dapAnFileId'],
  DapAn:      ['maBai', 'maDe', 'dapAn'],
  KetQua:     ['maBai', 'sbd', 'maDe', 'tong', 'p1', 'p2', 'p3', 'vang', 'traLoi', 'capNhat', 'daSua', 'nguon'],
  NhatKySua:  ['thoiGian', 'nguoiSua', 'maBai', 'sbd', 'truong', 'giaTriCu', 'giaTriMoi', 'lyDo'],
  GopY:       ['thoiGian', 'noiDung', 'lienHe', 'daXuLy', 'id', 'nguon', 'nguoiGui', 'sbd'],
  NhatKyLoi:  ['thoiGian', 'action', 'user', 'message'],
  BuoiHoc:    ['maBuoi', 'maNhom', 'ngay', 'trangThai', 'dd', 'ghiChu', 'nguoiGhi', 'capNhat', 'nopBoi', 'nopLuc', 'chotBoi', 'chotLuc', 'taoBoi', 'taoLuc'],
  DiemDanhLog:['thoiGian', 'maBuoi', 'sbd', 'cu', 'moi', 'nguoi', 'loai', 'lyDo'],
  YeuCauSua:  ['id', 'thoiGian', 'maBuoi', 'maNhom', 'sbd', 'cu', 'moi', 'lyDo', 'nguoiGui', 'lienHe', 'trangThai', 'duyetBoi', 'duyetLuc', 'phanHoi'],
  ChoXuLy:    ['id', 'maBai', 'sbd', 'maDe', 'tong', 'p1', 'p2', 'p3', 'traLoi', 'lyDo', 'ghiChu', 'tenFile', 'capNhat'],
  CaiDat:     ['key', 'value']
};

function setup() {
  var props = _props();
  var id = props.getProperty('SYSTEM_SHEET_ID');
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('HỆ THỐNG ĐIỂM - DỮ LIỆU (đừng chỉnh tay)');
    id = ss.getId();
    props.setProperty('SYSTEM_SHEET_ID', id);
  }
  props.setProperty('ROSTER_SHEET_ID', ROSTER_SHEET_ID);

  // Tạo các tab nếu chưa có
  Object.keys(SYS_TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(SYS_TABS[name]);
  });
  var def = ss.getSheetByName('Sheet1');
  if (def) { try { ss.deleteSheet(def); } catch (e) {} }

  // Thêm các cột mới vào tab cũ (nâng cấp bản nền cũ, không mất dữ liệu)
  ensureSysColumns(ss);

  // Cài đặt mặc định
  seedSetting(ss, 'hienThiThuHang', 'false');
  seedSetting(ss, 'hienDiemTB', 'false');
  seedSetting(ss, 'soNgayGiuAnh', '30');
  seedSetting(ss, 'tenHienThi', TEN_HE_THONG);
  Object.keys(CAI_DAT_MAC_DINH).forEach(function (k) { if (k !== 'emailNhanThongBao') seedSetting(ss, k, CAI_DAT_MAC_DINH[k]); });
  seedSetting(ss, 'emailNhanThongBao', (Session.getEffectiveUser().getEmail() || ''));

  // Tài khoản mặc định: admin (không dải), giáo viên 'me' giữ dải 000–199
  seedUser(ss, 'admin', 'admin', 'Quản trị', '', '', '');
  seedUser(ss, 'me', 'teacher', 'Cô Lâm Cường', ROSTER_SHEET_ID, 0, 199);

  // Nếu 'me' đã tồn tại từ bản cũ mà thiếu file/dải thì bổ sung
  var me = findUserRow('me');
  if (me) {
    if (!me.data[me.H['rosterSheetId']]) me.sheet.getRange(me.row, me.H['rosterSheetId'] + 1).setValue(ROSTER_SHEET_ID);
    if (me.data[me.H['rangeStart']] === '' || me.data[me.H['rangeStart']] === undefined) me.sheet.getRange(me.row, me.H['rangeStart'] + 1).setValue(0);
    if (me.data[me.H['rangeEnd']] === '' || me.data[me.H['rangeEnd']] === undefined) me.sheet.getRange(me.row, me.H['rangeEnd'] + 1).setValue(199);
  }

  // Thêm cột TRẠNG THÁI / SBD CŨ vào file danh sách và đăng ký nhóm của 'me'
  var report = syncTeacher('me');

  Logger.log('XONG. Sheet hệ thống: ' + ss.getUrl());
  Logger.log('Đăng nhập admin:     admin / ' + DEFAULT_TEMP_PW);
  Logger.log('Đăng nhập giáo viên: me / ' + DEFAULT_TEMP_PW);
  Logger.log('Đã đăng ký ' + report.registered.length + ' nhóm cho cô Cường (dải 000–199).');
  if (report.conflicts.length) Logger.log('Cảnh báo: ' + JSON.stringify(report.conflicts));
  return ss.getUrl();
}

/** Thêm các cột thiếu vào từng tab hệ thống theo SYS_TABS */
function ensureSysColumns(ss) {
  Object.keys(SYS_TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (c) { return String(c).trim(); });
    SYS_TABS[name].forEach(function (col) {
      if (header.indexOf(col) === -1) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(col);
        header.push(col);
      }
    });
  });
}

function seedSetting(ss, key, value) {
  var sh = ss.getSheetByName('CaiDat');
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (String(rows[i][0]) === key) return;
  sh.appendRow([key, value]);
}

function seedUser(ss, username, role, fullName, rosterId, rangeStart, rangeEnd) {
  var sh = ss.getSheetByName('NguoiDung');
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (String(rows[i][0]) === username) return;
  var salt = randomSalt();
  var hash = hashPassword(DEFAULT_TEMP_PW, salt);
  // Thứ tự khớp SYS_TABS.NguoiDung
  sh.appendRow([username, role, fullName, hash, salt, true, true, rosterId || '',
    (rangeStart === '' ? '' : rangeStart), (rangeEnd === '' ? '' : rangeEnd), nowIso()]);
}

/** Thêm header TRẠNG THÁI / SBD CŨ cho một tab danh sách nếu thiếu */
function ensureRosterColumns(sheet) {
  var values = sheet.getDataRange().getValues();
  var headerRow = -1, headers = [];
  for (var r = 0; r < Math.min(values.length, 6); r++) {
    var up = values[r].map(function (c) { return String(c).trim().toUpperCase(); });
    if (up.indexOf('SBD') !== -1) { headerRow = r; headers = up; break; }
  }
  if (headerRow === -1) return false;
  var nextCol = sheet.getLastColumn() + 1;
  ['TRẠNG THÁI', 'SBD CŨ'].forEach(function (h) {
    if (headers.indexOf(h) === -1) { sheet.getRange(headerRow + 1, nextCol).setValue(h); nextCol++; }
  });
  return true;
}


/* ===================== Score.gs ===================== */
/** CHẤM ĐIỂM — dùng khi sửa câu trả lời của một em, và để đối chiếu với điểm TNMaker */

function round2(x) { return Math.round(x * 100 + 1e-9) / 100; }

function normP3(s) {
  return String(s == null ? '' : s).trim().replace(/\./g, ',').replace(/_+$/, '');
}
function normP2(s) {
  var t = String(s == null ? '' : s), out = '';
  for (var i = 0; i < 4; i++) {
    var c = (t.charAt(i) || '_').toUpperCase();
    out += (c === 'Đ' || c === 'S') ? c : '_';
  }
  return out;
}

/**
 * cfg: cấu hình bài {p1:{on,soCau,diem}, p2:{on,soCau,d1..d4}, p3:{on,soCau,diem}}
 * tl:  câu trả lời của em {p1:['A',...], p2:['ĐSSĐ',...], p3:['8','0,31',...]}
 * key: đáp án đúng cùng dạng
 * Trả về điểm từng phần, tổng, và mảng đúng/sai từng câu (null nếu thiếu đáp án).
 */
function chamBai(cfg, tl, key) {
  tl = tl || {}; key = key || {};
  var out = { p1: 0, p2: 0, p3: 0, tong: 0, thieuKey: false, dung: { p1: [], p2: [], p3: [] } };

  if (cfg.p1 && cfg.p1.on) {
    var n1 = Number(cfg.p1.soCau) || 0, per1 = n1 ? Number(cfg.p1.diem) / n1 : 0;
    for (var i = 0; i < n1; i++) {
      var k1 = key.p1 && key.p1[i] ? String(key.p1[i]).trim().toUpperCase() : '';
      if (!k1) { out.thieuKey = true; out.dung.p1.push(null); continue; }
      var a1 = tl.p1 && tl.p1[i] ? String(tl.p1[i]).trim().toUpperCase() : '';
      var ok1 = a1 === k1;
      out.dung.p1.push(ok1 ? 1 : 0);
      if (ok1) out.p1 += per1;
    }
  }

  if (cfg.p2 && cfg.p2.on) {
    var n2 = Number(cfg.p2.soCau) || 0;
    var pts = [0, Number(cfg.p2.d1) || 0, Number(cfg.p2.d2) || 0, Number(cfg.p2.d3) || 0, Number(cfg.p2.d4) || 0];
    for (var j = 0; j < n2; j++) {
      var k2 = key.p2 && key.p2[j] ? normP2(key.p2[j]) : '';
      if (!k2 || k2.indexOf('_') !== -1) { out.thieuKey = true; out.dung.p2.push(null); continue; }
      var a2 = normP2(tl.p2 && tl.p2[j]);
      var cnt = 0;
      for (var m = 0; m < 4; m++) if (a2.charAt(m) === k2.charAt(m)) cnt++;
      out.dung.p2.push(cnt);
      out.p2 += pts[cnt];
    }
  }

  if (cfg.p3 && cfg.p3.on) {
    var n3 = Number(cfg.p3.soCau) || 0, per3 = n3 ? Number(cfg.p3.diem) / n3 : 0;
    for (var q = 0; q < n3; q++) {
      var k3 = key.p3 ? normP3(key.p3[q]) : '';
      if (!k3) { out.thieuKey = true; out.dung.p3.push(null); continue; }
      var a3 = normP3(tl.p3 && tl.p3[q]);
      var ok3 = a3 !== '' && a3 === k3;
      out.dung.p3.push(ok3 ? 1 : 0);
      if (ok3) out.p3 += per3;
    }
  }

  out.p1 = round2(out.p1); out.p2 = round2(out.p2); out.p3 = round2(out.p3);
  out.tong = round2(out.p1 + out.p2 + out.p3);
  return out;
}


/* ===================== Drive.gs ===================== */
/** TỔ CHỨC THƯ MỤC DRIVE */

function getRootFolder() {
  var p = _props();
  var id = p.getProperty('DRIVE_ROOT_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('HỆ THỐNG ĐIỂM - Lớp học thêm');
  p.setProperty('DRIVE_ROOT_ID', f.getId());
  return f;
}

function getTeacherFolder(username) {
  var p = _props();
  var key = 'DRIVEGV_' + username;
  var id = p.getProperty(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var root = getRootFolder();
  var name = 'GV - ' + username;
  var it = root.getFoldersByName(name);
  var f = it.hasNext() ? it.next() : root.createFolder(name);
  p.setProperty(key, f.getId());
  return f;
}

function safeName(s) {
  return String(s || '').replace(/[\\/:*?"<>|]/g, '-').trim().substring(0, 80) || 'Bài';
}

function createBaiFolder(username, tenBai, ngay) {
  var tf = getTeacherFolder(username);
  var name = (ngay ? ngay + ' - ' : '') + safeName(tenBai);
  return tf.createFolder(name);
}

/** Thư mục con "anh" của một bài (Giai đoạn 3) */
function getAnhFolder(baiFolderId) {
  var folder = DriveApp.getFolderById(baiFolderId);
  var it = folder.getFoldersByName('anh');
  return it.hasNext() ? it.next() : folder.createFolder('anh');
}

/** Xuất một Google Sheet thành blob .xlsx */
function exportSpreadsheetAsXlsxBlob(ssId, filename) {
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx';
  var res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Không xuất được file xlsx (' + res.getResponseCode() + ')');
  return res.getBlob().setName(filename);
}

/** Lưu blob vào thư mục, thay file cùng tên nếu có */
function saveBlobToFolder(folder, blob, name) {
  var it = folder.getFilesByName(name);
  while (it.hasNext()) it.next().setTrashed(true);
  return folder.createFile(blob.setName(name));
}


/* ===================== Auth.gs ===================== */
/** ĐĂNG NHẬP, PHIÊN, TÀI KHOẢN, GIÁO VIÊN */

var LOGIN_MAX_FAIL = 5;         // sai 5 lần liên tiếp thì khóa tạm
var LOGIN_LOCK_MIN = 15;        // trong 15 phút

function _users() { return getSystemSheet().getSheetByName('NguoiDung'); }
function _phien() { return getSystemSheet().getSheetByName('Phien'); }

function findUserRow(username) {
  var sh = _users();
  var rows = sh.getDataRange().getValues();
  var H = idx(rows[0]);
  var want = String(username || '').trim().toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][H['username']]).trim().toLowerCase() === want) return { row: i + 1, data: rows[i], H: H, sheet: sh };
  }
  return null;
}
function userField(u, k) { return u.H[k] === undefined ? '' : u.data[u.H[k]]; }

function authError(msg) { var e = new Error(msg); e.code = 'AUTH'; return e; }

/* ---- Thu hồi phiên: đổi/đặt lại mật khẩu hoặc khóa tài khoản thì mọi phiên cũ hết hiệu lực ngay ---- */
/* Thứ tự "cấp phiên → thu hồi → cấp phiên" luôn tăng dần, kể cả khi rơi cùng một mili-giây:
   thu hồi không sớm hơn phiên gần nhất đã cấp; phiên mới luôn sau lần thu hồi gần nhất. */
function issuedAt(username) {
  var k = 'ISS_' + String(username).toLowerCase(), c = CacheService.getScriptCache(), v = c.get(k);
  if (v === null || v === undefined) { v = _props().getProperty(k) || '0'; c.put(k, v, 21600); }
  return Number(v) || 0;
}
function setIssued(username, t) {
  var k = 'ISS_' + String(username).toLowerCase();
  _props().setProperty(k, String(t)); CacheService.getScriptCache().put(k, String(t), 21600);
}
function revokeSessions(username) {
  var k = 'REV_' + String(username).toLowerCase(), v = String(Math.max(Date.now(), issuedAt(username)));
  _props().setProperty(k, v);
  CacheService.getScriptCache().put(k, v, 21600);
}
function revokedAt(username) {
  var k = 'REV_' + String(username).toLowerCase(), c = CacheService.getScriptCache(), v = c.get(k);
  if (v === null || v === undefined) { v = _props().getProperty(k) || '0'; c.put(k, v, 21600); }
  return Number(v) || 0;
}

function apiLogin(p) {
  var username = String(p.username || '').trim();
  var pw = String(p.password || '');
  if (!username || !pw) throw new Error('Nhập tài khoản và mật khẩu');
  var cache = CacheService.getScriptCache(), fk = 'lf_' + username.toLowerCase();
  var fails = Number(cache.get(fk)) || 0;
  if (fails >= LOGIN_MAX_FAIL) throw new Error('Đăng nhập sai quá ' + LOGIN_MAX_FAIL + ' lần. Thử lại sau ' + LOGIN_LOCK_MIN + ' phút, hoặc dùng "Quên mật khẩu".');
  var u = findUserRow(username);
  var dungMk = u && hashPassword(pw, userField(u, 'salt')) === userField(u, 'passHash');
  if (!dungMk) {
    cache.put(fk, String(fails + 1), LOGIN_LOCK_MIN * 60);
    var con = LOGIN_MAX_FAIL - fails - 1;
    throw new Error('Sai tài khoản hoặc mật khẩu' + (con > 0 && con <= 2 ? ' (còn ' + con + ' lần thử)' : ''));
  }
  if (!toBool(userField(u, 'active'))) throw new Error('Tài khoản đã bị khóa. Liên hệ quản trị.');
  cache.remove(fk);
  var uname = String(userField(u, 'username')), role = userField(u, 'role'), fullName = userField(u, 'fullName');
  var token = createSession(uname, role, fullName);
  return { token: token, user: { username: uname, role: role, fullName: fullName, mustChange: toBool(userField(u, 'mustChange')) } };
}

function createSession(username, role, fullName) {
  var token = uuid() + uuid().replace(/-/g, '');
  var now = Math.max(Date.now(), revokedAt(username) + 1);     // phiên mới luôn sau lần thu hồi gần nhất, kể cả cùng mili-giây
  setIssued(username, now);
  var exp = new Date(now + SESSION_TTL_HOURS * 3600 * 1000).toISOString();
  var obj = { token: token, username: username, role: role, fullName: fullName, expiry: exp, iat: now };
  _phien().appendRow([token, username, role, fullName, exp]);
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(obj), 6 * 3600);
  return token;
}

function getSession(token) {
  if (!token) return null;
  var cached = CacheService.getScriptCache().get('sess_' + token);
  var s = null;
  if (cached) { s = JSON.parse(cached); }
  else {
    var rows = _phien().getDataRange().getValues();
    var H = idx(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][H['token']]) === token) {
        s = { token: token, username: rows[i][H['username']], role: rows[i][H['role']], fullName: rows[i][H['fullName']], expiry: rows[i][H['expiry']], iat: new Date(rows[i][H['expiry']]).getTime() - SESSION_TTL_HOURS * 3600 * 1000 };
        break;
      }
    }
  }
  if (!s) return null;
  if (new Date(s.expiry).getTime() < Date.now()) return null;
  var rev = revokedAt(s.username);
  if (rev && (s.iat || 0) <= rev) return null;               // tạo trước hoặc đúng lúc thu hồi: hết hiệu lực
  return s;
}

function currentUser(p) {
  var s = getSession(p.token);
  if (!s) throw authError('Phiên đã hết hạn, hãy đăng nhập lại');
  return { username: s.username, role: s.role, fullName: s.fullName };
}
/** Dùng cho chức năng vừa công khai vừa có thể đăng nhập (góp ý): không ném lỗi */
function optionalUser(p) { try { return p.token ? currentUser(p) : null; } catch (e) { return null; } }

function apiLogout(p) {
  if (p.token) {
    CacheService.getScriptCache().remove('sess_' + p.token);
    var sh = _phien();
    var rows = sh.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) if (String(rows[i][0]) === p.token) sh.deleteRow(i + 1);
  }
  return { ok: true };
}

function setPassword(u, newPw, mustChange) {
  var salt = randomSalt();
  u.sheet.getRange(u.row, u.H['salt'] + 1).setValue(salt);
  u.sheet.getRange(u.row, u.H['passHash'] + 1).setValue(hashPassword(newPw, salt));
  u.sheet.getRange(u.row, u.H['mustChange'] + 1).setValue(!!mustChange);
}
function checkNewPassword(pw) {
  pw = String(pw || '');
  if (pw.length < 6) throw new Error('Mật khẩu mới cần ít nhất 6 ký tự');
  if (pw.length > 100) throw new Error('Mật khẩu quá dài');
  return pw;
}

/** Đổi mật khẩu: các máy khác bị đăng xuất, máy đang dùng nhận phiên mới */
function apiChangePassword(p) {
  var u0 = currentUser(p);
  var u = findUserRow(u0.username);
  if (!u) throw new Error('Không tìm thấy tài khoản');
  if (hashPassword(String(p.oldPassword || ''), userField(u, 'salt')) !== userField(u, 'passHash')) throw new Error('Mật khẩu hiện tại không đúng');
  var pw = checkNewPassword(p.newPassword);
  if (pw === String(p.oldPassword || '')) throw new Error('Mật khẩu mới phải khác mật khẩu cũ');
  setPassword(u, pw, false);
  revokeSessions(u0.username);
  Utilities.sleep(5);
  var token = createSession(u0.username, u0.role, u0.fullName);
  return { ok: true, token: token };
}

/* ---- Quên mật khẩu: gửi mã 6 số tới email của tài khoản (quản trị nhập email cho từng tài khoản) ---- */
function quenMatKhau(p) {
  var username = String(p.username || '').trim();
  if (!username) throw new Error('Nhập tên tài khoản');
  var res = { ok: true, msg: 'Nếu tài khoản có email, mã đặt lại đã được gửi tới email đó. Mã có hiệu lực 15 phút.' };
  var cache = CacheService.getScriptCache(), rk = 'rs_' + username.toLowerCase();
  if (cache.get(rk + '_wait')) return res;               // mỗi phút gửi tối đa một lần
  cache.put(rk + '_wait', '1', 60);
  var u = findUserRow(username);
  if (!u || !toBool(userField(u, 'active'))) return res;  // không tiết lộ tài khoản có tồn tại hay không
  var email = String(userField(u, 'email') || '').trim();
  if (!email) return res;
  var code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put(rk, JSON.stringify({ h: sha256Hex(code), n: 0 }), 900);
  try {
    MailApp.sendEmail({ to: email, name: 'Hệ thống điểm', subject: '[Điểm] Mã đặt lại mật khẩu: ' + code,
      body: 'Mã đặt lại mật khẩu cho tài khoản "' + username + '": ' + code + '\n\nMã có hiệu lực 15 phút. Nếu không phải bạn yêu cầu, hãy bỏ qua email này.' });
  } catch (e) {}
  return res;
}

function datLaiMatKhau(p) {
  var username = String(p.username || '').trim();
  var cache = CacheService.getScriptCache(), rk = 'rs_' + username.toLowerCase();
  var st = cache.get(rk);
  if (!st) throw new Error('Mã đã hết hạn hoặc chưa được gửi. Bấm "Gửi mã" lại.');
  st = JSON.parse(st);
  if (st.n >= 5) { cache.remove(rk); throw new Error('Nhập sai mã quá nhiều lần. Hãy gửi mã mới.'); }
  if (sha256Hex(String(p.code || '').trim()) !== st.h) {
    st.n++; cache.put(rk, JSON.stringify(st), 900);
    throw new Error('Mã không đúng');
  }
  var pw = checkNewPassword(p.newPassword);
  var u = findUserRow(username);
  if (!u) throw new Error('Không tìm thấy tài khoản');
  setPassword(u, pw, false);
  cache.remove(rk);
  cache.remove('lf_' + username.toLowerCase());
  revokeSessions(username);
  return { ok: true };
}

/* ---- Dải mã ---- */

function rangesOverlap(aS, aE, bS, bE) { return aS <= bE && bS <= aE; }

function checkRangeFree(rangeStart, rangeEnd, exceptUsername) {
  if (rangeStart === '' || rangeStart === undefined || rangeStart === null || rangeEnd === '' || rangeEnd === undefined || rangeEnd === null) return;
  var s = parseInt(rangeStart, 10), e = parseInt(rangeEnd, 10);
  if (isNaN(s) || isNaN(e) || s < 0 || e > 999 || s > e) throw new Error('Dải mã không hợp lệ (0–999, đầu ≤ cuối)');
  sysRows('NguoiDung').forEach(function (r) {
    if (r.role !== 'teacher') return;
    if (exceptUsername && String(r.username).toLowerCase() === String(exceptUsername).toLowerCase()) return;
    if (r.rangeStart === '' || r.rangeStart === undefined) return;
    if (rangesOverlap(s, e, parseInt(r.rangeStart, 10), parseInt(r.rangeEnd, 10)))
      throw new Error('Dải mã trùng với giáo viên ' + r.username + ' (' + pad3(r.rangeStart) + '–' + pad3(r.rangeEnd) + ')');
  });
}

/* ---- Quản trị tài khoản (admin) ---- */

function listUsers() {
  sysTouch('NguoiDung');
  return {
    users: sysRows('NguoiDung').map(function (r) {
      return {
        username: r.username, role: r.role, fullName: r.fullName, email: String(r.email || ''),
        active: toBool(r.active), mustChange: toBool(r.mustChange),
        rosterSheetId: r.rosterSheetId, rangeStart: r.rangeStart, rangeEnd: r.rangeEnd
      };
    })
  };
}

function listTeachers() {
  var users = sysRows('NguoiDung').filter(function (r) { return r.role === 'teacher'; });
  var nhom = sysRows('Nhom');
  return {
    teachers: users.map(function (u) {
      return {
        username: u.username, fullName: u.fullName, active: toBool(u.active),
        rosterSheetId: u.rosterSheetId, rangeStart: u.rangeStart, rangeEnd: u.rangeEnd,
        groupCount: nhom.filter(function (n) { return n.teacher === u.username; }).length
      };
    })
  };
}

function checkEmail(email) {
  email = String(email || '').trim();
  if (email && !/^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(email)) throw new Error('Email không hợp lệ');
  return email;
}

function createUser(p) {
  var username = String(p.username || '').trim();
  if (!/^[A-Za-z0-9._-]{2,40}$/.test(username)) throw new Error('Tên tài khoản 2–40 ký tự, chỉ gồm chữ không dấu, số, dấu chấm, gạch');
  if (findUserRow(username)) throw new Error('Tài khoản đã tồn tại');
  var role = p.role === 'admin' ? 'admin' : 'teacher';
  var rangeStart = '', rangeEnd = '';
  if (role === 'teacher') {
    rangeStart = (p.rangeStart === undefined || p.rangeStart === '') ? '' : parseInt(p.rangeStart, 10);
    rangeEnd = (p.rangeEnd === undefined || p.rangeEnd === '') ? '' : parseInt(p.rangeEnd, 10);
    checkRangeFree(rangeStart, rangeEnd, null);
  }
  var email = checkEmail(p.email);
  var sh = _users(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), nCol = sh.getLastColumn();
  var row = []; for (var i = 0; i < nCol; i++) row.push('');
  var salt = randomSalt();
  var set = function (k, v) { if (H[k] !== undefined) row[H[k]] = v; };
  set('username', username); set('role', role); set('fullName', String(p.fullName || '').trim());
  set('passHash', hashPassword(DEFAULT_TEMP_PW, salt)); set('salt', salt); set('mustChange', true); set('active', true);
  set('rosterSheetId', String(p.rosterSheetId || '').trim()); set('rangeStart', rangeStart); set('rangeEnd', rangeEnd);
  set('createdAt', nowIso()); set('email', email);
  sh.appendRow(row);
  sysTouch();
  return { ok: true, tempPassword: DEFAULT_TEMP_PW };
}

/** Sửa thông tin tài khoản: họ tên, email, file danh sách, dải mã */
function setUserInfo(p) {
  var u = findUserRow(p.username);
  if (!u) throw new Error('Không tìm thấy tài khoản');
  var isTeacher = userField(u, 'role') === 'teacher';
  if (isTeacher && p.rangeStart !== undefined && p.rangeEnd !== undefined) checkRangeFree(p.rangeStart, p.rangeEnd, p.username);
  var put = function (k, v) { if (u.H[k] !== undefined) u.sheet.getRange(u.row, u.H[k] + 1).setValue(v); };
  if (p.fullName !== undefined) put('fullName', String(p.fullName || '').trim());
  if (p.email !== undefined) { if (u.H['email'] !== undefined) setTextCell(u.sheet, u.row, u.H['email'] + 1, checkEmail(p.email)); }
  if (isTeacher && p.rosterSheetId !== undefined) put('rosterSheetId', String(p.rosterSheetId || '').trim());
  if (isTeacher && p.rangeStart !== undefined) put('rangeStart', p.rangeStart === '' ? '' : parseInt(p.rangeStart, 10));
  if (isTeacher && p.rangeEnd !== undefined) put('rangeEnd', p.rangeEnd === '' ? '' : parseInt(p.rangeEnd, 10));
  sysTouch();
  return { ok: true };
}
/** Giữ cho tương thích bản cũ */
function setTeacherRoster(p) { return setUserInfo(p); }

function setUserActive(p, me) {
  var u = findUserRow(p.username);
  if (!u) throw new Error('Không tìm thấy tài khoản');
  if (me && String(me.username).toLowerCase() === String(p.username).toLowerCase() && !p.active) throw new Error('Không thể tự khóa tài khoản đang dùng');
  u.sheet.getRange(u.row, u.H['active'] + 1).setValue(!!p.active);
  if (!p.active) revokeSessions(userField(u, 'username'));
  sysTouch();
  return { ok: true };
}

function adminResetPassword(p) {
  var u = findUserRow(p.username);
  if (!u) throw new Error('Không tìm thấy tài khoản');
  setPassword(u, DEFAULT_TEMP_PW, true);
  revokeSessions(userField(u, 'username'));
  CacheService.getScriptCache().remove('lf_' + String(p.username).toLowerCase());
  return { ok: true, tempPassword: DEFAULT_TEMP_PW };
}


/* ===================== Students.gs ===================== */
/** NHÓM VÀ HỌC SINH — mỗi giáo viên một file danh sách riêng */

/** Đọc một tab thành nhóm. Trả null nếu tab không có cột SBD. */
function readGroupFromSheet(sheet) {
  var values = sheet.getDataRange().getValues();
  var headerRow = -1, cols = {};
  for (var r = 0; r < Math.min(values.length, 6); r++) {
    var up = values[r].map(function (c) { return String(c).trim().toUpperCase(); });
    if (up.indexOf('SBD') !== -1) {
      headerRow = r;
      up.forEach(function (h, i) { if (h) cols[h] = i; });
      break;
    }
  }
  if (headerRow === -1) return null;

  var students = [], maNhom = '';
  for (var i = headerRow + 1; i < values.length; i++) {
    var raw = values[i];
    var sbd = normSbd(raw[cols['SBD']]);
    var ho = cols['HỌ'] !== undefined ? String(raw[cols['HỌ']] || '').trim() : '';
    var ten = cols['TÊN'] !== undefined ? String(raw[cols['TÊN']] || '').trim() : '';
    if (!sbd && !ho && !ten) continue;
    if (sbd && !maNhom) maNhom = sbd.substring(0, 3);
    students.push({
      sbd: sbd, ho: ho, ten: ten,
      lop: cols['LỚP'] !== undefined ? String(raw[cols['LỚP']] || '').trim() : '',
      trangThai: (cols['TRẠNG THÁI'] !== undefined ? String(raw[cols['TRẠNG THÁI']] || '').trim() : '') || 'Đang học',
      sbdCu: cols['SBD CŨ'] !== undefined ? normSbd(raw[cols['SBD CŨ']]) : '',
      rowIndex: i + 1
    });
  }
  return { sheetName: sheet.getName(), tenNhom: sheet.getName().trim(), maNhom: maNhom, headerRow: headerRow + 1, cols: cols, students: students };
}

/*
 * Ảnh chụp toàn bộ các tab của một file danh sách. Đọc file danh sách là việc chậm nhất của hệ thống
 * (mở file + đọc từng tab), nên giữ trong bộ nhớ đệm 60 giây. Thao tác ghi qua web xóa bộ nhớ ngay;
 * sửa thẳng trên Google Sheet thì sau tối đa 60 giây (hoặc bấm "Tải lại") web thấy thay đổi.
 * Các bước cần chính xác tuyệt đối (đối soát, ghép dòng chờ) luôn đọc mới (fresh = true).
 */
function rosterSnapshot(fileId, fresh) {
  var key = 'ros_' + fileId;
  if (!fresh) { var hit = cacheGetBig(key); if (hit) return hit; }
  var tabs = [];
  openRoster(fileId).getSheets().forEach(function (sh) {
    var g = readGroupFromSheet(sh);
    if (g) tabs.push(g);
  });
  cachePutBig(key, tabs, 600);     // 10 phút; sửa qua web thì xóa ngay, nút "Tải lại" luôn đọc mới
  return tabs;
}
function rosterTouch(fileId) {
  if (fileId) cacheDel('ros_' + fileId);
  try { pubTouch(); } catch (e) {}          // tên học sinh trên trang tra điểm có thể đã đổi
}
/** Tìm tab của một nhóm trong ảnh chụp: ưu tiên đúng tên tab, không thấy thì dò theo mã nhóm */
function findTab(tabs, sheetName, maNhom) {
  var byName = tabs.filter(function (t) { return String(t.sheetName) === String(sheetName); })[0];
  return byName || tabs.filter(function (t) { return t.maNhom === maNhom; })[0] || null;
}

/** Danh mục nhóm: maNhom -> {tenNhom, teacher, rosterSheetId, sheetName, _row} */
function getNhomMap() {
  var m = {};
  sysRows('Nhom').forEach(function (r) {
    var ma = normMa(r.maNhom);
    if (ma) m[ma] = { tenNhom: r.tenNhom, teacher: r.teacher, rosterSheetId: r.rosterSheetId, sheetName: r.sheetName, _row: r._row };
  });
  return m;
}

function canAccessGroup(u, maNhom) {
  if (u.role === 'admin') return true;
  var info = getNhomMap()[normMa(maNhom)];
  return info && info.teacher === u.username;
}
function requireGroupAccess(u, maNhom) {
  if (!canAccessGroup(u, maNhom)) throw new Error('Không có quyền với nhóm ' + maNhom);
}

/** Tìm nhóm theo mã: dò trong file của giáo viên đã đăng ký */
function getGroupByMaNhom(maNhom) {
  maNhom = normMa(maNhom);
  var reg = getNhomMap()[maNhom];
  if (!reg || !reg.rosterSheetId) return null;
  var file = openRoster(reg.rosterSheetId);
  var sh = file.getSheetByName(reg.sheetName);
  if (!sh) {
    var sheets = file.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var g0 = readGroupFromSheet(sheets[i]);
      if (g0 && g0.maNhom === maNhom) { sh = sheets[i]; break; }
    }
  }
  if (!sh) return null;
  return { sheet: sh, group: readGroupFromSheet(sh), reg: reg };
}

/** Đồng bộ nhóm của một giáo viên: quét file của họ, đăng ký các tab vào sổ chung */
function syncTeacher(username) {
  var user = findUserRow(username);
  if (!user) throw new Error('Không tìm thấy tài khoản ' + username);
  var rosterId = user.data[user.H['rosterSheetId']];
  if (!rosterId) throw new Error('Giáo viên ' + username + ' chưa gán file danh sách');
  var rs = user.data[user.H['rangeStart']] === '' ? 0 : parseInt(user.data[user.H['rangeStart']], 10);
  var re = user.data[user.H['rangeEnd']] === '' ? 999 : parseInt(user.data[user.H['rangeEnd']], 10);

  var ss = getSystemSheet();
  var nhomSheet = ss.getSheetByName('Nhom');
  var Hn = idx(nhomSheet.getRange(1, 1, 1, nhomSheet.getLastColumn()).getValues()[0]);
  var reg = sysRows('Nhom');

  var registered = [], conflicts = [];
  var file = openRoster(rosterId);
  file.getSheets().forEach(function (sh) {
    ensureRosterColumns(sh);
    var g = readGroupFromSheet(sh);
    if (!g || !g.maNhom) return;
    var pfx = parseInt(g.maNhom, 10);
    if (pfx < rs || pfx > re) { conflicts.push({ maNhom: g.maNhom, tenNhom: g.tenNhom, lyDo: 'ngoài dải ' + pad3(rs) + '–' + pad3(re) }); return; }
    var otherOwner = reg.filter(function (r) { return normMa(r.maNhom) === g.maNhom && r.teacher !== username; })[0];
    if (otherOwner) { conflicts.push({ maNhom: g.maNhom, tenNhom: g.tenNhom, lyDo: 'trùng mã với ' + otherOwner.teacher }); return; }

    var total = g.students.length;
    var active = g.students.filter(function (s) { return s.trangThai === 'Đang học'; }).length;
    var existing = reg.filter(function (r) { return normMa(r.maNhom) === g.maNhom && r.teacher === username; })[0];
    if (existing) {
      nhomSheet.getRange(existing._row, Hn['tenNhom'] + 1).setValue(g.tenNhom);
      nhomSheet.getRange(existing._row, Hn['sheetName'] + 1).setValue(g.sheetName);
      nhomSheet.getRange(existing._row, Hn['rosterSheetId'] + 1).setValue(rosterId);
      nhomSheet.getRange(existing._row, Hn['total'] + 1).setValue(total);
      nhomSheet.getRange(existing._row, Hn['active'] + 1).setValue(active);
    } else {
      nhomSheet.appendRow([g.maNhom, g.tenNhom, g.sheetName, username, rosterId, total, active]);
      reg.push({ maNhom: g.maNhom, teacher: username, _row: nhomSheet.getLastRow() });
    }
    registered.push({ maNhom: g.maNhom, tenNhom: g.tenNhom });
  });
  rosterTouch(rosterId); sysTouch();
  return { registered: registered, conflicts: conflicts };
}

function apiSyncGroups(p, u) {
  var target = (u.role === 'admin' && p.username) ? p.username : u.username;
  if (u.role !== 'admin' && target !== u.username) throw new Error('Không đủ quyền');
  return syncTeacher(target);
}

/** Danh sách nhóm người dùng được xem (đếm số liệu sống từ file) */
function listGroups(u, fresh) {
  var reg = sysRows('Nhom').filter(function (r) { return (u.role === 'admin' || r.teacher === u.username) && normMa(r.maNhom); });
  var byFile = {};
  reg.forEach(function (r) { (byFile[r.rosterSheetId] = byFile[r.rosterSheetId] || []).push(r); });

  var out = [];
  Object.keys(byFile).forEach(function (fileId) {
    var tabs = [];
    try { tabs = rosterSnapshot(fileId, !!fresh); } catch (e) {}
    byFile[fileId].forEach(function (r) {
      var ma = normMa(r.maNhom), g = findTab(tabs, r.sheetName, ma);
      out.push({
        maNhom: ma, tenNhom: g ? g.tenNhom : String(r.tenNhom), teacher: r.teacher,
        total: g ? g.students.length : (Number(r.total) || 0),
        active: g ? g.students.filter(function (s) { return s.trangThai === 'Đang học'; }).length : (Number(r.active) || 0)
      });
    });
  });
  out.sort(function (a, b) { return a.maNhom < b.maNhom ? -1 : 1; });
  return { groups: out };
}

function listStudents(p, u) {
  requireGroupAccess(u, p.maNhom);
  var ma = normMa(p.maNhom), reg = getNhomMap()[ma];
  if (!reg) throw new Error('Không tìm thấy nhóm ' + p.maNhom);
  var g = findTab(rosterSnapshot(reg.rosterSheetId, !!p.fresh), reg.sheetName, ma);
  if (!g) throw new Error('Không tìm thấy tab của nhóm ' + ma + ' trong file danh sách');
  return { maNhom: ma, tenNhom: g.tenNhom, students: g.students };
}

function nextSbd(g) {
  var max = 0;
  g.students.forEach(function (s) {
    if (s.sbd && s.sbd.substring(0, 3) === g.maNhom) {
      var n = parseInt(s.sbd.substring(3), 10);
      if (n > max) max = n;
    }
  });
  return g.maNhom + ('00' + (max + 1)).slice(-3);
}

function upsertStudent(p, u) {
  requireGroupAccess(u, p.maNhom);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var found = getGroupByMaNhom(p.maNhom);
    if (!found) throw new Error('Không tìm thấy nhóm ' + p.maNhom);
    ensureRosterColumns(found.sheet);
    found = getGroupByMaNhom(p.maNhom);
    var sh = found.sheet, g = found.group, cols = g.cols;

    var ho = (p.ho || '').trim(), ten = (p.ten || '').trim();
    if (!ho && !ten) throw new Error('Phải nhập họ tên');
    var sbd = normSbd(p.sbd);
    var targetRow = null;
    if (sbd) { var ex = g.students.filter(function (s) { return s.sbd === sbd; })[0]; if (ex) targetRow = ex.rowIndex; }
    var isNew = !targetRow;
    if (isNew) {
      if (sbd && g.students.filter(function (s) { return s.sbd === sbd; }).length) throw new Error('SBD đã tồn tại: ' + sbd);
      if (!sbd) sbd = nextSbd(g);
      targetRow = (g.students.length ? g.students[g.students.length - 1].rowIndex : g.headerRow) + 1;
      if (cols['STT'] !== undefined) sh.getRange(targetRow, cols['STT'] + 1).setValue(g.students.length + 1);
      if (cols['TRẠNG THÁI'] !== undefined) sh.getRange(targetRow, cols['TRẠNG THÁI'] + 1).setValue('Đang học');
    }
    if (cols['HỌ'] !== undefined) sh.getRange(targetRow, cols['HỌ'] + 1).setValue(ho);
    if (cols['TÊN'] !== undefined) sh.getRange(targetRow, cols['TÊN'] + 1).setValue(ten);
    if (p.lop !== undefined && cols['LỚP'] !== undefined) sh.getRange(targetRow, cols['LỚP'] + 1).setValue((p.lop || '').trim());
    if (cols['SBD'] !== undefined) setTextCell(sh, targetRow, cols['SBD'] + 1, sbd);
    rosterTouch(found.reg.rosterSheetId);
    return { ok: true, sbd: sbd, isNew: isNew };
  } finally { lock.releaseLock(); }
}

function setStudentStatus(p, u) {
  requireGroupAccess(u, p.maNhom);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var found = getGroupByMaNhom(p.maNhom);
    if (!found) throw new Error('Không tìm thấy nhóm ' + p.maNhom);
    var sh = found.sheet, g = found.group;
    if (g.cols['TRẠNG THÁI'] === undefined) { ensureRosterColumns(sh); found = getGroupByMaNhom(p.maNhom); sh = found.sheet; g = found.group; }
    var st = g.students.filter(function (s) { return s.sbd === normSbd(p.sbd); })[0];
    if (!st) throw new Error('Không tìm thấy SBD ' + p.sbd);
    sh.getRange(st.rowIndex, g.cols['TRẠNG THÁI'] + 1).setValue(p.trangThai || 'Đang học');
    rosterTouch(found.reg.rosterSheetId);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function transferStudent(p, u) {
  requireGroupAccess(u, p.fromMaNhom);
  requireGroupAccess(u, p.toMaNhom);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var from = getGroupByMaNhom(p.fromMaNhom);
    var to = getGroupByMaNhom(p.toMaNhom);
    if (!from || !to) throw new Error('Không tìm thấy nhóm');
    ensureRosterColumns(from.sheet); ensureRosterColumns(to.sheet);
    from = getGroupByMaNhom(p.fromMaNhom); to = getGroupByMaNhom(p.toMaNhom);

    var oldSbd = normSbd(p.sbd);
    var st = from.group.students.filter(function (s) { return s.sbd === oldSbd; })[0];
    if (!st) throw new Error('Không tìm thấy SBD ' + p.sbd + ' trong nhóm ' + p.fromMaNhom);

    var newSbd = nextSbd(to.group);
    var tsh = to.sheet, tc = to.group.cols;
    var newRow = (to.group.students.length ? to.group.students[to.group.students.length - 1].rowIndex : to.group.headerRow) + 1;
    if (tc['STT'] !== undefined) tsh.getRange(newRow, tc['STT'] + 1).setValue(to.group.students.length + 1);
    if (tc['HỌ'] !== undefined) tsh.getRange(newRow, tc['HỌ'] + 1).setValue(st.ho);
    if (tc['TÊN'] !== undefined) tsh.getRange(newRow, tc['TÊN'] + 1).setValue(st.ten);
    if (tc['LỚP'] !== undefined) tsh.getRange(newRow, tc['LỚP'] + 1).setValue(st.lop);
    if (tc['SBD'] !== undefined) setTextCell(tsh, newRow, tc['SBD'] + 1, newSbd);
    if (tc['TRẠNG THÁI'] !== undefined) tsh.getRange(newRow, tc['TRẠNG THÁI'] + 1).setValue('Đang học');
    if (tc['SBD CŨ'] !== undefined) setTextCell(tsh, newRow, tc['SBD CŨ'] + 1, oldSbd);

    var fc = from.group.cols;
    if (fc['TRẠNG THÁI'] !== undefined) from.sheet.getRange(st.rowIndex, fc['TRẠNG THÁI'] + 1).setValue('Đã chuyển');
    rosterTouch(from.reg.rosterSheetId); rosterTouch(to.reg.rosterSheetId);
    return { ok: true, newSbd: newSbd };
  } finally { lock.releaseLock(); }
}

/** Kiểm tra cấu trúc danh sách người dùng được xem */
function validateRoster(u) {
  var issues = [], allSbd = {};
  var reg = sysRows('Nhom').filter(function (r) { return (u.role === 'admin' || r.teacher === u.username) && normMa(r.maNhom); });
  var byFile = {};
  reg.forEach(function (r) { (byFile[r.rosterSheetId] = byFile[r.rosterSheetId] || []).push(r); });

  Object.keys(byFile).forEach(function (fileId) {
    var tabs;
    try { tabs = rosterSnapshot(fileId, true); } catch (e) { issues.push({ group: '(file)', type: 'Không mở được file', detail: String(fileId) }); return; }
    byFile[fileId].forEach(function (r) {
      var g = findTab(tabs, r.sheetName, normMa(r.maNhom));
      if (!g) { issues.push({ group: String(r.tenNhom), type: 'Thiếu tab', detail: String(r.sheetName) }); return; }
      var seen = {};
      g.students.forEach(function (s) {
        if (s.sbd) {
          (allSbd[s.sbd] = allSbd[s.sbd] || []).push(g.tenNhom);
          if (seen[s.sbd]) issues.push({ group: g.tenNhom, type: 'Trùng SBD trong nhóm', detail: s.sbd });
          seen[s.sbd] = true;
          if (s.sbd.substring(0, 3) !== normMa(r.maNhom)) issues.push({ group: g.tenNhom, type: 'SBD sai mã nhóm', detail: s.sbd + ' (nhóm ' + normMa(r.maNhom) + ')' });
          if (!s.ho && !s.ten) issues.push({ group: g.tenNhom, type: 'Có SBD nhưng thiếu tên', detail: s.sbd });
        } else if (s.ho || s.ten) {
          issues.push({ group: g.tenNhom, type: 'Có tên nhưng thiếu SBD', detail: (s.ho + ' ' + s.ten).trim() });
        }
      });
    });
  });
  Object.keys(allSbd).forEach(function (sbd) { if (allSbd[sbd].length > 1) issues.push({ group: allSbd[sbd].join(', '), type: 'SBD trùng nhiều nhóm', detail: sbd }); });
  return { ok: issues.length === 0, issues: issues };
}


/* ===================== Bai.gs ===================== */
/** BÀI KIỂM TRA, ĐÁP ÁN, SINH FILE */

function _bai() { return getSystemSheet().getSheetByName('BaiKiemTra'); }
function _dapan() { return getSystemSheet().getSheetByName('DapAn'); }

function normMaDe(x) {
  var s = String(x == null ? '' : x).trim();
  if (s === '') return '';
  if (/^\d+$/.test(s)) return ('00' + parseInt(s, 10)).slice(-3);
  return s;
}

/** Sheets tự đổi chuỗi ngày thành kiểu Date; đưa về chuỗi yyyy-MM-dd cho web */
function ngayStr(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function findBai(maBai) {
  return sysRows('BaiKiemTra').filter(function (r) { return String(r.maBai) === String(maBai); })[0] || null;
}
function requireBaiAccess(u, baiRow) {
  if (!baiRow) throw new Error('Không tìm thấy bài');
  if (u.role !== 'admin' && baiRow.teacher !== u.username) throw new Error('Không có quyền với bài này');
}

/** Tổng điểm tối đa theo cấu hình phải bằng 10 */
function validateCauHinh(cfg) {
  var total = 0;
  if (cfg.p1 && cfg.p1.on) {
    if (!(cfg.p1.soCau > 0)) throw new Error('Phần I: số câu phải > 0');
    total += Number(cfg.p1.diem) || 0;
  }
  if (cfg.p2 && cfg.p2.on) {
    if (!(cfg.p2.soCau > 0)) throw new Error('Phần II: số câu phải > 0');
    total += (Number(cfg.p2.soCau) || 0) * (Number(cfg.p2.d4) || 0);
  }
  if (cfg.p3 && cfg.p3.on) {
    if (!(cfg.p3.soCau > 0)) throw new Error('Phần III: số câu phải > 0');
    total += Number(cfg.p3.diem) || 0;
  }
  if (Math.abs(total - 10) > 0.01) throw new Error('Tổng điểm tối đa phải bằng 10 (đang ' + total.toFixed(2) + ')');
}

function createBai(p, u) {
  var cfg = p.cauHinh;
  if (typeof cfg === 'string') cfg = JSON.parse(cfg);
  if (cfg && cfg.loai === 'ngoai') cfg = chuanCfgNgoai(cfg);
  else {
    validateCauHinh(cfg);
    cfg.maDeList = (cfg.maDeList || []).map(normMaDe).filter(function (x) { return x; });
    if (!cfg.maDeList.length) throw new Error('Phải có ít nhất 1 mã đề');
  }

  // Nhóm làm bài phải thuộc giáo viên
  var nhom = getNhomMap();
  var cacNhom = (p.cacNhom || []).map(normMa).filter(function (x) { return x; });
  cacNhom.forEach(function (ma) {
    var info = nhom[ma];
    if (!info) throw new Error('Nhóm ' + ma + ' chưa đăng ký');
    if (u.role !== 'admin' && info.teacher !== u.username) throw new Error('Nhóm ' + ma + ' không thuộc bạn');
  });

  var teacher = (u.role === 'admin' && p.teacher) ? p.teacher : u.username;
  var folder = isWorker() ? { getId: function () { return ''; } } : createBaiFolder(teacher, p.tenBai, p.ngay);   // Worker: thư mục Drive tạo lúc cần, qua cầu nối
  var maBai;                                       // mã không bao giờ trùng, kể cả 2 bài tạo cùng một mili-giây
  do { maBai = 'B' + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 4).toUpperCase(); } while (sysRows('BaiKiemTra').some(function (r) { return String(r.maBai) === maBai; }));
  var sh = _bai(), row = sh.getLastRow() + 1;
  var vals = [maBai, (p.tenBai || '').trim(), p.ngay || '', teacher, cacNhom.join(','), JSON.stringify(cfg), 'Nháp', folder.getId(), '', nowIso()];
  sh.getRange(row, 1, 1, vals.length).setNumberFormat('@');   // lưu dạng chữ: Sheets không tự đổi "103,104", "28/8"… thành số/ngày
  sh.getRange(row, 1, 1, vals.length).setValues([vals]);
  sysTouch();
  return { ok: true, maBai: maBai, driveFolderId: folder.getId() };
}

function updateBai(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  var cfg = p.cauHinh; if (typeof cfg === 'string') cfg = JSON.parse(cfg);
  if (laBaiNgoai(bai) || (cfg && cfg.loai === 'ngoai')) cfg = chuanCfgNgoai(cfg || {});
  else {
    validateCauHinh(cfg);
    cfg.maDeList = (cfg.maDeList || []).map(normMaDe).filter(function (x) { return x; });
    if (!cfg.maDeList.length) throw new Error('Phải có ít nhất 1 mã đề');
  }
  var sh = _bai(); var H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  if (p.tenBai !== undefined) setTextCell(sh, bai._row, H['tenBai'] + 1, (p.tenBai || '').trim());
  if (p.ngay !== undefined) setTextCell(sh, bai._row, H['ngay'] + 1, p.ngay || '');
  if (p.cacNhom !== undefined) setTextCell(sh, bai._row, H['cacNhom'] + 1, (p.cacNhom || []).map(normMa).filter(Boolean).join(','));
  sh.getRange(bai._row, H['cauHinh'] + 1).setValue(JSON.stringify(cfg));
  sysTouch();
  return { ok: true };
}

/** Đếm số dòng theo mã bài trong một tab (đọc một cột) */
function demTheoBai(tab, col1) {
  var sh = getSystemSheet().getSheetByName(tab), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, col1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { var k = String(r[0]); out[k] = (out[k] || 0) + 1; });
  return out;
}

function listBai(u, p) {
  var rows = sysRows('BaiKiemTra').filter(function (r) { return u.role === 'admin' || r.teacher === u.username; });
  rows.sort(function (a, b) {
    var x = ngayStr(a.ngay) || String(a.createdAt), y = ngayStr(b.ngay) || String(b.createdAt);
    return x < y ? 1 : (x > y ? -1 : (String(b.createdAt) < String(a.createdAt) ? -1 : 1));
  });
  var kq = demTheoBai('KetQua', 1), cx = demTheoBai('ChoXuLy', 2);
  return {
    bai: rows.map(function (r) {
      var cfg = parseJsonSafe(r.cauHinh, {});
      return {
        maBai: String(r.maBai), tenBai: String(r.tenBai), ngay: ngayStr(r.ngay), teacher: r.teacher, cacNhom: parseNhomList(r.cacNhom),
        trangThai: r.trangThai, soMaDe: (cfg.maDeList || []).length, luuTru: toBool(r.luuTru),
        soKetQua: kq[String(r.maBai)] || 0, soChoXuLy: cx[String(r.maBai)] || 0, createdAt: String(r.createdAt || '')
      };
    })
  };
}

function getBaiDapAn(maBai) {
  var map = {};
  sysRows('DapAn').filter(function (r) { return String(r.maBai) === String(maBai); }).forEach(function (r) {
    var d = {}; try { d = JSON.parse(r.dapAn || '{}'); } catch (e) {}
    map[normMaDe(r.maDe)] = d;
  });
  return map;
}

function getBai(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  var cfg = {}; try { cfg = JSON.parse(bai.cauHinh || '{}'); } catch (e) {}
  var hasFile = false, fileId = String(bai.dapAnFileId || '');
  if (fileId) hasFile = true;
  else if (!isWorker()) {
    try {
      var folder = DriveApp.getFolderById(bai.driveFolderId);
      var it = folder.getFilesByName('dap-an - ' + safeName(bai.tenBai) + '.xlsx');
      if (it.hasNext()) { hasFile = true; fileId = it.next().getId(); }
    } catch (e) {}
  }
  return {
    maBai: bai.maBai, tenBai: bai.tenBai, ngay: ngayStr(bai.ngay), teacher: bai.teacher,
    cacNhom: parseNhomList(bai.cacNhom),
    cauHinh: cfg, maDeList: cfg.maDeList || [], trangThai: bai.trangThai,
    driveFolderId: bai.driveFolderId,
    dapAn: getBaiDapAn(bai.maBai),
    soKetQua: countKetQua(bai.maBai),
    soChoXuLy: countChoXuLy(bai.maBai),
    luuTru: toBool(bai.luuTru),
    dapAnFile: hasFile ? { fileId: fileId, viewUrl: 'https://drive.google.com/file/d/' + fileId + '/view', downloadUrl: 'https://drive.google.com/uc?export=download&id=' + fileId } : null
  };
}

/**
 * Lưu đáp án (được lưu dở). p.dapAn = {maDe:{p1:[],p2:[],p3:[]}}.
 * Có p.maDe thì chỉ thay đáp án của đúng mã đề đó (web tự lưu từng mã khi đang nhập).
 */
function saveDapAn(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  var incoming = p.dapAn; if (typeof incoming === 'string') incoming = JSON.parse(incoming);
  incoming = incoming || {};
  var only = p.maDe ? normMaDe(p.maDe) : null;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = _dapan();
    var del = sysRows('DapAn').filter(function (r) {
      return String(r.maBai) === String(p.maBai) && (!only || normMaDe(r.maDe) === only);
    }).map(function (r) { return r._row; });
    if (del.length) deleteRowsList(sh, del);
    var rows = Object.keys(incoming).filter(function (md) { return !only || normMaDe(md) === only; })
      .map(function (md) { return [String(p.maBai), normMaDe(md), JSON.stringify(incoming[md])]; });
    if (rows.length) {
      var r0 = sh.getLastRow() + 1;
      sh.getRange(r0, 1, rows.length, 2).setNumberFormat('@');
      sh.getRange(r0, 1, rows.length, 3).setValues(rows);
    }
    sysTouch();
    return { ok: true, luc: nowIso() };
  } finally { lock.releaseLock(); }
}

/** Sinh file đáp án .xlsx đúng mẫu TNMaker, lưu vào Drive, trả link */
function genDapAnFile(p, u) {
  var bai = getBai(p, u);
  var cfg = bai.cauHinh, maDeList = bai.maDeList;
  if (!maDeList.length) throw new Error('Chưa có mã đề');

  var parts = [];
  if (cfg.p1 && cfg.p1.on) parts.push(['p1', cfg.p1.soCau, 'Phần I']);
  if (cfg.p2 && cfg.p2.on) parts.push(['p2', cfg.p2.soCau, 'Phần II']);
  if (cfg.p3 && cfg.p3.on) parts.push(['p3', cfg.p3.soCau, 'Phần III']);

  // Kiểm tra đủ đáp án
  var thieu = [];
  maDeList.forEach(function (md) {
    var da = bai.dapAn[md] || {};
    parts.forEach(function (pt) {
      var arr = da[pt[0]] || [];
      for (var i = 0; i < pt[1]; i++) {
        var v = arr[i];
        if (v === undefined || v === null || String(v).trim() === '') thieu.push('Mã ' + md + ' · ' + pt[2] + ' câu ' + (i + 1));
      }
    });
  });
  if (thieu.length) throw new Error('Còn thiếu đáp án: ' + thieu.slice(0, 8).join('; ') + (thieu.length > 8 ? ' …' : ''));

  // Dựng bảng câu × mã đề
  var header = ['CAU\\MA DE'].concat(maDeList);
  var dataRows = [];
  var cau = 1;
  parts.forEach(function (pt) {
    for (var i = 0; i < pt[1]; i++) {
      var row = [cau];
      maDeList.forEach(function (md) {
        var arr = (bai.dapAn[md] || {})[pt[0]] || [];
        var val = arr[i] !== undefined ? arr[i] : '';
        if (pt[0] === 'p3') val = String(val).replace('.', ',');
        else val = String(val);
        row.push(val);
      });
      dataRows.push(row); cau++;
    }
  });

  var tmp = SpreadsheetApp.create('tmp-dapan-' + bai.maBai);
  try {
    // Tab 1 để trống; tab 2 chứa đáp án
    var s2 = tmp.insertSheet('DapAn');
    var nCol = header.length, nRow = dataRows.length + 1;
    s2.getRange(1, 1, 1, nCol).setNumberFormat('@');              // dòng tiêu đề là text
    if (dataRows.length) s2.getRange(2, 2, dataRows.length, nCol - 1).setNumberFormat('@'); // đáp án là text
    s2.getRange(1, 1, nRow, nCol).setValues([header].concat(dataRows));
    SpreadsheetApp.flush();

    var blob = exportSpreadsheetAsXlsxBlob(tmp.getId(), 'dap-an.xlsx');
    var folder = DriveApp.getFolderById(bai.driveFolderId);
    var name = 'dap-an - ' + safeName(bai.tenBai) + '.xlsx';
    var file = saveBlobToFolder(folder, blob, name);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);


    return { ok: true, fileId: file.getId(), viewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view', downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId() };
  } finally {
    try { DriveApp.getFileById(tmp.getId()).setTrashed(true); } catch (e) {}
  }
}

/** Trang một bài: thông tin bài + ảnh phiếu trong MỘT lượt gọi (trước đây 2 lượt) */
function getBaiFull(p, u) {
  var b = getBai(p, u);
  if (isWorker()) { var bb = findBai(p.maBai); b.anh = { count: Number(bb && bb.soAnh) || 0, folderUrl: bb && bb.driveFolderId ? 'https://drive.google.com/drive/folders/' + bb.driveFolderId : '' }; }
  else { try { b.anh = anhInfo(p, u); } catch (e) { b.anh = { count: 0, folderUrl: '', loi: String(e.message || e) }; } }
  return b;
}

/** Lưu trữ bài: ẩn khỏi danh sách đang làm (vẫn giữ nguyên dữ liệu, điểm đã công bố vẫn tra được) */
function luuTruBai(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  var sh = _bai(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  if (H['luuTru'] === undefined) throw new Error('Dữ liệu chưa được nâng cấp, hãy tải lại trang');
  sh.getRange(bai._row, H['luuTru'] + 1).setValue(p.luuTru ? 'TRUE' : '');
  sysTouch();
  return { ok: true, luuTru: !!p.luuTru };
}

/**
 * Xóa hẳn một bài: kết quả, dòng chờ, đáp án, bài, và thư mục trên Drive (vào thùng rác Drive).
 * Bắt buộc gõ lại đúng tên bài để xác nhận. Nhật ký sửa điểm được giữ lại.
 */
function xoaBai(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  if (String(p.xacNhan || '').trim() !== String(bai.tenBai).trim()) throw new Error('Gõ lại đúng tên bài để xác nhận xóa');
  var daCongBo = bai.trangThai === 'Đã công bố';
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var kqRows = kqIndex(bai.maBai).map(function (x) { return x.row; });
    if (kqRows.length) deleteRowsList(_kq(), kqRows);
    var cx = getSystemSheet().getSheetByName('ChoXuLy');
    if (cx) { var cr = cxlLoad(bai.maBai).map(function (x) { return x._row; }); if (cr.length) deleteRowsList(cx, cr); }
    var da = sysRows('DapAn').filter(function (r) { return String(r.maBai) === String(bai.maBai); }).map(function (r) { return r._row; });
    if (da.length) deleteRowsList(_dapan(), da);
    var cur = findBai(bai.maBai);
    if (cur) deleteRowsList(_bai(), [cur._row]);
    try { DriveApp.getFolderById(bai.driveFolderId).setTrashed(true); } catch (e) {}
    CacheService.getScriptCache().remove('baimeta_' + bai.maBai);
    logSua(u, bai.maBai, '(bài)', 'xóa bài', String(bai.tenBai), '(đã xóa)', 'xóa hẳn bài');
    sysTouch();
    if (daCongBo) pubTouch();
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/** Lưu một file (ví dụ bảng điểm Excel do trình duyệt tạo) vào thư mục của bài trên Drive */
function luuFileDrive(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  var data = String(p.data || '');
  if (!data) throw new Error('File rỗng');
  if (data.length > 14 * 1024 * 1024) throw new Error('File quá lớn');
  var name = String(p.tenFile || 'file.xlsx').replace(/[\\/:*?"<>|]/g, '-').trim().slice(0, 120) || 'file.xlsx';
  var mime = /\.xlsx$/i.test(name) ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/octet-stream';
  var folder = DriveApp.getFolderById(bai.driveFolderId);
  var file = saveBlobToFolder(folder, Utilities.newBlob(Utilities.base64Decode(data), mime, name), name);
  return { ok: true, url: 'https://drive.google.com/file/d/' + file.getId() + '/view' };
}


/* ===================== Diem.gs ===================== */
/** ĐIỂM: nhập file TNMaker, đối soát, xem kết quả, sửa điểm (có nhật ký), công bố */

var DIEM_TOL = 0.011;
var TRANG_THAI_CHUA_CO_DIEM = ['Nháp', 'Chờ chấm', 'Sẵn sàng chấm'];

/* ---------------- Sheet KetQua ---------------- */

function _kq() {
  var ss = getSystemSheet();
  var cache = CacheService.getScriptCache();
  if (!cache.get('kqcols_ok2')) {
    ensureSysColumns(ss);                    // thêm cột daSua, nguon nếu bản cũ chưa có
    cache.put('kqcols_ok2', '1', 21600);
  }
  return ss.getSheetByName('KetQua');
}

function kqHeader(sh) {
  return idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
}

/** Danh sách {row, sbd} các dòng của một bài (chỉ đọc 2 cột đầu cho nhanh) */
function kqIndex(maBai) {
  var sh = _kq(), last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, 2).getValues(), out = [];
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]) === String(maBai)) out.push({ row: i + 2, sbd: normSbd(v[i][1]) });
  }
  return out;
}

function countKetQua(maBai) {
  try { return kqIndex(maBai).length; } catch (e) { return 0; }
}

function parseJsonSafe(s, def) {
  try { return JSON.parse(s); } catch (e) { return def; }
}

/** Đọc đầy đủ các dòng điểm của một bài */
function kqLoad(maBai) {
  var ix = kqIndex(maBai);
  if (!ix.length) return [];
  var sh = _kq(), H = kqHeader(sh), nCol = sh.getLastColumn();
  var r0 = ix[0].row, r1 = ix[ix.length - 1].row;
  var vals = sh.getRange(r0, 1, r1 - r0 + 1, nCol).getValues(), out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (String(v[H['maBai']]) !== String(maBai)) continue;
    out.push({
      _row: r0 + i, sbd: normSbd(v[H['sbd']]), maDe: normMaDe(v[H['maDe']]),
      tong: Number(v[H['tong']]), p1: Number(v[H['p1']]), p2: Number(v[H['p2']]), p3: Number(v[H['p3']]),
      tl: parseJsonSafe(v[H['traLoi']], { p1: [], p2: [], p3: [] }),
      daSua: toBool(v[H['daSua']]),
      nguon: H['nguon'] !== undefined ? (normSbd(v[H['nguon']]) || '') : ''     // SBD ghi trong file TNMaker, nếu đã được gán sang em khác
    });
  }
  return out;
}

function kqRowArray(H, nCol, bai, o) {
  var a = []; for (var i = 0; i < nCol; i++) a.push('');
  a[H['maBai']] = bai; a[H['sbd']] = o.sbd; a[H['maDe']] = o.maDe;
  a[H['tong']] = o.tong; a[H['p1']] = o.p1; a[H['p2']] = o.p2; a[H['p3']] = o.p3;
  a[H['traLoi']] = JSON.stringify(o.tl || {}); a[H['capNhat']] = nowIso();
  a[H['daSua']] = o.daSua ? 'TRUE' : '';
  if (H['nguon'] !== undefined) a[H['nguon']] = o.nguon || '';
  return a;
}

/** Xóa các dòng (số dòng tuyệt đối), gom thành từng đoạn liền nhau để ít lệnh nhất */
function deleteRowsList(sh, rows) {
  rows = rows.slice().sort(function (a, b) { return b - a; });
  var i = 0;
  while (i < rows.length) {
    var end = rows[i], start = end, j = i + 1;
    while (j < rows.length && rows[j] === start - 1) { start = rows[j]; j++; }
    sh.deleteRows(start, end - start + 1);
    i = j;
  }
  sysTouch();
}

function tStr(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
  return v === null || v === undefined ? '' : String(v);
}

function logSua(user, maBai, sbd, truong, cu, moi, lyDo) {
  var sh = getSystemSheet().getSheetByName('NhatKySua'), row = sh.getLastRow() + 1;
  var vals = [nowIso(), user.username, maBai, sbd, truong, String(cu), String(moi), lyDo || ''];
  sh.getRange(row, 1, 1, vals.length).setNumberFormat('@');   // lý do kiểu "1/2" không bị Sheets đổi thành ngày
  sh.getRange(row, 1, 1, vals.length).setValues([vals]);
  sysTouch('NhatKySua');
}

/* ---------------- Danh sách học sinh ---------------- */

/** Đọc mọi nhóm mà người dùng được xem: maNhom -> {tenNhom, teacher, students[]}. fresh = đọc thẳng file */
function loadAccessibleGroups(u, fresh) {
  var nhom = getNhomMap(), byFile = {};
  Object.keys(nhom).forEach(function (ma) {
    var info = nhom[ma];
    if (u.role !== 'admin' && info.teacher !== u.username) return;
    if (!info.rosterSheetId) return;
    (byFile[info.rosterSheetId] = byFile[info.rosterSheetId] || []).push({ ma: ma, info: info });
  });
  var out = {};
  Object.keys(byFile).forEach(function (fid) {
    var tabs;
    try { tabs = rosterSnapshot(fid, !!fresh); } catch (e) { return; }
    byFile[fid].forEach(function (x) {
      var g = findTab(tabs, x.info.sheetName, x.ma);
      if (g) out[x.ma] = { maNhom: x.ma, tenNhom: g.tenNhom, teacher: x.info.teacher, students: g.students };
    });
  });
  return out;
}

function assertGroupAccessBySbd(u, sbd) {
  var info = getNhomMap()[sbd.substring(0, 3)];
  if (!info) throw new Error('SBD ' + sbd + ' thuộc nhóm chưa đăng ký');
  if (u.role !== 'admin' && info.teacher !== u.username) throw new Error('SBD ' + sbd + ' không thuộc nhóm của bạn');
}

/* ---------------- Đối chiếu đáp án ---------------- */

function keyComplete(cfg, k) {
  if (!k) return false;
  var i;
  if (cfg.p1 && cfg.p1.on) for (i = 0; i < cfg.p1.soCau; i++) if (!k.p1 || !String(k.p1[i] || '').trim()) return false;
  if (cfg.p2 && cfg.p2.on) for (i = 0; i < cfg.p2.soCau; i++) if (!k.p2 || normP2(k.p2[i]).indexOf('_') !== -1) return false;
  if (cfg.p3 && cfg.p3.on) for (i = 0; i < cfg.p3.soCau; i++) if (!k.p3 || !normP3(k.p3[i])) return false;
  return true;
}

/** a = đáp án trên hệ thống, b = đáp án trong file TNMaker */
function keyDiffs(cfg, a, b) {
  var d = [], i;
  if (cfg.p1 && cfg.p1.on) for (i = 0; i < cfg.p1.soCau; i++) {
    var x1 = String((a.p1 || [])[i] || '').toUpperCase(), y1 = String((b.p1 || [])[i] || '').toUpperCase();
    if (x1 !== y1) d.push('Phần I câu ' + (i + 1) + ': hệ thống ' + x1 + ', file ' + y1);
  }
  if (cfg.p2 && cfg.p2.on) for (i = 0; i < cfg.p2.soCau; i++) {
    var x2 = normP2((a.p2 || [])[i]), y2 = normP2((b.p2 || [])[i]);
    if (x2 !== y2) d.push('Phần II câu ' + (i + 1) + ': hệ thống ' + x2 + ', file ' + y2);
  }
  if (cfg.p3 && cfg.p3.on) for (i = 0; i < cfg.p3.soCau; i++) {
    var x3 = normP3((a.p3 || [])[i]), y3 = normP3((b.p3 || [])[i]);
    if (x3 !== y3) d.push('Phần III câu ' + (i + 1) + ': hệ thống ' + x3 + ', file ' + y3);
  }
  return d;
}

function shapeOfCfg(cfg) {
  return {
    n1: (cfg.p1 && cfg.p1.on) ? Number(cfg.p1.soCau) : 0,
    n2: (cfg.p2 && cfg.p2.on) ? Number(cfg.p2.soCau) : 0,
    n3: (cfg.p3 && cfg.p3.on) ? Number(cfg.p3.soCau) : 0
  };
}

/* ---------------- Xem trước và đối soát ---------------- */

function readRows(p) {
  var rows = p.rows;
  if (typeof rows === 'string') rows = JSON.parse(rows);
  if (!rows || !rows.length) throw new Error('File không có dòng điểm nào');
  if (rows.length > 3000) throw new Error('File quá lớn (trên 3000 dòng)');
  return rows;
}

/**
 * Chưa ghi gì. Trả về từng dòng kèm các vấn đề cần xử lý, danh sách em vắng,
 * kết quả đối chiếu với điểm TNMaker, và các cảnh báo về đáp án.
 */
function previewDiem(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  if (laBaiNgoai(bai)) throw new Error('Bài này chỉ có điểm tổng: dùng mục "Nhập điểm" của bài');
  var cfg = parseJsonSafe(bai.cauHinh, {});
  var rows = readRows(p);
  var keys = p.keys || {};
  if (typeof keys === 'string') keys = JSON.parse(keys);

  var groups = loadAccessibleGroups(u, true);
  var bySbd = {}, byCu = {};
  Object.keys(groups).forEach(function (ma) {
    groups[ma].students.forEach(function (s) {
      if (s.sbd) (bySbd[s.sbd] = bySbd[s.sbd] || []).push({ ho: s.ho, ten: s.ten, maNhom: ma, trangThai: s.trangThai });
      if (s.sbdCu) byCu[s.sbdCu] = s.sbd;
    });
  });

  var existing = {}, existCount = 0, edited = [], byNguon = {};
  kqLoad(bai.maBai).forEach(function (r) { existing[r.sbd] = r; existCount++; if (r.daSua) edited.push(r.sbd); if (r.nguon) byNguon[r.nguon] = r; });

  var cacNhom = parseNhomList(bai.cacNhom);
  var maDeList = (cfg.maDeList || []).map(normMaDe);

  var count = {};
  rows.forEach(function (r) { var s = normSbd(r.sbd); count[s] = (count[s] || 0) + 1; });

  var out = [], present = {}, khop = 0, lech = 0, khongKiemTra = 0;
  rows.forEach(function (r, i) {
    var sbd = normSbd(r.sbd), maDe = normMaDe(r.maDe);
    var o = { i: i, sbd: sbd, maDe: maDe, tong: r.tong, issues: [] };
    var list = bySbd[sbd] || [];
    if (!/^\d{6}$/.test(sbd)) {
      o.issues.push({ code: 'KHONG_CO_SBD', level: 'error', msg: 'SBD không hợp lệ: ' + r.sbd });
    } else if (!list.length) {
      o.issues.push({ code: 'KHONG_CO_SBD', level: 'error', msg: 'SBD không có trong danh sách' });
    } else if (list.length > 1) {
      o.issues.push({ code: 'TRUNG_DS', level: 'error', msg: 'SBD đang dùng cho ' + list.length + ' học sinh: ' +
        list.map(function (x) { return (x.ho + ' ' + x.ten).trim() + ' (nhóm ' + x.maNhom + ')'; }).join('; ') + '. Hãy sửa danh sách trước' });
    } else {
      var st = list[0];
      o.ho = st.ho; o.ten = st.ten; o.maNhom = st.maNhom;
      if (st.trangThai === 'Đã nghỉ') o.issues.push({ code: 'DA_NGHI', level: 'warn', msg: 'Học sinh này đã nghỉ' });
      if (st.trangThai === 'Đã chuyển') {
        o.goiY = byCu[sbd] || '';
        o.issues.push({ code: 'DA_CHUYEN', level: 'warn', msg: 'Học sinh đã chuyển nhóm' + (o.goiY ? ', SBD mới là ' + o.goiY : '') });
      }
      if (cacNhom.length && cacNhom.indexOf(st.maNhom) === -1) o.issues.push({ code: 'NGOAI_BAI', level: 'warn', msg: 'Nhóm ' + st.maNhom + ' chưa nằm trong danh sách nhóm làm bài (sẽ được thêm vào)' });
    }
    var gan = byNguon[sbd];
    if (gan && o.issues.length && (o.issues[0].code === 'KHONG_CO_SBD' || o.issues[0].code === 'TRUNG_DS')) {
      var who = (bySbd[gan.sbd] || [])[0];
      o.issues = [{ code: 'DA_GAN', level: 'warn', msg: 'Lần nhập trước đã gán dòng này cho SBD ' + gan.sbd + (who ? ' (' + (who.ho + ' ' + who.ten).trim() + ')' : '') }];
      o.daGan = gan.sbd;
      if (gan.daSua) o.issues.push({ code: 'DA_SUA_TAY', level: 'warn', msg: 'Kết quả của em này đã được sửa tay trước đó' });
    }
    if (count[sbd] > 1) o.issues.push({ code: 'TRUNG_FILE', level: 'warn', msg: 'SBD xuất hiện ' + count[sbd] + ' lần trong file' });
    if (maDe && maDeList.indexOf(maDe) === -1) o.issues.push({ code: 'MA_DE_LA', level: 'warn', msg: 'Mã đề ' + maDe + ' chưa có trong bài (sẽ được thêm vào)' });
    if (existing[sbd]) {
      o.daCo = true;
      if (existing[sbd].daSua) o.issues.push({ code: 'DA_SUA_TAY', level: 'warn', msg: 'Kết quả này đã được sửa tay trước đó' });
    }
    if (r.tl && keys[maDe]) {
      var sc = chamBai(cfg, r.tl, keys[maDe]);
      if (sc.thieuKey) khongKiemTra++;
      else if (Math.abs(sc.tong - r.tong) > DIEM_TOL || Math.abs(sc.p1 - r.p1) > DIEM_TOL || Math.abs(sc.p2 - r.p2) > DIEM_TOL || Math.abs(sc.p3 - r.p3) > DIEM_TOL) {
        lech++;
        o.issues.push({ code: 'LECH_DIEM', level: 'warn', msg: 'Web tính ' + sc.tong + ' điểm, TNMaker ' + r.tong + ' (kiểm tra cấu hình bài trên app TNMaker)' });
      } else khop++;
    } else khongKiemTra++;
    if (/^\d{6}$/.test(sbd)) present[sbd] = true;
    out.push(o);
  });

  var maHienCo = cacNhom.length ? cacNhom : Object.keys(out.reduce(function (acc, o) { if (o.maNhom) acc[o.maNhom] = 1; return acc; }, {}));
  var absent = [];
  maHienCo.forEach(function (ma) {
    var g = groups[ma]; if (!g) return;
    g.students.forEach(function (s) {
      if (s.trangThai === 'Đang học' && s.sbd && !present[s.sbd] && !existing[s.sbd]) absent.push({ sbd: s.sbd, ho: s.ho, ten: s.ten, maNhom: ma });
    });
  });
  absent.sort(function (a, b) { return a.sbd < b.sbd ? -1 : 1; });

  var dapAn = getBaiDapAn(bai.maBai), keyConflicts = [];
  Object.keys(keys).forEach(function (md) {
    var e = dapAn[normMaDe(md)];
    if (e && keyComplete(cfg, e)) {
      var d = keyDiffs(cfg, e, keys[md]);
      if (d.length) keyConflicts.push({ maDe: normMaDe(md), soCau: d.length, chiTiet: d.slice(0, 6) });
    }
  });

  var shapeWarn = '';
  if (p.shape) {
    var sc0 = shapeOfCfg(cfg), sh = p.shape;
    if (sc0.n1 !== sh.n1 || sc0.n2 !== sh.n2 || sc0.n3 !== sh.n3) {
      shapeWarn = 'File có Phần I ' + sh.n1 + ' câu, Phần II ' + sh.n2 + ' câu, Phần III ' + sh.n3 + ' câu; cấu hình bài trên web là ' +
        sc0.n1 + ' / ' + sc0.n2 + ' / ' + sc0.n3 + '. Nên sửa cấu hình bài cho khớp để tính lại điểm chính xác.';
    }
  }

  var tenNhom = {};
  Object.keys(groups).forEach(function (ma) { tenNhom[ma] = groups[ma].tenNhom; });

  return {
    rows: out, absent: absent, tenNhom: tenNhom,
    existing: { count: existCount, edited: edited },
    keyConflicts: keyConflicts, shapeWarn: shapeWarn,
    doiChieu: { khop: khop, lech: lech, khongKiemTra: khongKiemTra },
    tenBai: bai.tenBai, choXuLy: countChoXuLy(bai.maBai)
  };
}

/* ---------------- Ghi điểm ---------------- */

function numOrThrow(v, name, sbd) {
  var n = Number(v);
  if (!isFinite(n) || n < 0 || n > 10) throw new Error('Điểm ' + name + ' không hợp lệ (SBD ' + sbd + ')');
  return n;
}

/**
 * Ghi các dòng đã xử lý xong (client đã áp dụng lựa chọn: bỏ qua / gán SBD khác).
 * p.mode: 'merge' (mặc định, cập nhật các em trong file, giữ nguyên các em khác) hoặc 'replace' (xóa hết cũ).
 */
function commitDiem(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  if (laBaiNgoai(bai)) throw new Error('Bài này chỉ có điểm tổng: dùng mục "Nhập điểm" của bài');
  var cfg = parseJsonSafe(bai.cauHinh, {});
  var mode = p.mode === 'replace' ? 'replace' : 'merge';
  var rows = p.rows;
  if (typeof rows === 'string') rows = JSON.parse(rows);
  rows = rows || [];
  var deferred = p.deferred || [];
  if (typeof deferred === 'string') deferred = JSON.parse(deferred);
  var purge = p.purge || [];
  if (typeof purge === 'string') purge = JSON.parse(purge);
  if (!rows.length && !deferred.length && mode !== 'replace') throw new Error('Không có dòng nào để ghi');
  var keys = p.keys || {};
  if (typeof keys === 'string') keys = JSON.parse(keys);

  var nhom = getNhomMap(), seen = {}, clean = [], addNhom = {}, maDeNew = {}, purgeSbd = {}, maDeKeys = {};
  purge.forEach(function (x) { var n = normSbd(x); if (n) purgeSbd[n] = true; });
  var pendRows = deferred.map(function (r) {
    var md = normMaDe(r.maDe), sb = normSbd(r.sbd) || String(r.sbd == null ? '' : r.sbd);
    if (md) maDeKeys[md] = true;
    if (normSbd(r.sbd)) purgeSbd[normSbd(r.sbd)] = true;
    var num = function (v) { var n = Number(v); return isFinite(n) ? n : 0; };
    return { sbd: sb, maDe: md, tong: num(r.tong), p1: num(r.p1), p2: num(r.p2), p3: num(r.p3), tl: r.tl || null,
      lyDo: String(r.lyDo || '').slice(0, 30), ghiChu: String(r.ghiChu || '').slice(0, 300) };
  });
  rows.forEach(function (r) {
    var sbd = normSbd(r.sbd);
    if (!/^\d{6}$/.test(sbd)) throw new Error('SBD không hợp lệ: ' + r.sbd);
    var ma = sbd.substring(0, 3), info = nhom[ma];
    if (!info) throw new Error('SBD ' + sbd + ' thuộc nhóm ' + ma + ' chưa đăng ký');
    if (u.role !== 'admin' && info.teacher !== u.username) throw new Error('SBD ' + sbd + ' không thuộc nhóm của bạn');
    if (seen[sbd]) throw new Error('SBD ' + sbd + ' bị trùng sau khi xử lý');
    seen[sbd] = true;
    var maDe = normMaDe(r.maDe);
    if (!maDe) throw new Error('Thiếu mã đề (SBD ' + sbd + ')');
    addNhom[ma] = true; maDeNew[maDe] = true;
    var srcS = normSbd(r.srcSbd);
    purgeSbd[srcS || sbd] = true;
    clean.push({ nguon: (srcS && srcS !== sbd) ? srcS : '', sbd: sbd, maDe: maDe, tong: numOrThrow(r.tong, 'tổng', sbd), p1: numOrThrow(r.p1, 'P1', sbd), p2: numOrThrow(r.p2, 'P2', sbd), p3: numOrThrow(r.p3, 'P3', sbd), tl: r.tl || null, daSua: false });
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = _kq(), H = kqHeader(sh), nCol = sh.getLastColumn();
    var ix = kqIndex(bai.maBai), del = [];
    if (mode === 'replace') ix.forEach(function (x) { del.push(x.row); });
    else ix.forEach(function (x) { if (seen[x.sbd]) del.push(x.row); });
    if (del.length) deleteRowsList(sh, del);

    kqAppendRows(bai.maBai, clean);

    // Đáp án theo file TNMaker (điểm đã chấm theo nó) ghi vào bảng DapAn
    Object.keys(maDeNew).forEach(function (m) { maDeKeys[m] = true; });
    var da = _dapan(), mdKeys = Object.keys(keys).map(normMaDe).filter(function (m) { return maDeKeys[m]; });
    if (mdKeys.length) {
      var daDel = [];
      sysRows('DapAn').forEach(function (r) {
        if (String(r.maBai) === String(bai.maBai) && mdKeys.indexOf(normMaDe(r.maDe)) !== -1) daDel.push(r._row);
      });
      if (daDel.length) deleteRowsList(da, daDel);
      var daRows = [];
      Object.keys(keys).forEach(function (md) {
        var m = normMaDe(md); if (!maDeKeys[m]) return;
        daRows.push([bai.maBai, m, JSON.stringify(keys[md])]);
      });
      if (daRows.length) da.getRange(da.getLastRow() + 1, 1, daRows.length, 3).setValues(daRows);
    }

    // Dòng chờ: những SBD có trong file này được thay theo lựa chọn mới (đã ghi / để chờ / bỏ hẳn)
    cxlDelete(bai.maBai, function (r) { return mode === 'replace' || purgeSbd[r.sbd]; });
    cxlAppend(bai.maBai, pendRows, p.tenFile);

    var tt = (clean.length || mode === 'replace') ? capNhatBaiSauKhiGhi(bai, cfg, maDeNew, addNhom) : findBai(bai.maBai).trangThai;
    logSua(u, bai.maBai, '(file)', 'nhập file điểm', mode === 'replace' ? 'xóa hết kết quả cũ' : (del.length + ' dòng cũ bị thay'),
      clean.length + ' dòng' + (pendRows.length ? ', ' + pendRows.length + ' dòng để xử lý sau' : '') + (p.tenFile ? ' từ ' + p.tenFile : ''), '');
    sysTouch();
    if (tt === 'Đã công bố') pubTouch();
    return { ok: true, soDong: clean.length, soChoXuLy: countChoXuLy(bai.maBai), trangThai: tt };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Xem kết quả ---------------- */

function getKetQua(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  return cachedResult('kq_' + bai.maBai + '_' + u.username, 900, function () { return getKetQua_(p, u, bai); }, p.fresh);
}
function getKetQua_(p, u, bai) {
  var cfg = parseJsonSafe(bai.cauHinh, {});
  var res = kqLoad(bai.maBai), groups = loadAccessibleGroups(u, !!p.fresh);
  var cacNhom = parseNhomList(bai.cacNhom);
  var maSet = {};
  cacNhom.forEach(function (m) { maSet[m] = true; });
  res.forEach(function (r) { maSet[r.sbd.substring(0, 3)] = true; });

  var bySbd = {};
  res.forEach(function (r) { bySbd[r.sbd] = r; });
  var used = {}, outGroups = [];
  Object.keys(maSet).sort().forEach(function (ma) {
    var g = groups[ma]; if (!g) return;
    var list = [];
    g.students.forEach(function (s) {
      var r = bySbd[s.sbd];
      if (s.trangThai !== 'Đang học' && !r) return;
      if (r) used[s.sbd] = true;
      list.push({ sbd: s.sbd, ho: s.ho, ten: s.ten, lop: s.lop, trangThai: s.trangThai, sbdCu: s.sbdCu,
        kq: r ? { maDe: r.maDe, tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, daSua: r.daSua, tl: r.tl } : null });
    });
    list.sort(function (a, b) { return a.sbd < b.sbd ? -1 : 1; });
    outGroups.push({ maNhom: ma, tenNhom: g.tenNhom, students: list });
  });
  var bySbdAll = {};
  Object.keys(groups).forEach(function (ma) {
    groups[ma].students.forEach(function (s) { if (s.sbd) (bySbdAll[s.sbd] = bySbdAll[s.sbd] || []).push({ ho: s.ho, ten: s.ten, maNhom: ma, trangThai: s.trangThai }); });
  });
  var choXuLy = cxlLoad(bai.maBai).map(function (r) {
    var o = { id: r.id, sbd: r.sbd, maDe: r.maDe, tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, lyDo: r.lyDo, ghiChu: r.ghiChu, tenFile: r.tenFile };
    var m = bySbdAll[r.sbd] || [];
    if (m.length === 1) o.match = { ho: m[0].ho, ten: m[0].ten, maNhom: m[0].maNhom, trangThai: m[0].trangThai, coKetQua: !!bySbd[r.sbd] };
    else if (m.length > 1) o.trungDS = m.length;
    return o;
  });
  var orphan = [];
  res.forEach(function (r) {
    if (!used[r.sbd] && groups[r.sbd.substring(0, 3)]) orphan.push({ sbd: r.sbd, kq: { maDe: r.maDe, tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, daSua: r.daSua, tl: r.tl } });
  });
  return {
    bai: { maBai: bai.maBai, tenBai: bai.tenBai, ngay: ngayStr(bai.ngay), trangThai: bai.trangThai, ngayCongBo: tStr(bai.ngayCongBo), cauHinh: cfg },
    groups: outGroups, orphan: orphan, choXuLy: choXuLy, keys: getBaiDapAn(bai.maBai)
  };
}

/* ---------------- Sửa điểm ---------------- */

function normTraLoi(cfg, tl) {
  var out = { p1: [], p2: [], p3: [] }, i;
  tl = tl || {};
  var n1 = (cfg.p1 && cfg.p1.on) ? Number(cfg.p1.soCau) : 0;
  var n2 = (cfg.p2 && cfg.p2.on) ? Number(cfg.p2.soCau) : 0;
  var n3 = (cfg.p3 && cfg.p3.on) ? Number(cfg.p3.soCau) : 0;
  for (i = 0; i < n1; i++) { var a = String((tl.p1 || [])[i] || '').trim().toUpperCase(); out.p1.push(/^[A-D]$/.test(a) ? a : ''); }
  for (i = 0; i < n2; i++) out.p2.push(normP2((tl.p2 || [])[i]));
  for (i = 0; i < n3; i++) out.p3.push(normP3((tl.p3 || [])[i]));
  return out;
}

/**
 * Sửa kết quả của một em (hoặc nhập tay cho em vắng). Bắt buộc ghi lý do.
 * Hoặc gửi p.traLoi (+ p.maDe): web tính lại điểm; hoặc gửi trực tiếp p.tong / p.p1 / p.p2 / p.p3.
 */
function editKetQua(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  var sbd = normSbd(p.sbd);
  if (!/^\d{6}$/.test(sbd)) throw new Error('SBD không hợp lệ');
  assertGroupAccessBySbd(u, sbd);
  var lyDo = String(p.lyDo || '').trim();
  if (!lyDo) throw new Error('Phải ghi lý do sửa');
  var cfg = parseJsonSafe(bai.cauHinh, {});

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var cur = kqLoad(bai.maBai).filter(function (r) { return r.sbd === sbd; })[0] || null;
    var maDe = normMaDe(p.maDe || (cur && cur.maDe));
    if (!maDe) throw new Error('Thiếu mã đề');
    var next = cur ? { maDe: cur.maDe, tong: cur.tong, p1: cur.p1, p2: cur.p2, p3: cur.p3, tl: cur.tl } :
      { maDe: maDe, tong: 0, p1: 0, p2: 0, p3: 0, tl: { p1: [], p2: [], p3: [] } };
    next.maDe = maDe;

    if (p.traLoi) {
      var key = getBaiDapAn(bai.maBai)[maDe];
      if (!key) throw new Error('Chưa có đáp án mã đề ' + maDe + ', không tính lại được. Hãy sửa điểm trực tiếp.');
      var tl = normTraLoi(cfg, p.traLoi);
      var sc = chamBai(cfg, tl, key);
      if (sc.thieuKey) throw new Error('Đáp án mã đề ' + maDe + ' chưa đủ, không tính lại được. Hãy sửa điểm trực tiếp.');
      next.tl = tl; next.p1 = sc.p1; next.p2 = sc.p2; next.p3 = sc.p3; next.tong = sc.tong;
    } else {
      if (p.p1 !== undefined && p.p1 !== '') next.p1 = numOrThrow(p.p1, 'P1', sbd);
      if (p.p2 !== undefined && p.p2 !== '') next.p2 = numOrThrow(p.p2, 'P2', sbd);
      if (p.p3 !== undefined && p.p3 !== '') next.p3 = numOrThrow(p.p3, 'P3', sbd);
      if (p.tong !== undefined && p.tong !== '') next.tong = numOrThrow(p.tong, 'tổng', sbd);
      else next.tong = round2(next.p1 + next.p2 + next.p3);
    }

    var changes = [];
    if (!cur || cur.maDe !== next.maDe) changes.push(['mã đề', cur ? cur.maDe : '(chưa có)', next.maDe]);
    ['tong', 'p1', 'p2', 'p3'].forEach(function (f) {
      var a = cur ? cur[f] : '', b = next[f];
      if (!cur || Math.abs(Number(a) - Number(b)) > 1e-9) changes.push([f === 'tong' ? 'tổng điểm' : f.toUpperCase(), cur ? a : '(chưa có)', b]);
    });
    if (p.traLoi && (!cur || JSON.stringify(cur.tl) !== JSON.stringify(next.tl))) changes.push(['câu trả lời', cur ? 'bản cũ' : '(chưa có)', 'đã sửa']);
    if (!changes.length) throw new Error('Không có gì thay đổi');

    var sh = _kq(), H = kqHeader(sh), nCol = sh.getLastColumn();
    var rowArr = kqRowArray(H, nCol, bai.maBai, { nguon: cur ? cur.nguon : '', sbd: sbd, maDe: next.maDe, tong: next.tong, p1: next.p1, p2: next.p2, p3: next.p3, tl: next.tl, daSua: true });
    if (cur) {
      sh.getRange(cur._row, 1, 1, nCol).setValues([rowArr]);
    } else {
      var start = sh.getLastRow() + 1;
      sh.getRange(start, H['sbd'] + 1, 1, 2).setNumberFormat('@');
      sh.getRange(start, 1, 1, nCol).setValues([rowArr]);
    }
    changes.forEach(function (c) { logSua(u, bai.maBai, sbd, c[0], c[1], c[2], lyDo); });

    var bs = _bai(), HB = idx(bs.getRange(1, 1, 1, bs.getLastColumn()).getValues()[0]);
    if (TRANG_THAI_CHUA_CO_DIEM.indexOf(bai.trangThai) !== -1) bs.getRange(bai._row, HB['trangThai'] + 1).setValue('Đã nhập điểm');
    sysTouch();
    if (bai.trangThai === 'Đã công bố') pubTouch();
    return { ok: true, kq: { maDe: next.maDe, tong: next.tong, p1: next.p1, p2: next.p2, p3: next.p3, daSua: true, tl: next.tl } };
  } finally {
    lock.releaseLock();
  }
}

/** Xóa kết quả của một em (nhập nhầm). Bắt buộc ghi lý do. */
function xoaKetQua(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  var sbd = normSbd(p.sbd);
  assertGroupAccessBySbd(u, sbd);
  var lyDo = String(p.lyDo || '').trim();
  if (!lyDo) throw new Error('Phải ghi lý do xóa');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var cur = kqLoad(bai.maBai).filter(function (r) { return r.sbd === sbd; })[0];
    if (!cur) throw new Error('Em này chưa có kết quả');
    deleteRowsList(_kq(), [cur._row]);
    logSua(u, bai.maBai, sbd, 'xóa kết quả', cur.tong + ' điểm (mã ' + cur.maDe + ')', '(đã xóa)', lyDo);
    if (bai.trangThai === 'Đã công bố') pubTouch();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Công bố, nhật ký ---------------- */

function setCongBo(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  var bs = _bai(), HB = idx(bs.getRange(1, 1, 1, bs.getLastColumn()).getValues()[0]);
  var tt;
  if (p.congBo) {
    if (!countKetQua(bai.maBai)) throw new Error('Bài chưa có điểm để công bố');
    tt = 'Đã công bố';
    bs.getRange(bai._row, HB['trangThai'] + 1).setValue(tt);
    bs.getRange(bai._row, HB['ngayCongBo'] + 1).setValue(nowIso());
  } else {
    tt = (bai.trangThai === 'Đã công bố') ? 'Đã nhập điểm' : bai.trangThai;
    bs.getRange(bai._row, HB['trangThai'] + 1).setValue(tt);
    bs.getRange(bai._row, HB['ngayCongBo'] + 1).setValue('');
  }
  logSua(u, bai.maBai, '(bài)', p.congBo ? 'công bố' : 'ẩn điểm', bai.trangThai, tt, '');
  sysTouch(); pubTouch();
  return { ok: true, trangThai: tt };
}

function getNhatKy(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  var rows = sysRows('NhatKySua').filter(function (r) { return String(r.maBai) === String(bai.maBai); });
  rows.reverse();
  return {
    nhatKy: rows.slice(0, 200).map(function (r) {
      return { thoiGian: tStr(r.thoiGian), nguoiSua: r.nguoiSua, sbd: normSbd(r.sbd) || String(r.sbd), truong: r.truong, cu: String(r.giaTriCu), moi: String(r.giaTriMoi), lyDo: r.lyDo };
    })
  };
}

/* ---------------- Dùng chung khi ghi điểm ---------------- */

/** Ghi thêm các dòng kết quả vào cuối bảng KetQua (sbd, mã đề giữ dạng chữ để không mất số 0 đầu) */
function kqAppendRows(maBai, clean) {
  if (!clean.length) return;
  var sh = _kq(), H = kqHeader(sh), nCol = sh.getLastColumn();
  var arr = clean.map(function (o) { return kqRowArray(H, nCol, maBai, o); });
  var start = sh.getLastRow() + 1;
  sh.getRange(start, H['sbd'] + 1, arr.length, 2).setNumberFormat('@');
  sh.getRange(start, 1, arr.length, nCol).setValues(arr);
}

/** Sau khi ghi điểm: thêm mã đề / nhóm mới vào bài, chuyển trạng thái. Trả về trạng thái mới. */
function capNhatBaiSauKhiGhi(bai, cfg, maDeNew, addNhom) {
  var bs = _bai(), HB = idx(bs.getRange(1, 1, 1, bs.getLastColumn()).getValues()[0]);
  var curBai = findBai(bai.maBai);
  var maDeList = (cfg.maDeList || []).map(normMaDe);
  Object.keys(maDeNew).forEach(function (m) { if (maDeList.indexOf(m) === -1) maDeList.push(m); });
  maDeList.sort();
  cfg.maDeList = maDeList;
  bs.getRange(curBai._row, HB['cauHinh'] + 1).setValue(JSON.stringify(cfg));
  var cacNhom = parseNhomList(curBai.cacNhom);
  Object.keys(addNhom).forEach(function (m) { if (cacNhom.indexOf(m) === -1) cacNhom.push(m); });
  cacNhom.sort();
  setTextCell(bs, curBai._row, HB['cacNhom'] + 1, cacNhom.join(','));
  var tt = curBai.trangThai;
  if (tt !== 'Đã công bố') { tt = 'Đã nhập điểm'; bs.getRange(curBai._row, HB['trangThai'] + 1).setValue(tt); }
  sysTouch();
  return tt;
}

/* ---------------- Dòng điểm CHỜ XỬ LÝ ----------------
 * Khi nhập file, dòng nào chưa quyết được (em mới học chưa kịp thêm vào danh sách, SBD trùng…) được GIỮ LẠI
 * ở đây thay vì bị vứt. Thêm/sửa danh sách lớp xong thì ghép lại lúc nào cũng được.
 */

function _cxl() {
  var ss = getSystemSheet(), sh = ss.getSheetByName('ChoXuLy');
  if (!sh) { sh = ss.insertSheet('ChoXuLy'); sh.appendRow(SYS_TABS.ChoXuLy); }
  return sh;
}

function cxlLoad(maBai) {
  var sh = getSystemSheet().getSheetByName('ChoXuLy');
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  var H = idx(vals[0]), out = [];
  for (var i = 1; i < vals.length; i++) {
    var v = vals[i];
    if (String(v[H['maBai']]) !== String(maBai)) continue;
    var raw = String(v[H['sbd']]);
    out.push({
      _row: i + 1, id: String(v[H['id']]), sbd: normSbd(raw) || raw, maDe: normMaDe(v[H['maDe']]),
      tong: Number(v[H['tong']]), p1: Number(v[H['p1']]), p2: Number(v[H['p2']]), p3: Number(v[H['p3']]),
      tl: parseJsonSafe(v[H['traLoi']], null), lyDo: String(v[H['lyDo']]), ghiChu: String(v[H['ghiChu']]), tenFile: String(v[H['tenFile']])
    });
  }
  return out;
}

function countChoXuLy(maBai) {
  try { return cxlLoad(maBai).length; } catch (e) { return 0; }
}

function cxlAppend(maBai, rows, tenFile) {
  if (!rows.length) return;
  var sh = _cxl(), n = SYS_TABS.ChoXuLy.length, now = nowIso(), base = Date.now().toString(36);
  var arr = rows.map(function (r, i) {
    return ['X' + base + '_' + i + '_' + Math.random().toString(36).slice(2, 6), maBai, r.sbd, r.maDe, r.tong, r.p1, r.p2, r.p3,
      JSON.stringify(r.tl || null), r.lyDo || '', r.ghiChu || '', tenFile || '', now];
  });
  var start = sh.getLastRow() + 1;
  sh.getRange(start, 1, arr.length, 4).setNumberFormat('@');          // id, maBai, sbd, maDe giữ dạng chữ
  sh.getRange(start, 1, arr.length, n).setValues(arr);
}

/** Xóa các dòng chờ của bài thỏa điều kiện fn(dòng) */
function cxlDelete(maBai, fn) {
  var sh = getSystemSheet().getSheetByName('ChoXuLy');
  if (!sh) return 0;
  var rows = cxlLoad(maBai).filter(fn).map(function (r) { return r._row; });
  if (rows.length) deleteRowsList(sh, rows);
  return rows.length;
}

/**
 * Xử lý các dòng chờ.
 *  - p.items = [{id, action:'assign', to:'104034'} | {id, action:'discard'}]
 *  - hoặc p.auto = true: tự ghép những dòng có SBD đã khớp ĐÚNG MỘT học sinh đang học và chưa có kết quả.
 * Lỗi từng dòng được báo riêng trong `errors`, các dòng còn lại vẫn được xử lý.
 */
function giaiQuyetChoXuLy(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  var cfg = parseJsonSafe(bai.cauHinh, {});
  if (cfg.loai === 'ngoai') return ngoaiGiaiQuyet(p, u, bai);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var pend = cxlLoad(bai.maBai);
    if (!pend.length) return { ok: true, resolved: 0, discarded: 0, errors: [], conLai: 0 };
    var groups = loadAccessibleGroups(u, true), bySbd = {};
    Object.keys(groups).forEach(function (ma) {
      groups[ma].students.forEach(function (s) { if (s.sbd) (bySbd[s.sbd] = bySbd[s.sbd] || []).push(s); });
    });
    var existing = {};
    kqLoad(bai.maBai).forEach(function (r) { existing[r.sbd] = r; });

    var items = p.items || [];
    if (typeof items === 'string') items = JSON.parse(items);
    if (p.auto) {
      items = [];
      pend.forEach(function (r) {
        var m = bySbd[r.sbd] || [];
        if (m.length === 1 && m[0].trangThai === 'Đang học' && !existing[r.sbd] && r.maDe) items.push({ id: r.id, action: 'assign', to: r.sbd });
      });
    }

    var byId = {};
    pend.forEach(function (r) { byId[r.id] = r; });
    var errors = [], delRows = [], write = [], addNhom = {}, maDeNew = {}, used = {}, resolved = 0, discarded = 0;
    items.forEach(function (it) {
      var r = byId[it.id];
      if (!r) return;                                                   // đã được xử lý trước đó
      if (it.action === 'discard') {
        delRows.push(r._row); discarded++;
        logSua(u, bai.maBai, r.sbd, 'bỏ dòng chờ', r.tong + ' điểm (mã ' + r.maDe + ')', '(đã bỏ)', 'bỏ dòng chờ xử lý');
        return;
      }
      var to = normSbd(it.to);
      if (!/^\d{6}$/.test(to)) { errors.push({ id: r.id, msg: 'SBD không hợp lệ' }); return; }
      try { assertGroupAccessBySbd(u, to); } catch (e) { errors.push({ id: r.id, msg: e.message }); return; }
      var m = bySbd[to] || [];
      if (m.length !== 1) { errors.push({ id: r.id, msg: m.length ? 'SBD ' + to + ' đang dùng cho nhiều học sinh' : 'SBD ' + to + ' chưa có trong danh sách lớp' }); return; }
      if (existing[to] || used[to]) { errors.push({ id: r.id, msg: (m[0].ho + ' ' + m[0].ten).trim() + ' đã có kết quả bài này' }); return; }
      if (!r.maDe) { errors.push({ id: r.id, msg: 'Dòng này thiếu mã đề' }); return; }
      used[to] = true;
      write.push({ nguon: (r.sbd && r.sbd !== to) ? r.sbd : '', sbd: to, maDe: r.maDe, tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, tl: r.tl, daSua: false });
      addNhom[to.substring(0, 3)] = true; maDeNew[r.maDe] = true;
      delRows.push(r._row); resolved++;
      logSua(u, bai.maBai, to, 'xử lý dòng chờ', 'dòng SBD "' + r.sbd + '" trong file', r.tong + ' điểm (mã ' + r.maDe + ')', 'ghép dòng chờ xử lý');
    });

    kqAppendRows(bai.maBai, write);
    var sh = getSystemSheet().getSheetByName('ChoXuLy');
    if (delRows.length) deleteRowsList(sh, delRows);
    var tt = write.length ? capNhatBaiSauKhiGhi(bai, cfg, maDeNew, addNhom) : bai.trangThai;
    sysTouch();
    if (tt === 'Đã công bố' && write.length) pubTouch();
    return { ok: true, resolved: resolved, discarded: discarded, errors: errors, conLai: countChoXuLy(bai.maBai), trangThai: tt };
  } finally {
    lock.releaseLock();
  }
}


/* ===================== DiemNgoai.gs ===================== */
/**
 * ĐIỂM TỪ NGUỒN KHÁC (học kỳ ở trường, bài khác): bài chỉ có điểm tổng.
 * File/bảng chỉ cần Họ tên + Điểm (có thêm SBD hoặc Lớp thì khớp chắc hơn).
 * Khớp tên với danh sách các nhóm của bài; tên trùng, gõ thiếu dấu, không thấy => cô chọn, hoặc vào "chờ xử lý".
 */
function chuanTen(s) { return String(s || '').normalize('NFC').toLowerCase().replace(/\s+/g, ' ').trim(); }
function boDau(s) { return chuanTen(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd'); }
function laBaiNgoai(bai) { return parseJsonSafe(bai.cauHinh, {}).loai === 'ngoai'; }
function chuanCfgNgoai(cfg) {
  var t = Number(cfg.thang) || 10;
  if (t < 1 || t > 100) throw new Error('Thang điểm từ 1 đến 100');
  return { loai: 'ngoai', thang: t, nguon: String(cfg.nguon || '').trim().slice(0, 60) };
}
function ngoaiBai(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  var cfg = parseJsonSafe(bai.cauHinh, {});
  if (cfg.loai !== 'ngoai') throw new Error('Bài này là bài trắc nghiệm, nhập điểm bằng file TNMaker');
  return { bai: bai, cfg: cfg, thang: Number(cfg.thang) || 10 };
}
/** Học sinh trong các nhóm của bài */
function ngoaiHocSinh(bai) {
  var nhom = getNhomMap(), out = [];
  parseNhomList(bai.cacNhom).forEach(function (ma) {
    var reg = nhom[ma]; if (!reg) return;
    var g = findTab(rosterSnapshot(reg.rosterSheetId, false), reg.sheetName, ma); if (!g) return;
    g.students.forEach(function (s) { if (s.sbd) out.push({ sbd: s.sbd, ten: ((s.ho || '') + ' ' + (s.ten || '')).trim(), lop: String(s.lop || ''), trangThai: s.trangThai, maNhom: ma }); });
  });
  return out;
}
function ngoaiDiem(v, thang) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(String(v).trim().replace(',', '.'));
  if (!isFinite(n) || n < 0 || n > thang) return NaN;
  return Math.round(n * 100) / 100;
}
function ngoaiPick(h) { return { sbd: h.sbd, ten: h.ten, lop: h.lop }; }

/** Đối chiếu các dòng (tên, điểm, [sbd], [lớp]) với danh sách; chưa ghi gì */
function ngoaiDoiChieu(p, u) {
  var x = ngoaiBai(p, u), hs = ngoaiHocSinh(x.bai);
  var bySbd = {}, byTen = {}, byKd = {};
  hs.forEach(function (h) {
    bySbd[h.sbd] = h;
    (byTen[chuanTen(h.ten)] = byTen[chuanTen(h.ten)] || []).push(h);
    (byKd[boDau(h.ten)] = byKd[boDau(h.ten)] || []).push(h);
  });
  var rows = p.rows; if (typeof rows === 'string') rows = JSON.parse(rows);
  if (!rows || !rows.length) throw new Error('Chưa có dòng điểm nào');
  if (rows.length > 600) throw new Error('Tối đa 600 dòng mỗi lần');
  var hep = function (c, lop) {                                   // thu hẹp ứng viên: theo lớp, rồi theo em đang học
    if (c.length > 1 && lop) { var cl = c.filter(function (h) { return chuanTen(h.lop) === chuanTen(lop); }); if (cl.length) c = cl; }
    if (c.length > 1) { var dh = c.filter(function (h) { return h.trangThai === 'Đang học'; }); if (dh.length) c = dh; }
    return c;
  };
  var out = rows.map(function (r, i) {
    var o = { i: i, ten: String(r.ten || '').trim(), lop: String(r.lop || '').trim(), sbdFile: normSbd(r.sbd || ''), diemGoc: String(r.diem === undefined ? '' : r.diem).trim(), ungVien: [] };
    o.diem = ngoaiDiem(r.diem, x.thang);
    if (o.diem === null || isNaN(o.diem)) { o.trangThai = 'loi-diem'; return o; }
    if (o.sbdFile && bySbd[o.sbdFile]) { o.trangThai = 'khop'; o.sbd = o.sbdFile; o.tenDs = bySbd[o.sbdFile].ten; return o; }
    var c = o.ten ? hep(byTen[chuanTen(o.ten)] || [], o.lop) : [];
    if (c.length === 1) { o.trangThai = 'khop'; o.sbd = c[0].sbd; o.tenDs = c[0].ten; return o; }
    if (c.length > 1) { o.trangThai = 'trung-ten'; o.ungVien = c.map(ngoaiPick); return o; }
    var g = o.ten ? hep(byKd[boDau(o.ten)] || [], o.lop) : [];
    if (g.length === 1) { o.trangThai = 'gan-dung'; o.sbd = g[0].sbd; o.tenDs = g[0].ten; o.ungVien = g.map(ngoaiPick); return o; }
    if (g.length > 1) { o.trangThai = 'trung-ten'; o.ungVien = g.map(ngoaiPick); return o; }
    var tenGoi = boDau(o.ten).split(' ').pop();
    o.ungVien = tenGoi ? hs.filter(function (h) { return boDau(h.ten).split(' ').pop() === tenGoi; }).slice(0, 8).map(ngoaiPick) : [];
    o.trangThai = 'khong-thay';
    return o;
  });
  var dem = {}; out.forEach(function (o) { if (o.sbd) dem[o.sbd] = (dem[o.sbd] || 0) + 1; });
  out.forEach(function (o) { if (o.sbd && dem[o.sbd] > 1) o.trungDong = true; });
  return { thang: x.thang, rows: out, hocSinh: hs.filter(function (h) { return h.trangThai === 'Đang học'; }).map(ngoaiPick), daCo: countKetQua(x.bai.maBai) };
}

/** Ghi điểm: gan = [{sbd, diem}], cho = [{ten, diem, lop}] (vào chờ xử lý). cheDo 'thay' (mặc định) | 'gop' */
function ngoaiLuu(p, u) {
  var x = ngoaiBai(p, u), by = {};
  ngoaiHocSinh(x.bai).forEach(function (h) { by[h.sbd] = h; });
  var gan = p.gan, cho = p.cho;
  if (typeof gan === 'string') gan = JSON.parse(gan); if (typeof cho === 'string') cho = JSON.parse(cho);
  gan = gan || []; cho = cho || [];
  var seen = {}, clean = [];
  gan.forEach(function (g) {
    var sbd = normSbd(g.sbd);
    if (!by[sbd]) throw new Error('SBD ' + g.sbd + ' không thuộc các nhóm của bài');
    if (seen[sbd]) throw new Error('Em ' + by[sbd].ten + ' (' + sbd + ') đang có 2 dòng điểm, hãy bỏ bớt 1 dòng');
    seen[sbd] = 1;
    var d = ngoaiDiem(g.diem, x.thang);
    if (d === null || isNaN(d)) throw new Error('Điểm của em ' + sbd + ' không hợp lệ (0 đến ' + x.thang + ')');
    clean.push({ sbd: sbd, maDe: '', tong: d, p1: '', p2: '', p3: '', tl: {}, daSua: false, nguon: 'ngoai' });
  });
  if (!clean.length && !cho.length) throw new Error('Chưa có dòng nào để lưu');
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var thay = p.cheDo !== 'gop', old = kqIndex(x.bai.maBai);
    if (thay) { deleteRowsList(_kq(), old.map(function (o) { return o.row; })); cxlDelete(x.bai.maBai, function () { return true; }); }
    else deleteRowsList(_kq(), old.filter(function (o) { return seen[o.sbd]; }).map(function (o) { return o.row; }));
    kqAppendRows(x.bai.maBai, clean);
    cxlAppend(x.bai.maBai, cho.map(function (c) {
      var d = ngoaiDiem(c.diem, x.thang);
      return { sbd: '', maDe: '', tong: (d === null || isNaN(d)) ? '' : d, p1: '', p2: '', p3: '', tl: null, lyDo: 'Không khớp tên trong danh sách',
        ghiChu: 'Tên trong file: ' + String(c.ten || '').slice(0, 80) + (c.lop ? ' (lớp ' + String(c.lop).slice(0, 20) + ')' : '') };
    }), String(p.tenFile || 'dán từ bảng tính').slice(0, 120));
    logSua(u, x.bai.maBai, '(cả bài)', 'nhập điểm tổng', thay ? 'thay toàn bộ' : 'gộp thêm', clean.length + ' em, ' + cho.length + ' dòng chờ', 'nhập điểm từ nguồn khác');
    var tt = capNhatBaiSauKhiGhi(x.bai, x.cfg, {}, {});
    if (tt === 'Đã công bố') { try { pubTouch(); } catch (e) {} }
    return { ok: true, daGhi: clean.length, cho: cho.length, trangThai: tt };
  } finally { lock.releaseLock(); }
}

/** Sửa điểm một em (bài chỉ có điểm tổng) */
function ngoaiSuaDiem(p, u) {
  var x = ngoaiBai(p, u), sbd = normSbd(p.sbd), d = ngoaiDiem(p.diem, x.thang);
  if (d === null || isNaN(d)) throw new Error('Điểm phải từ 0 đến ' + x.thang);
  if (!String(p.lyDo || '').trim()) throw new Error('Ghi lý do sửa');
  var ix = kqIndex(x.bai.maBai).filter(function (o) { return o.sbd === sbd; })[0];
  if (!ix) throw new Error('Chưa có điểm của em ' + sbd);
  var sh = _kq(), H = kqHeader(sh), cu = sh.getRange(ix.row, H['tong'] + 1).getValues()[0][0];
  sh.getRange(ix.row, H['tong'] + 1).setValue(d);
  sh.getRange(ix.row, H['daSua'] + 1).setValue('TRUE');
  sh.getRange(ix.row, H['capNhat'] + 1).setValue(nowIso());
  logSua(u, x.bai.maBai, sbd, 'tong', cu, d, p.lyDo);
  if (x.bai.trangThai === 'Đã công bố') { try { pubTouch(); } catch (e) {} }
  return { ok: true };
}

/** Ghép dòng chờ xử lý của bài điểm tổng: items = [{id, action:'assign', to} | {id, action:'discard'}] */
function ngoaiGiaiQuyet(p, u, bai) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var pend = cxlLoad(bai.maBai), items = p.items || [];
    if (typeof items === 'string') items = JSON.parse(items);
    var by = {}; ngoaiHocSinh(bai).forEach(function (h) { by[h.sbd] = h; });
    var co = {}; kqIndex(bai.maBai).forEach(function (o) { co[o.sbd] = 1; });
    var errors = [], add = [], del = {}, resolved = 0, discarded = 0;
    items.forEach(function (it) {
      var r = pend.filter(function (x) { return x.id === String(it.id); })[0];
      if (!r) { errors.push('Không tìm thấy dòng chờ'); return; }
      if (it.action === 'discard') { del[r.id] = 1; discarded++; logSua(u, bai.maBai, '', 'bỏ dòng chờ', r.ghiChu + ', ' + r.tong + ' điểm', '(đã bỏ)', 'bỏ dòng chờ xử lý'); return; }
      var to = normSbd(it.to);
      if (!by[to]) { errors.push('SBD ' + it.to + ' không thuộc nhóm của bài'); return; }
      if (co[to]) { errors.push('Em ' + by[to].ten + ' (' + to + ') đã có điểm'); return; }
      if (isNaN(r.tong) || String(r.tong) === '') { errors.push('Dòng "' + r.ghiChu + '" không có điểm hợp lệ'); return; }
      co[to] = 1; del[r.id] = 1; resolved++;
      add.push({ sbd: to, maDe: '', tong: Number(r.tong), p1: '', p2: '', p3: '', tl: {}, daSua: false, nguon: 'ngoai' });
    });
    kqAppendRows(bai.maBai, add);
    cxlDelete(bai.maBai, function (x) { return del[x.id]; });
    if (bai.trangThai === 'Đã công bố') { try { pubTouch(); } catch (e) {} }
    return { ok: true, resolved: resolved, discarded: discarded, errors: errors, conLai: pend.length - resolved - discarded };
  } finally { lock.releaseLock(); }
}


/* ===================== Public.gs ===================== */
/**
 * TRANG TRA ĐIỂM CÔNG KHAI — học sinh/phụ huynh nhập SBD để xem điểm các bài ĐÃ CÔNG BỐ.
 *
 * Tốc độ: mỗi nhóm có một "chỉ mục" gom sẵn điểm đã công bố của nhóm đó (kèm thứ hạng, điểm TB),
 * giữ trong bộ nhớ đệm 1 giờ. Một lượt tra chỉ đọc chỉ mục của đúng nhóm em đó, không đọc cả bảng điểm.
 * Mọi thay đổi ảnh hưởng tới điểm đã công bố gọi pubTouch() để chỉ mục được dựng lại ở lượt tra sau.
 */

function pubVer() { return _props().getProperty('PUB_VER') || '0'; }
function pubTouch() { _props().setProperty('PUB_VER', Date.now().toString(36) + Math.random().toString(36).slice(2, 7)); }

/** Cấu hình hiển thị cho trang công khai */
function publicConfig() {
  var c = CacheService.getScriptCache(), k = 'pubcfg_' + pubVer();
  var hit = c.get(k);
  if (hit) return JSON.parse(hit);
  var out = {
    tenHienThi: String(getSetting('tenHienThi', '') || 'Tra điểm'),
    loiNhan: String(getSetting('loiNhanTraDiem', '') || ''),
    hienThiThuHang: toBool(getSetting('hienThiThuHang', false)),
    hienDiemTB: toBool(getSetting('hienDiemTB', false)),
    choPhepGopY: toBool(getSetting('choPhepGopY', true)),
    hienDiemCaoNhat: toBool(getSetting('hienDiemCaoNhat', true)),
    hienChuyenCan: toBool(getSetting('hienChuyenCan', false))
  };
  c.put(k, JSON.stringify(out), 3600);
  return out;
}

function maxDiemPhan(cfg) {
  return [
    (cfg.p1 && cfg.p1.on) ? Number(cfg.p1.diem) || 0 : 0,
    (cfg.p2 && cfg.p2.on) ? (Number(cfg.p2.soCau) || 0) * (Number(cfg.p2.d4) || 0) : 0,
    (cfg.p3 && cfg.p3.on) ? Number(cfg.p3.diem) || 0 : 0
  ];
}

/** Chỉ mục điểm đã công bố của MỘT nhóm (có bộ nhớ đệm) */
function pubGroupIndex(ma) {
  var key = 'pubg_' + pubVer() + '_' + ma;
  var hit = cacheGetBig(key);
  if (hit) return hit;
  var built = buildPubGroupIndex(ma);
  cachePutBig(key, built, 3600);
  return built;
}

function buildPubGroupIndex(ma) {
  var out = { tenNhom: '', hs: {}, kq: {}, bai: {} };
  var info = getNhomMap()[ma];
  if (!info) return out;
  out.tenNhom = String(info.tenNhom || '');
  try {
    var g = findTab(rosterSnapshot(info.rosterSheetId, false), info.sheetName, ma);
    if (g) {
      out.tenNhom = g.tenNhom;
      g.students.forEach(function (s) { if (s.sbd) out.hs[s.sbd] = { ho: s.ho, ten: s.ten, cu: s.sbdCu || '', tt: s.trangThai }; });
    }
  } catch (e) { /* không đọc được danh sách: vẫn trả điểm, chỉ thiếu tên */ }

  var pub = {};
  sysRows('BaiKiemTra').forEach(function (b) { if (b.trangThai === 'Đã công bố') pub[String(b.maBai)] = b; });
  if (!Object.keys(pub).length) return out;

  var sh = _kq(), last = sh.getLastRow();
  if (last < 2) return out;
  var H = kqHeader(sh);
  var need = Math.max(H['maBai'], H['sbd'], H['maDe'], H['tong'], H['p1'], H['p2'], H['p3']) + 1;
  var v = sh.getRange(2, 1, last - 1, need).getValues(), byBai = {};
  for (var i = 0; i < v.length; i++) {
    var mb = String(v[i][H['maBai']]);
    if (!pub[mb]) continue;
    var sbd = normSbd(v[i][H['sbd']]);
    if (!sbd || sbd.substring(0, 3) !== ma) continue;
    (byBai[mb] = byBai[mb] || []).push({
      sbd: sbd, m: normMaDe(v[i][H['maDe']]), t: Number(v[i][H['tong']]),
      p: [Number(v[i][H['p1']]), Number(v[i][H['p2']]), Number(v[i][H['p3']])]
    });
  }
  Object.keys(byBai).forEach(function (mb) {
    var list = byBai[mb], b = pub[mb], cfg = parseJsonSafe(b.cauHinh, {});
    var sum = 0, cao = 0;
    list.forEach(function (x) { sum += x.t; if (x.t > cao) cao = x.t; });
    out.bai[mb] = {
      ten: String(b.tenBai), ngay: ngayStr(b.ngay), cb: tStr(b.ngayCongBo),
      n: list.length, tb: round2(sum / list.length), cao: round2(cao), mx: maxDiemPhan(cfg),
      loai: cfg.loai === 'ngoai' ? 'ngoai' : '', thang: Number(cfg.thang) || 10, nguon: cfg.nguon || ''
    };
    list.forEach(function (x) {
      var hang = 1;                                   // cùng điểm thì cùng hạng (1, 2, 2, 4…)
      list.forEach(function (y) { if (y.t > x.t + 1e-9) hang++; });
      (out.kq[x.sbd] = out.kq[x.sbd] || []).push({ b: mb, m: x.m, t: x.t, p: x.p, h: hang });
    });
  });
  return out;
}

function pubEntry(e, gi, cfgPub, viaSbd) {
  var b = gi.bai[e.b] || {};
  return {
    maBai: e.b, tenBai: b.ten, ngay: b.ngay, maDe: e.m,
    tong: e.t, p1: e.p[0], p2: e.p[1], p3: e.p[2], mx: b.mx || [0, 0, 0], loai: b.loai || '', thang: b.thang || 10, nguon: b.nguon || '',
    hang: cfgPub.hienThiThuHang ? e.h : null,
    soHs: cfgPub.hienThiThuHang ? b.n : null,
    tbNhom: cfgPub.hienDiemTB ? b.tb : null,
    caoNhat: cfgPub.hienDiemCaoNhat ? (b.cao !== undefined ? b.cao : null) : null,
    sbdCu: viaSbd || null
  };
}

/** Chuỗi SBD của một em: SBD hiện tại + các SBD cũ (em chuyển nhóm) */
function sbdChain(sbd) {
  var chain = [sbd], cur = sbd;
  for (var k = 0; k < 4; k++) {
    var h = pubGroupIndex(cur.substring(0, 3)).hs[cur];
    if (!h || !h.cu || chain.indexOf(h.cu) !== -1) break;
    chain.push(h.cu);
    cur = h.cu;
  }
  return chain;
}

/** Tra điểm theo SBD (không cần đăng nhập) */
function traDiem(p) {
  var raw = String(p.sbd == null ? '' : p.sbd).replace(/\s/g, '');
  if (!/^\d{6}$/.test(raw)) throw new Error('Số báo danh gồm đúng 6 chữ số');
  var sbd = raw;
  var cfgPub = publicConfig();
  var gi = pubGroupIndex(sbd.substring(0, 3));
  var hs = gi.hs[sbd];
  var chain = sbdChain(sbd), list = [];
  chain.forEach(function (s, i) {
    var g = i === 0 ? gi : pubGroupIndex(s.substring(0, 3));
    (g.kq[s] || []).forEach(function (e) { list.push(pubEntry(e, g, cfgPub, i === 0 ? null : s)); });
  });
  if (!hs && !list.length) return { found: false, sbd: sbd };
  list.sort(function (a, b) { return String(b.ngay || '') < String(a.ngay || '') ? -1 : (String(b.ngay || '') > String(a.ngay || '') ? 1 : 0); });
  return {
    found: true, sbd: sbd,
    ho: hs ? hs.ho : '', ten: hs ? hs.ten : '', tenNhom: gi.tenNhom,
    daChuyen: hs ? hs.tt === 'Đã chuyển' : false,
    hienThiThuHang: cfgPub.hienThiThuHang, hienDiemTB: cfgPub.hienDiemTB, hienDiemCaoNhat: cfgPub.hienDiemCaoNhat,
    chuyenCan: cfgPub.hienChuyenCan ? pubChuyenCan(sbd) : null,
    bai: list
  };
}

/** Chi tiết một bài của một em: từng câu chọn gì, đáp án đúng, đúng/sai */
function chiTietBai(p) {
  var raw = String(p.sbd == null ? '' : p.sbd).replace(/\s/g, '');
  if (!/^\d{6}$/.test(raw)) throw new Error('Số báo danh gồm đúng 6 chữ số');
  var sbd = raw;
  var bai = findBai(p.maBai);
  if (!bai || bai.trangThai !== 'Đã công bố') throw new Error('Bài này chưa được công bố');
  if (laBaiNgoai(bai)) throw new Error('Bài này chỉ có điểm tổng, không có chi tiết từng câu');
  var chain = sbdChain(sbd);
  var r = kqLoad(bai.maBai).filter(function (x) { return chain.indexOf(x.sbd) !== -1; })[0];
  if (!r) throw new Error('Không có kết quả của số báo danh này trong bài');
  var cfg = parseJsonSafe(bai.cauHinh, {});
  var key = getBaiDapAn(bai.maBai)[r.maDe] || null;
  var sc = (key && r.tl) ? chamBai(cfg, r.tl, key) : null;
  return {
    tenBai: String(bai.tenBai), ngay: ngayStr(bai.ngay), maDe: r.maDe,
    tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, mx: maxDiemPhan(cfg),
    cfg: {
      p1: cfg.p1 && cfg.p1.on ? { soCau: cfg.p1.soCau, diem: cfg.p1.diem } : null,
      p2: cfg.p2 && cfg.p2.on ? { soCau: cfg.p2.soCau, d: [cfg.p2.d1, cfg.p2.d2, cfg.p2.d3, cfg.p2.d4] } : null,
      p3: cfg.p3 && cfg.p3.on ? { soCau: cfg.p3.soCau, diem: cfg.p3.diem } : null
    },
    tl: r.tl || null, key: key, dung: sc ? sc.dung : null, daSua: r.daSua
  };
}

/**
 * Góp ý / báo lỗi. Công khai (u = null) hoặc từ giáo viên đã đăng nhập.
 * Chống spam: ô ẩn "website" (máy tự điền thì bỏ), tối đa 30 góp ý công khai mỗi 10 phút.
 */
function guiGopY(p, u) {
  if (p.website) return { ok: true };
  var nd = String(p.noiDung || '').trim();
  if (nd.length < 5) throw new Error('Nội dung góp ý quá ngắn (ít nhất 5 ký tự)');
  if (nd.length > 2000) nd = nd.slice(0, 2000);
  if (!u && !toBool(getSetting('choPhepGopY', true))) throw new Error('Hiện chưa nhận góp ý qua trang này');
  var c = CacheService.getScriptCache(), n = Number(c.get('gopy_rate')) || 0;
  if (!u && n >= 30) throw new Error('Đang có quá nhiều góp ý, hãy thử lại sau ít phút');
  c.put('gopy_rate', String(n + 1), 600);

  var sh = getSystemSheet().getSheetByName('GopY');
  var H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), nCol = sh.getLastColumn();
  var row = []; for (var i = 0; i < nCol; i++) row.push('');
  var set = function (k, v) { if (H[k] !== undefined) row[H[k]] = v; };
  set('id', 'G' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5));
  set('thoiGian', nowIso());
  set('noiDung', nd);
  set('lienHe', String(p.lienHe || '').trim().slice(0, 200));
  set('daXuLy', '');
  set('nguon', u ? 'giáo viên' : 'trang tra điểm');
  set('nguoiGui', u ? u.username : '');
  set('sbd', normSbd(p.sbd) || '');
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, nCol).setNumberFormat('@');
  sh.getRange(r, 1, 1, nCol).setValues([row]);
  sysTouch('GopY');

  try {
    var to = String(getSetting('emailNhanThongBao', '') || '').trim();
    if (to) MailApp.sendEmail({ to: to, subject: '[Điểm] Góp ý mới' + (u ? ' từ ' + u.username : ''), name: 'Hệ thống điểm',
      body: nd + '\n\n' + (p.lienHe ? 'Liên hệ: ' + p.lienHe + '\n' : '') + (p.sbd ? 'SBD: ' + p.sbd + '\n' : '') + 'Nguồn: ' + (u ? 'giáo viên ' + u.username : 'trang tra điểm') });
  } catch (e) {}
  return { ok: true };
}


/* ===================== ThongKe.gs ===================== */
/** THỐNG KÊ MỘT BÀI — để giáo viên biết chữa câu nào, bổ túc cho em nào */

function _median(arr) {
  if (!arr.length) return null;
  var a = arr.slice().sort(function (x, y) { return x - y; }), m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : round2((a[m - 1] + a[m]) / 2);
}

function _tomTat(list) {
  var n = list.length, out = { n: n, tb: null, trungVi: null, cao: null, thap: null, p1: null, p2: null, p3: null, pho: [] };
  for (var k = 0; k <= 10; k++) out.pho.push(0);
  if (!n) return out;
  var t = list.map(function (r) { return r.tong; }), s = 0, s1 = 0, s2 = 0, s3 = 0;
  list.forEach(function (r) {
    s += r.tong; s1 += r.p1; s2 += r.p2; s3 += r.p3;
    var bin = Math.min(10, Math.max(0, Math.floor(r.tong + 1e-9)));
    out.pho[bin]++;
  });
  out.tb = round2(s / n); out.trungVi = _median(t);
  out.cao = Math.max.apply(null, t); out.thap = Math.min.apply(null, t);
  out.p1 = round2(s1 / n); out.p2 = round2(s2 / n); out.p3 = round2(s3 / n);
  return out;
}

function thongKeBai(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  if (parseJsonSafe(bai.cauHinh, {}).loai === 'ngoai') throw new Error('Bài chỉ có điểm tổng: xem phổ điểm ở trang Kết quả');
  return cachedResult('tk_' + bai.maBai + '_' + u.username, 900, function () { return thongKeBai_(p, u, bai); }, p.fresh);
}
function thongKeBai_(p, u, bai) {
  var cfg = parseJsonSafe(bai.cauHinh, {});
  var mx = maxDiemPhan(cfg);
  var res = kqLoad(bai.maBai), keys = getBaiDapAn(bai.maBai);
  var groups = loadAccessibleGroups(u, false);
  var nguong = Number(getSetting('nguongBoTuc', 5)) || 5;

  // ---- tổng quan, theo nhóm
  var byGroup = {};
  res.forEach(function (r) { var ma = r.sbd.substring(0, 3); (byGroup[ma] = byGroup[ma] || []).push(r); });
  var maList = parseNhomList(bai.cacNhom);
  Object.keys(byGroup).forEach(function (m) { if (maList.indexOf(m) === -1) maList.push(m); });
  maList.sort();
  var names = {};
  var nhomOut = maList.map(function (ma) {
    var g = groups[ma], list = byGroup[ma] || [];
    var dangHoc = g ? g.students.filter(function (s) { return s.trangThai === 'Đang học'; }) : [];
    if (g) g.students.forEach(function (s) { names[s.sbd] = (s.ho + ' ' + s.ten).trim(); });
    var co = {};
    list.forEach(function (r) { co[r.sbd] = true; });
    var t = _tomTat(list);
    t.maNhom = ma; t.tenNhom = g ? g.tenNhom : ma;
    t.vang = dangHoc.filter(function (s) { return !co[s.sbd]; }).length;
    return t;
  });

  // ---- theo câu, tách theo mã đề (các mã đã trộn nên câu 1 mã này khác câu 1 mã kia)
  var theoMa = {}, coChiTiet = 0;
  res.forEach(function (r) {
    var k = keys[r.maDe];
    if (!k || !r.tl) return;
    var sc = chamBai(cfg, r.tl, k);
    if (sc.thieuKey) return;
    coChiTiet++;
    var M = theoMa[r.maDe];
    if (!M) {
      M = theoMa[r.maDe] = { maDe: r.maDe, n: 0, p1: [], p2: [], p3: [] };
      (k.p1 || []).forEach(function (x, i) { if (cfg.p1 && cfg.p1.on && i < cfg.p1.soCau) M.p1.push({ cau: i + 1, dapAn: x, dung: 0, chon: { A: 0, B: 0, C: 0, D: 0, '': 0 } }); });
      if (cfg.p2 && cfg.p2.on) for (var j = 0; j < cfg.p2.soCau; j++) M.p2.push({ cau: j + 1, dapAn: normP2((k.p2 || [])[j]), y: [0, 0, 0, 0], du4: 0 });
      if (cfg.p3 && cfg.p3.on) for (var q = 0; q < cfg.p3.soCau; q++) M.p3.push({ cau: q + 1, dapAn: normP3((k.p3 || [])[q]), dung: 0, sai: {} });
    }
    M.n++;
    M.p1.forEach(function (c, i) {
      var a = String((r.tl.p1 || [])[i] || '').toUpperCase();
      c.chon[/^[A-D]$/.test(a) ? a : '']++;
      if (sc.dung.p1[i] === 1) c.dung++;
    });
    M.p2.forEach(function (c, i) {
      var a = normP2((r.tl.p2 || [])[i]);
      for (var y = 0; y < 4; y++) if (a.charAt(y) === c.dapAn.charAt(y)) c.y[y]++;
      if (sc.dung.p2[i] === 4) c.du4++;
    });
    M.p3.forEach(function (c, i) {
      if (sc.dung.p3[i] === 1) c.dung++;
      else { var a = normP3((r.tl.p3 || [])[i]) || '(bỏ trống)'; c.sai[a] = (c.sai[a] || 0) + 1; }
    });
  });
  var pct = function (x, n) { return n ? Math.round(x * 1000 / n) / 10 : 0; };
  var khoNhat = [];
  var maOut = Object.keys(theoMa).sort().map(function (md) {
    var M = theoMa[md];
    var o = {
      maDe: md, n: M.n,
      p1: M.p1.map(function (c) {
        var tl = pct(c.dung, M.n);
        khoNhat.push({ maDe: md, phan: 'I', cau: c.cau, tiLe: tl });
        var nhieu = ['A', 'B', 'C', 'D'].filter(function (x) { return x !== c.dapAn; }).sort(function (a, b) { return c.chon[b] - c.chon[a]; })[0];
        return { cau: c.cau, dapAn: c.dapAn, tiLe: tl, chon: c.chon, nhieuNhat: c.chon[nhieu] ? nhieu : null };
      }),
      p2: M.p2.map(function (c) {
        var tl = pct(c.du4, M.n);
        khoNhat.push({ maDe: md, phan: 'II', cau: c.cau, tiLe: tl });
        return { cau: c.cau, dapAn: c.dapAn, tiLeDu4: tl, y: c.y.map(function (x) { return pct(x, M.n); }) };
      }),
      p3: M.p3.map(function (c) {
        var tl = pct(c.dung, M.n);
        khoNhat.push({ maDe: md, phan: 'III', cau: c.cau, tiLe: tl });
        var sai = Object.keys(c.sai).map(function (k) { return { gt: k, n: c.sai[k] }; }).sort(function (a, b) { return b.n - a.n; }).slice(0, 3);
        return { cau: c.cau, dapAn: c.dapAn, tiLe: tl, saiHay: sai };
      })
    };
    return o;
  });
  khoNhat.sort(function (a, b) { return a.tiLe - b.tiLe; });

  // ---- cần bổ túc
  var boTuc = [];
  res.forEach(function (r) {
    var ly = [];
    if (r.tong < nguong) ly.push('dưới ' + nguong + ' điểm');
    ['p1', 'p2', 'p3'].forEach(function (k, i) { if (mx[i] > 0 && r[k] < mx[i] / 2) ly.push('Phần ' + ['I', 'II', 'III'][i] + ' dưới một nửa'); });
    if (ly.length) boTuc.push({ sbd: r.sbd, ten: names[r.sbd] || '', tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, lyDo: ly });
  });
  boTuc.sort(function (a, b) { return a.tong - b.tong; });

  // ---- so với các bài trước của cùng nhóm (tối đa 6 bài gần nhất)
  var cungNhom = sysRows('BaiKiemTra').filter(function (b) {
    if (u.role !== 'admin' && b.teacher !== u.username) return false;
    var ms = parseNhomList(b.cacNhom);
    return ms.some(function (m) { return maList.indexOf(m) !== -1; });
  });
  cungNhom.sort(function (a, b) { var x = ngayStr(a.ngay) || String(a.createdAt), y = ngayStr(b.ngay) || String(b.createdAt); return x < y ? -1 : (x > y ? 1 : 0); });
  var vt = cungNhom.map(function (b) { return String(b.maBai); }).indexOf(String(bai.maBai));
  var chon = cungNhom.slice(Math.max(0, vt - 5), vt + 1);
  var chonSet = {};
  chon.forEach(function (b) { chonSet[String(b.maBai)] = { maBai: String(b.maBai), tenBai: String(b.tenBai), ngay: ngayStr(b.ngay), nhom: {} }; });
  var shk = _kq(), last = shk.getLastRow();
  if (last >= 2 && chon.length > 1) {
    var H = kqHeader(shk);
    var v = shk.getRange(2, 1, last - 1, Math.max(H['maBai'], H['sbd'], H['tong']) + 1).getValues();
    v.forEach(function (row) {
      var o = chonSet[String(row[H['maBai']])];
      if (!o) return;
      var ma = normSbd(row[H['sbd']]).substring(0, 3);
      if (maList.indexOf(ma) === -1) return;
      var g = o.nhom[ma] = o.nhom[ma] || { s: 0, n: 0 };
      g.s += Number(row[H['tong']]); g.n++;
    });
  } else if (chon.length === 1) {
    res.forEach(function (r) { var ma = r.sbd.substring(0, 3); var g = chonSet[String(bai.maBai)].nhom[ma] = chonSet[String(bai.maBai)].nhom[ma] || { s: 0, n: 0 }; g.s += r.tong; g.n++; });
  }
  var xuHuong = chon.map(function (b) {
    var o = chonSet[String(b.maBai)], nh = {};
    Object.keys(o.nhom).forEach(function (m) { nh[m] = o.nhom[m].n ? round2(o.nhom[m].s / o.nhom[m].n) : null; });
    return { maBai: o.maBai, tenBai: o.tenBai, ngay: o.ngay, laBaiNay: o.maBai === String(bai.maBai), tbNhom: nh };
  }).filter(function (x) { return Object.keys(x.tbNhom).length; });

  return {
    bai: { maBai: bai.maBai, tenBai: String(bai.tenBai), ngay: ngayStr(bai.ngay), trangThai: bai.trangThai },
    mx: mx, nguong: nguong,
    tongQuan: _tomTat(res), nhom: nhomOut,
    coChiTiet: coChiTiet, theoMa: maOut, khoNhat: khoNhat.slice(0, 8),
    boTuc: boTuc, xuHuong: xuHuong
  };
}


/* ===================== DiemDanh.gs ===================== */
/**
 * ĐIỂM DANH, CHUYÊN CẦN
 *
 * Mỗi buổi học là MỘT dòng trong tab BuoiHoc; trạng thái cả lớp gói trong một ô JSON {sbd: mã}.
 * Mã: C = có mặt, M = đi muộn, P = vắng có phép, K = vắng không phép.
 *
 * Chống bao che:
 *  - Lớp trưởng chỉ điểm danh được buổi HÔM NAY của nhóm mình; qua ngày hoặc cô chốt thì khóa.
 *  - Mọi thay đổi ghi vào DiemDanhLog (ai, lúc nào, từ gì sang gì); sửa sau khi đã gửi bị gắn cờ.
 *  - Đối chiếu với bài kiểm tra cùng ngày (vắng mà có bài làm / có mặt mà không có bài làm).
 *  - Sau khi khóa: chỉ sửa qua "yêu cầu sửa" do cô duyệt, hoặc cô sửa thẳng kèm lý do.
 */

var DD_CODES = { C: 'Có mặt', M: 'Đi muộn', P: 'Vắng có phép', K: 'Vắng không phép' };

function homNay() { return Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd'); }
function ddSheet() { return getSystemSheet().getSheetByName('BuoiHoc'); }

function ddRows(maNhom) {
  return sysRows('BuoiHoc').filter(function (r) { return !maNhom || normMa(r.maNhom) === normMa(maNhom); }).map(ddObj);
}
function ddObj(r) {
  return {
    _row: r._row, maBuoi: String(r.maBuoi), maNhom: normMa(r.maNhom), ngay: ngayStr(r.ngay), trangThai: String(r.trangThai || 'Đang mở'),
    dd: parseJsonSafe(r.dd, {}) || {}, ghiChu: parseJsonSafe(r.ghiChu, {}) || {},
    nguoiGhi: String(r.nguoiGhi || ''), capNhat: tStr(r.capNhat), nopBoi: String(r.nopBoi || ''), nopLuc: tStr(r.nopLuc),
    chotBoi: String(r.chotBoi || ''), chotLuc: tStr(r.chotLuc), taoBoi: String(r.taoBoi || '')
  };
}
function ddFind(maBuoi) { return ddRows().filter(function (b) { return b.maBuoi === String(maBuoi); })[0] || null; }

/** Ghi lại một buổi (cả dòng), giữ dạng chữ để Sheets không tự đổi kiểu */
function ddWrite(b) {
  var sh = ddSheet(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), n = sh.getLastColumn();
  var row = []; for (var i = 0; i < n; i++) row.push('');
  var set = function (k, v) { if (H[k] !== undefined) row[H[k]] = v; };
  set('maBuoi', b.maBuoi); set('maNhom', b.maNhom); set('ngay', b.ngay); set('trangThai', b.trangThai);
  set('dd', JSON.stringify(b.dd || {})); set('ghiChu', JSON.stringify(b.ghiChu || {}));
  set('nguoiGhi', b.nguoiGhi || ''); set('capNhat', nowIso()); set('nopBoi', b.nopBoi || ''); set('nopLuc', b.nopLuc || '');
  set('chotBoi', b.chotBoi || ''); set('chotLuc', b.chotLuc || ''); set('taoBoi', b.taoBoi || ''); set('taoLuc', b.taoLuc || nowIso());
  var r = b._row || (sh.getLastRow() + 1);
  sh.getRange(r, 1, 1, n).setNumberFormat('@');
  sh.getRange(r, 1, 1, n).setValues([row]);
  b._row = r;
  sysTouch('BuoiHoc');
  try { pubTouch(); } catch (e) {}
  return b;
}

function ddLog(entries) {
  if (!entries.length) return;
  var sh = getSystemSheet().getSheetByName('DiemDanhLog'), n = 8, r0 = sh.getLastRow() + 1;
  sh.getRange(r0, 1, entries.length, n).setNumberFormat('@');
  sh.getRange(r0, 1, entries.length, n).setValues(entries.map(function (e) {
    return [nowIso(), e.maBuoi, e.sbd, e.cu || '', e.moi || '', e.nguoi, e.loai || 'ghi', e.lyDo || ''];
  }));
  sysTouch('DiemDanhLog');
}

/* ---------------- quyền ---------------- */

/** Thông tin lớp trưởng: nhóm phụ trách (đọc từ tài khoản) */
function ltInfo(u) {
  var r = findUserRow(u.username);
  if (!r) throw new Error('Không tìm thấy tài khoản');
  var ma = normMa(userField(r, 'maNhom'));
  if (!ma) throw new Error('Tài khoản lớp trưởng chưa được gán nhóm');
  return { maNhom: ma, quanLy: String(userField(r, 'quanLy') || '') };
}
function isStaff(u) { return u.role === 'admin' || u.role === 'teacher'; }

/** Kiểm tra quyền với một nhóm; trả về {laLopTruong} */
function ddAccess(u, maNhom) {
  maNhom = normMa(maNhom);
  if (u.role === 'loptruong') {
    if (ltInfo(u).maNhom !== maNhom) throw new Error('Không có quyền với nhóm này');
    return { laLopTruong: true };
  }
  requireGroupAccess(u, maNhom);
  return { laLopTruong: false };
}

/** Danh sách học sinh đang học của nhóm (ảnh chụp danh sách, đọc mới nếu cần) */
function ddStudents(maNhom, fresh) {
  var reg = getNhomMap()[normMa(maNhom)];
  if (!reg) throw new Error('Nhóm ' + maNhom + ' chưa đăng ký');
  var g = findTab(rosterSnapshot(reg.rosterSheetId, !!fresh), reg.sheetName, normMa(maNhom));
  if (!g) throw new Error('Không đọc được danh sách nhóm ' + maNhom);
  return { tenNhom: g.tenNhom, students: g.students.filter(function (s) { return s.sbd && s.trangThai === 'Đang học'; }).map(function (s) { return { sbd: s.sbd, ho: s.ho, ten: s.ten }; }) };
}

/** Lớp trưởng chỉ sửa được buổi hôm nay chưa chốt */
function ddCanEdit(u, b) {
  if (isStaff(u)) return { ok: true, sauChot: b.trangThai === 'Đã chốt' || b.ngay < homNay() };
  if (b.ngay !== homNay()) return { ok: false, lyDo: 'Chỉ điểm danh được buổi của hôm nay. Muốn sửa buổi cũ, hãy gửi yêu cầu sửa cho cô.' };
  if (b.trangThai === 'Đã chốt') return { ok: false, lyDo: 'Cô đã chốt buổi này. Muốn sửa, hãy gửi yêu cầu sửa cho cô.' };
  return { ok: true, sauChot: false };
}

/* ---------------- buổi học ---------------- */

/** Mở (hoặc tạo) buổi của một ngày. Lớp trưởng: chỉ hôm nay, chỉ nhóm mình */
function ddMoBuoi(p, u) {
  var ma = u.role === 'loptruong' ? ltInfo(u).maNhom : normMa(p.maNhom);
  ddAccess(u, ma);
  var ngay = u.role === 'loptruong' ? homNay() : (String(p.ngay || '') || homNay());
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ngay)) throw new Error('Ngày không hợp lệ');
  if (ngay > homNay()) throw new Error('Không điểm danh trước cho ngày chưa tới');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var id = p.maBuoi ? String(p.maBuoi) : ma + '-' + ngay;
    var b = ddFind(id);
    if (!b) {
      if (p.maBuoi && !isStaff(u)) throw new Error('Không tìm thấy buổi học');
      b = ddWrite({ maBuoi: id, maNhom: ma, ngay: ngay, trangThai: 'Đang mở', dd: {}, ghiChu: {}, taoBoi: u.username, taoLuc: nowIso() });
    } else { ddAccess(u, b.maNhom); _KHONG_DOI = true; }
    return ddView(b, u, false);
  } finally { lock.releaseLock(); }
}

/** Dữ liệu một buổi để hiện màn điểm danh */
function ddView(b, u, withLog) {
  var st = ddStudents(b.maNhom, false);
  var edit = ddCanEdit(u, b);
  var out = {
    buoi: { maBuoi: b.maBuoi, maNhom: b.maNhom, ngay: b.ngay, trangThai: b.trangThai, nguoiGhi: b.nguoiGhi, capNhat: b.capNhat, nopBoi: b.nopBoi, nopLuc: b.nopLuc, chotBoi: b.chotBoi, chotLuc: b.chotLuc, homNay: b.ngay === homNay() },
    tenNhom: st.tenNhom, students: st.students, dd: b.dd, ghiChu: b.ghiChu,
    suaDuoc: edit.ok, lyDoKhoa: edit.lyDo || '', laLopTruong: u.role === 'loptruong'
  };
  if (withLog && isStaff(u)) out.nhatKy = sysRows('DiemDanhLog').filter(function (r) { return String(r.maBuoi) === b.maBuoi; }).reverse().slice(0, 300).map(function (r) {
    return { luc: tStr(r.thoiGian), sbd: normSbd(r.sbd) || String(r.sbd), cu: String(r.cu), moi: String(r.moi), nguoi: String(r.nguoi), loai: String(r.loai), lyDo: String(r.lyDo || '') };
  });
  return out;
}

function ddXemBuoi(p, u) {
  var b = ddFind(p.maBuoi);
  if (!b) throw new Error('Không tìm thấy buổi học');
  ddAccess(u, b.maNhom);
  return ddView(b, u, true);
}

/**
 * Lưu thay đổi điểm danh: p.ghi = [{sbd, tt: 'C'|'M'|'P'|'K', ghiChu?}].
 * Lớp trưởng: chỉ buổi hôm nay chưa chốt. Cô sửa buổi đã khóa phải ghi lý do.
 */
function ddLuu(p, u) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var b = ddFind(p.maBuoi);
    if (!b) throw new Error('Không tìm thấy buổi học');
    ddAccess(u, b.maNhom);
    var edit = ddCanEdit(u, b);
    if (!edit.ok) throw new Error(edit.lyDo);
    var lyDo = String(p.lyDo || '').trim();
    if (edit.sauChot && !lyDo) throw new Error('Buổi này đã khóa: cần ghi lý do khi sửa');
    var hop = {}; ddStudents(b.maNhom, false).students.forEach(function (s) { hop[s.sbd] = true; });
    var ghi = p.ghi; if (typeof ghi === 'string') ghi = JSON.parse(ghi);
    var logs = [], loai = edit.sauChot ? 'sửa sau chốt' : (b.trangThai === 'Đã nộp' && u.role === 'loptruong' ? 'sửa sau khi gửi' : 'ghi');
    (ghi || []).forEach(function (g) {
      var sbd = normSbd(g.sbd), tt = String(g.tt || '').toUpperCase();
      if (!hop[sbd] && !b.dd[sbd]) return;                     // không thuộc nhóm: bỏ qua
      if (!DD_CODES[tt]) return;
      var cu = b.dd[sbd] || 'C';
      if (g.ghiChu !== undefined) { var nt = String(g.ghiChu || '').trim().slice(0, 120); if (nt) b.ghiChu[sbd] = nt; else delete b.ghiChu[sbd]; }
      if (cu === tt && b.dd[sbd]) return;
      b.dd[sbd] = tt;
      if (cu !== tt) logs.push({ maBuoi: b.maBuoi, sbd: sbd, cu: cu, moi: tt, nguoi: u.username, loai: loai, lyDo: lyDo });
    });
    // em chưa được chạm vào: mặc định có mặt (ghi rõ để về sau không lệch khi danh sách đổi)
    Object.keys(hop).forEach(function (s) { if (!b.dd[s]) b.dd[s] = 'C'; });
    b.nguoiGhi = u.username;
    ddWrite(b);
    ddLog(logs);
    return { ok: true, doi: logs.length, trangThai: b.trangThai, capNhat: nowIso() };
  } finally { lock.releaseLock(); }
}

/** Lớp trưởng (hoặc cô) gửi điểm danh: đánh dấu đã nộp */
function ddNop(p, u) {
  var b = ddFind(p.maBuoi);
  if (!b) throw new Error('Không tìm thấy buổi học');
  ddAccess(u, b.maNhom);
  var edit = ddCanEdit(u, b);
  if (!edit.ok) throw new Error(edit.lyDo);
  if (b.trangThai !== 'Đã chốt') b.trangThai = 'Đã nộp';
  b.nopBoi = u.username; b.nopLuc = nowIso();
  ddWrite(b);
  ddLog([{ maBuoi: b.maBuoi, sbd: '(buổi)', cu: '', moi: 'đã gửi', nguoi: u.username, loai: 'gửi' }]);
  return { ok: true, trangThai: b.trangThai };
}

/** Cô chốt / mở lại buổi */
function ddChot(p, u) {
  if (!isStaff(u)) throw new Error('Không có quyền');
  var b = ddFind(p.maBuoi);
  if (!b) throw new Error('Không tìm thấy buổi học');
  ddAccess(u, b.maNhom);
  b.trangThai = p.chot ? 'Đã chốt' : (b.nopBoi ? 'Đã nộp' : 'Đang mở');
  b.chotBoi = p.chot ? u.username : ''; b.chotLuc = p.chot ? nowIso() : '';
  ddWrite(b);
  ddLog([{ maBuoi: b.maBuoi, sbd: '(buổi)', cu: '', moi: p.chot ? 'chốt' : 'mở lại', nguoi: u.username, loai: p.chot ? 'chốt' : 'mở lại' }]);
  return { ok: true, trangThai: b.trangThai };
}

function ddXoaBuoi(p, u) {
  if (!isStaff(u)) throw new Error('Không có quyền');
  var b = ddFind(p.maBuoi);
  if (!b) throw new Error('Không tìm thấy buổi học');
  ddAccess(u, b.maNhom);
  if (!String(p.lyDo || '').trim()) throw new Error('Ghi lý do xóa buổi');
  deleteRowsList(ddSheet(), [b._row]);
  ddLog([{ maBuoi: b.maBuoi, sbd: '(buổi)', cu: JSON.stringify(b.dd).slice(0, 200), moi: '(đã xóa)', nguoi: u.username, loai: 'xóa buổi', lyDo: p.lyDo }]);
  try { pubTouch(); } catch (e) {}
  return { ok: true };
}

/* ---------------- yêu cầu sửa (sau khi khóa) ---------------- */

function ddYcSheet() { return getSystemSheet().getSheetByName('YeuCauSua'); }

/** Tạo yêu cầu sửa. Từ lớp trưởng/cô (u) hoặc công khai từ phụ huynh/học sinh (u = null, có SBD) */
function ddYeuCau(p, u) {
  if (!u && p.website) return { ok: true };
  var b = ddFind(p.maBuoi);
  if (!b) throw new Error('Không tìm thấy buổi học');
  var sbd = normSbd(p.sbd), moi = String(p.moi || '').toUpperCase(), lyDo = String(p.lyDo || '').trim();
  if (!DD_CODES[moi]) throw new Error('Chọn trạng thái muốn sửa thành');
  if (lyDo.length < 5) throw new Error('Ghi rõ lý do (ít nhất 5 ký tự)');
  if (u) ddAccess(u, b.maNhom);
  else {
    if (!toBool(getSetting('hienChuyenCan', false))) throw new Error('Chưa mở tính năng này');
    var c = CacheService.getScriptCache(), n = Number(c.get('ddyc_rate')) || 0;
    if (n >= 30) throw new Error('Đang có quá nhiều yêu cầu, thử lại sau ít phút');
    c.put('ddyc_rate', String(n + 1), 600);
  }
  if (!b.dd[sbd] && sbd.substring(0, 3) !== b.maNhom) throw new Error('Số báo danh không thuộc buổi học này');
  var cu = b.dd[sbd] || 'C';
  if (cu === moi) throw new Error('Trạng thái hiện tại đã là "' + DD_CODES[moi] + '"');
  var sh = ddYcSheet(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), nC = sh.getLastColumn();
  var row = []; for (var i = 0; i < nC; i++) row.push('');
  var set = function (k, v) { if (H[k] !== undefined) row[H[k]] = v; };
  set('id', 'Y' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5)); set('thoiGian', nowIso());
  set('maBuoi', b.maBuoi); set('maNhom', b.maNhom); set('sbd', sbd); set('cu', cu); set('moi', moi); set('lyDo', lyDo.slice(0, 500));
  set('nguoiGui', u ? u.username : 'phụ huynh/học sinh'); set('lienHe', String(p.lienHe || '').slice(0, 120)); set('trangThai', 'Chờ duyệt');
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, nC).setNumberFormat('@');
  sh.getRange(r, 1, 1, nC).setValues([row]);
  sysTouch('YeuCauSua');
  return { ok: true };
}

function ddYeuCauList(u, maNhom) {
  var nhom = getNhomMap();
  return sysRows('YeuCauSua').filter(function (r) {
    var ma = normMa(r.maNhom), info = nhom[ma];
    if (maNhom && ma !== normMa(maNhom)) return false;
    return u.role === 'admin' || (info && info.teacher === u.username);
  }).map(function (r) {
    return { id: String(r.id), luc: tStr(r.thoiGian), maBuoi: String(r.maBuoi), maNhom: normMa(r.maNhom), sbd: normSbd(r.sbd), cu: String(r.cu), moi: String(r.moi),
      lyDo: String(r.lyDo || ''), nguoiGui: String(r.nguoiGui || ''), lienHe: String(r.lienHe || ''), trangThai: String(r.trangThai || ''), duyetBoi: String(r.duyetBoi || ''), _row: r._row };
  }).reverse();
}

/** Cô duyệt / từ chối yêu cầu sửa */
function ddDuyet(p, u) {
  if (!isStaff(u)) throw new Error('Không có quyền');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var y = ddYeuCauList(u).filter(function (x) { return x.id === String(p.id); })[0];
    if (!y) throw new Error('Không tìm thấy yêu cầu');
    if (y.trangThai !== 'Chờ duyệt') throw new Error('Yêu cầu này đã được xử lý');
    var sh = ddYcSheet(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
    if (p.dongY) {
      var b = ddFind(y.maBuoi);
      if (!b) throw new Error('Buổi học đã bị xóa');
      var cu = b.dd[y.sbd] || 'C';
      b.dd[y.sbd] = y.moi; ddWrite(b);
      ddLog([{ maBuoi: b.maBuoi, sbd: y.sbd, cu: cu, moi: y.moi, nguoi: u.username, loai: 'duyệt yêu cầu', lyDo: y.lyDo + ' (gửi bởi ' + y.nguoiGui + ')' }]);
    }
    setTextCell(sh, y._row, H['trangThai'] + 1, p.dongY ? 'Đã duyệt' : 'Từ chối');
    setTextCell(sh, y._row, H['duyetBoi'] + 1, u.username);
    setTextCell(sh, y._row, H['duyetLuc'] + 1, nowIso());
    if (H['phanHoi'] !== undefined) setTextCell(sh, y._row, H['phanHoi'] + 1, String(p.phanHoi || ''));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------------- tổng quan, phân tích cho cô ---------------- */

function ddDem(dd, list) {
  var c = { C: 0, M: 0, P: 0, K: 0 };
  list.forEach(function (s) { c[dd[s] || 'C']++; });
  return c;
}

/** Màn điểm danh của cô: buổi hôm nay từng nhóm + số yêu cầu chờ duyệt */
function ddTongQuan(p, u) {
  if (u.role === 'loptruong') {
    var me = ltInfo(u), st = ddStudents(me.maNhom, false), rows = ddRows(me.maNhom);
    rows.sort(function (a, b) { return a.ngay < b.ngay ? 1 : -1; });
    return { laLopTruong: true, maNhom: me.maNhom, tenNhom: st.tenNhom, homNay: homNay(),
      ganDay: rows.slice(0, 12).map(function (b) { var c = ddDem(b.dd, st.students.map(function (s) { return s.sbd; })); return { maBuoi: b.maBuoi, ngay: b.ngay, trangThai: b.trangThai, dem: c }; }) };
  }
  if (!isStaff(u)) throw new Error('Không có quyền');
  return cachedResult('ddtq_' + u.username + '_' + homNay(), 900, function () { return ddTongQuan_(u); }, p.fresh);
}
function ddTongQuan_(u) {
  var groups = listGroups(u).groups, today = homNay(), rowsAll = ddRows();
  var yc = ddYeuCauList(u).filter(function (y) { return y.trangThai === 'Chờ duyệt'; });
  return {
    homNay: today,
    nhom: groups.map(function (g) {
      var b = rowsAll.filter(function (x) { return x.maNhom === g.maNhom && x.ngay === today; })[0];
      return { maNhom: g.maNhom, tenNhom: g.tenNhom, siSo: g.active, homNay: b ? { maBuoi: b.maBuoi, trangThai: b.trangThai, nopBoi: b.nopBoi, nopLuc: b.nopLuc, dem: ddDem(b.dd, Object.keys(b.dd)) } : null,
        soBuoi: rowsAll.filter(function (x) { return x.maNhom === g.maNhom; }).length, choDuyet: yc.filter(function (y) { return y.maNhom === g.maNhom; }).length };
    }),
    choDuyet: yc.length
  };
}

/** Phân tích chuyên cần một nhóm: theo buổi, theo học sinh, cờ cảnh báo, yêu cầu chờ duyệt */
function ddPhanTich(p, u) {
  if (!isStaff(u)) throw new Error('Không có quyền');
  var ma = normMa(p.maNhom);
  ddAccess(u, ma);
  return cachedResult('ddpt_' + ma + '_' + u.username + '_' + homNay(), 900, function () { return ddPhanTich_(ma, u); }, p.fresh);
}
function ddPhanTich_(ma, u) {
  var nguong = Number(getSetting('nguongChuyenCan', 80)) || 80;
  var st = ddStudents(ma, false), list = st.students;
  var rows = ddRows(ma).sort(function (a, b) { return a.ngay < b.ngay ? -1 : 1; });
  var buoi = rows.map(function (b) {
    var c = ddDem(b.dd, list.map(function (s) { return s.sbd; })), n = list.length || 1;
    return { maBuoi: b.maBuoi, ngay: b.ngay, trangThai: b.trangThai, dem: c, tiLe: Math.round((c.C + c.M) * 1000 / n) / 10, nopBoi: b.nopBoi, nguoiGhi: b.nguoiGhi };
  });
  var hs = list.map(function (s) {
    var c = { C: 0, M: 0, P: 0, K: 0 }, soBuoi = 0;
    rows.forEach(function (b) { if (b.dd[s.sbd]) { c[b.dd[s.sbd]]++; soBuoi++; } });
    var tiLe = soBuoi ? Math.round((c.C + c.M) * 1000 / soBuoi) / 10 : null;
    return { sbd: s.sbd, ten: (s.ho + ' ' + s.ten).trim(), soBuoi: soBuoi, dem: c, tiLe: tiLe, duoiNguong: tiLe !== null && tiLe < nguong,
      gan: rows.slice(-12).map(function (b) { return b.dd[s.sbd] || ''; }) };
  });
  hs.sort(function (a, b) { return (a.tiLe === null ? 101 : a.tiLe) - (b.tiLe === null ? 101 : b.tiLe); });

  // Cờ cảnh báo
  var co = [];
  // (1) đối chiếu với bài kiểm tra cùng ngày của nhóm
  var baiTheoNgay = {};
  sysRows('BaiKiemTra').forEach(function (bk) {
    if (parseNhomList(bk.cacNhom).indexOf(ma) === -1) return;
    var ng = ngayStr(bk.ngay); if (ng) (baiTheoNgay[ng] = baiTheoNgay[ng] || []).push(bk);
  });
  var tenCua = {}; list.forEach(function (s) { tenCua[s.sbd] = (s.ho + ' ' + s.ten).trim(); });
  rows.forEach(function (b) {
    (baiTheoNgay[b.ngay] || []).forEach(function (bk) {
      var coBai = {}; kqIndex(bk.maBai).forEach(function (r) { if (r.sbd.substring(0, 3) === ma) coBai[r.sbd] = true; });
      var soCo = Object.keys(coBai).length;
      if (!soCo) return;
      list.forEach(function (s) {
        var tt = b.dd[s.sbd] || 'C';
        if ((tt === 'P' || tt === 'K') && coBai[s.sbd]) co.push({ loai: 'vang-co-bai', ngay: b.ngay, maBuoi: b.maBuoi, sbd: s.sbd, ten: tenCua[s.sbd], msg: 'Điểm danh vắng nhưng có bài làm "' + bk.tenBai + '"' });
        if ((tt === 'C' || tt === 'M') && !coBai[s.sbd] && soCo >= list.length * 0.5) co.push({ loai: 'comat-khong-bai', ngay: b.ngay, maBuoi: b.maBuoi, sbd: s.sbd, ten: tenCua[s.sbd], msg: 'Điểm danh có mặt nhưng không có bài làm "' + bk.tenBai + '"' });
      });
    });
  });
  // (2) lớp trưởng sửa sau khi đã gửi
  var setBuoi = {}; rows.forEach(function (b) { setBuoi[b.maBuoi] = b; });
  sysRows('DiemDanhLog').forEach(function (l) {
    if (!setBuoi[String(l.maBuoi)] || String(l.loai) !== 'sửa sau khi gửi') return;
    co.push({ loai: 'sua-sau-gui', ngay: setBuoi[String(l.maBuoi)].ngay, maBuoi: String(l.maBuoi), sbd: normSbd(l.sbd), ten: tenCua[normSbd(l.sbd)] || '',
      msg: String(l.nguoi) + ' sửa ' + (DD_CODES[String(l.cu)] || l.cu) + ' → ' + (DD_CODES[String(l.moi)] || l.moi) + ' sau khi đã gửi điểm danh (' + tStr(l.thoiGian) + ')' });
  });
  co.sort(function (a, b) { return a.ngay < b.ngay ? 1 : -1; });

  var tong = { C: 0, M: 0, P: 0, K: 0 }; buoi.forEach(function (b) { ['C', 'M', 'P', 'K'].forEach(function (k) { tong[k] += b.dem[k]; }); });
  var tongLuot = tong.C + tong.M + tong.P + tong.K;
  return {
    maNhom: ma, tenNhom: st.tenNhom, nguong: nguong, siSo: list.length,
    tongQuan: { soBuoi: rows.length, tiLe: tongLuot ? Math.round((tong.C + tong.M) * 1000 / tongLuot) / 10 : null, dem: tong, duoiNguong: hs.filter(function (h) { return h.duoiNguong; }).length },
    buoi: buoi.slice(-30), ngayGan: rows.slice(-12).map(function (b) { return b.ngay; }), hocSinh: hs, coCanhBao: co.slice(0, 60),
    yeuCau: ddYeuCauList(u, ma).filter(function (y) { return y.trangThai === 'Chờ duyệt'; }),
    lopTruong: ltList(u, ma)
  };
}

/* ---------------- tài khoản lớp trưởng ---------------- */

function ltList(u, maNhom) {
  return sysRows('NguoiDung').filter(function (r) {
    if (r.role !== 'loptruong') return false;
    if (maNhom && normMa(r.maNhom) !== normMa(maNhom)) return false;
    return u.role === 'admin' || String(r.quanLy) === u.username;
  }).map(function (r) { return { username: String(r.username), ten: String(r.fullName), maNhom: normMa(r.maNhom), active: toBool(r.active), mustChange: toBool(r.mustChange) }; });
}

/** Cô tạo tài khoản lớp trưởng từ một học sinh trong nhóm */
function ltTao(p, u) {
  if (!isStaff(u)) throw new Error('Không có quyền');
  var ma = normMa(p.maNhom); ddAccess(u, ma);
  var sbd = normSbd(p.sbd);
  var hs = ddStudents(ma, true).students.filter(function (s) { return s.sbd === sbd; })[0];
  if (!hs) throw new Error('Không tìm thấy học sinh ' + p.sbd + ' trong nhóm đang học');
  var username = 'lt' + sbd;
  if (findUserRow(username)) throw new Error('Em này đã có tài khoản lớp trưởng: ' + username);
  var pw = String(Math.floor(100000 + Math.random() * 900000));
  var sh = _users(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), n = sh.getLastColumn();
  var row = []; for (var i = 0; i < n; i++) row.push('');
  var salt = randomSalt();
  var set = function (k, v) { if (H[k] !== undefined) row[H[k]] = v; };
  set('username', username); set('role', 'loptruong'); set('fullName', (hs.ho + ' ' + hs.ten).trim());
  set('passHash', hashPassword(pw, salt)); set('salt', salt); set('mustChange', true); set('active', true);
  set('createdAt', nowIso()); set('maNhom', ma); set('quanLy', u.role === 'admin' ? (getNhomMap()[ma] || {}).teacher || u.username : u.username);
  sh.appendRow(row);
  var r = sh.getLastRow();
  if (H['maNhom'] !== undefined) setTextCell(sh, r, H['maNhom'] + 1, ma);
  sysTouch();
  return { ok: true, username: username, matKhau: pw };
}

function ltSua(p, u) {
  if (!isStaff(u)) throw new Error('Không có quyền');
  var lt = ltList(u).filter(function (x) { return x.username === String(p.username); })[0];
  if (!lt) throw new Error('Không tìm thấy tài khoản lớp trưởng');
  if (p.xoa) { var r = findUserRow(lt.username); deleteRowsList(r.sheet, [r.row]); revokeSessions(lt.username); return { ok: true }; }
  if (p.datLai) { var pw = String(Math.floor(100000 + Math.random() * 900000)); var ur = findUserRow(lt.username); setPassword(ur, pw, true); revokeSessions(lt.username); return { ok: true, matKhau: pw }; }
  return setUserActive({ username: lt.username, active: !!p.active }, u);
}

/* ---------------- công khai: chuyên cần của một em ---------------- */

function pubChuyenCan(sbd) {
  var ma = sbd.substring(0, 3);
  var key = 'pubdd_' + pubVer() + '_' + ma, idxDd = cacheGetBig(key);
  if (!idxDd) {
    idxDd = ddRows(ma).sort(function (a, b) { return a.ngay < b.ngay ? -1 : 1; }).map(function (b) { return { m: b.maBuoi, n: b.ngay, d: b.dd }; });
    cachePutBig(key, idxDd, 3600);
  }
  var c = { C: 0, M: 0, P: 0, K: 0 }, gan = [];
  idxDd.forEach(function (b) { var t = b.d[sbd]; if (!t) return; c[t]++; gan.push({ maBuoi: b.m, ngay: b.n, tt: t }); });
  var so = c.C + c.M + c.P + c.K;
  return { soBuoi: so, dem: c, tiLe: so ? Math.round((c.C + c.M) * 1000 / so) / 10 : null, gan: gan.slice(-15).reverse() };
}


/* ===================== Anh.gs ===================== */
/** ẢNH PHIẾU — nhận từ web, lưu vào thư mục anh/ của bài, cập nhật trạng thái, báo email */

/** Chạy tay MỘT LẦN sau khi thêm quyền gửi mail, để Google hỏi cấp quyền */
function authorize() {
  MailApp.getRemainingDailyQuota();
  DriveApp.getRootFolder();
  Logger.log('Đã cấp quyền xong. Giờ có thể Deploy phiên bản mới.');
}

function getSetting(key, def) {
  var rows = sysRows('CaiDat');
  for (var i = 0; i < rows.length; i++) if (String(rows[i].key) === key) return rows[i].value;
  return def;
}

function countFilesIn(folder) {
  var n = 0, it = folder.getFiles();
  while (it.hasNext()) { it.next(); n++; }
  return n;
}

/**
 * Thông tin gọn của bài (giáo viên, thư mục bài, thư mục ảnh), lưu cache 6 giờ.
 * Nhờ vậy mỗi lần tải một ảnh không phải đọc lại cả bảng và dò thư mục.
 */
function baiMeta(maBai, fresh) {
  var cache = CacheService.getScriptCache(), key = 'baimeta_' + maBai;
  if (!fresh) {
    var c = cache.get(key);
    if (c) return JSON.parse(c);
  }
  var bai = findBai(maBai);
  if (!bai) throw new Error('Không tìm thấy bài');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var anh = getAnhFolder(bai.driveFolderId);
    var meta = { teacher: bai.teacher, driveFolderId: bai.driveFolderId, anhFolderId: anh.getId() };
    cache.put(key, JSON.stringify(meta), 21600);
    return meta;
  } catch (e) {
    throw new Error('Không mở được thư mục của bài trên Drive (có thể đã bị xóa hoặc đổi quyền): ' + (e.message || e));
  } finally {
    lock.releaseLock();
  }
}

function assertMetaAccess(meta, u) {
  if (u.role !== 'admin' && meta.teacher !== u.username) throw new Error('Không có quyền với bài này');
}

/** Số ảnh hiện có + link thư mục ảnh */
function anhInfo(p, u) {
  var meta = baiMeta(p.maBai);
  assertMetaAccess(meta, u);
  var folder;
  try {
    folder = DriveApp.getFolderById(meta.anhFolderId);
    if (folder.isTrashed()) throw new Error('trashed');
  } catch (e) {
    meta = baiMeta(p.maBai, true);           // thư mục ảnh đã bị xóa/dọn: tạo lại
    folder = DriveApp.getFolderById(meta.anhFolderId);
  }
  return { count: countFilesIn(folder), folderUrl: 'https://drive.google.com/drive/folders/' + meta.anhFolderId };
}

/** Nhận ảnh base64 và tạo file trong anh/. Dùng p.images = [{name, data}] hoặc p.name + p.data */
function uploadAnh(p, u) {
  var meta = baiMeta(p.maBai);
  assertMetaAccess(meta, u);
  var imgs = p.images || [{ name: p.name, data: p.data }];
  if (!imgs.length) throw new Error('Không có ảnh để tải');
  var folder;
  try {
    folder = DriveApp.getFolderById(meta.anhFolderId);
  } catch (e) {
    CacheService.getScriptCache().remove('baimeta_' + p.maBai);
    throw new Error('Thư mục ảnh không còn. Hãy thử lại, hệ thống sẽ tạo lại.');
  }
  var added = 0;
  imgs.forEach(function (im) {
    var data = String(im.data || '');
    if (!data) throw new Error('Ảnh rỗng');
    if (data.length > 12 * 1024 * 1024) throw new Error('Ảnh quá lớn (trên 9 MB)');
    var name = String(im.name || ('anh-' + Date.now())).replace(/[^A-Za-z0-9._-]/g, '_');
    if (!/\.jpe?g$/i.test(name)) name += '.jpg';
    folder.createFile(Utilities.newBlob(Utilities.base64Decode(data), 'image/jpeg', name));
    added++;
  });
  return { ok: true, added: added };
}

/** Xóa (cho vào thùng rác Drive) mọi ảnh của bài để tải lại từ đầu */
function clearAnh(p, u) {
  var meta = baiMeta(p.maBai);
  assertMetaAccess(meta, u);
  var it = DriveApp.getFolderById(meta.anhFolderId).getFiles(), n = 0;
  while (it.hasNext()) { it.next().setTrashed(true); n++; }
  _props().deleteProperty('ANHMAIL_' + p.maBai);
  return { ok: true, removed: n };
}

/** Gọi sau khi tải xong: báo email nếu có ảnh mới (không đổi trạng thái bài) */
function finishAnh(p, u) {
  var meta = baiMeta(p.maBai);
  assertMetaAccess(meta, u);
  return thongBaoAnh(p.maBai);
}

/**
 * Nếu số ảnh tăng so với lần báo trước thì gửi email, kèm ghi chú đã có file đáp án trên hệ thống chưa.
 * File đáp án và ảnh chỉ là tiện ích, không phải điều kiện: hàm này không đổi trạng thái bài
 * và không bao giờ ném lỗi ra ngoài.
 */
function thongBaoAnh(maBai) {
  var out = { soAnh: 0, hasDapAn: false, emailSent: false };
  try {
    var bai = findBai(maBai);
    if (!bai) return out;
    try {
      var folder = DriveApp.getFolderById(bai.driveFolderId);
      out.hasDapAn = folder.getFilesByName('dap-an - ' + safeName(bai.tenBai) + '.xlsx').hasNext();
      out.soAnh = countFilesIn(getAnhFolder(bai.driveFolderId));
    } catch (e) { return out; }
    var props = _props(), key = 'ANHMAIL_' + maBai;
    var last = Number(props.getProperty(key)) || 0;
    if (out.soAnh > last) {
      out.emailSent = notifyAnh(bai, out.soAnh, out.hasDapAn);
      if (out.emailSent) props.setProperty(key, String(out.soAnh));
    }
  } catch (e) { /* bỏ qua */ }
  return out;
}

function notifyAnh(bai, soAnh, hasDapAn) {
  try {
    var to = String(getSetting('emailNhanThongBao', '') || '').trim();
    if (!to) return false;
    var cfg = {};
    try { cfg = JSON.parse(bai.cauHinh || '{}'); } catch (e) {}
    var body =
      'Có ảnh phiếu bài làm mới, có thể chấm bằng TNMaker.\n\n' +
      'Bài: ' + bai.tenBai + '\n' +
      'Ngày: ' + (ngayStr(bai.ngay) || '(chưa ghi)') + '\n' +
      'Giáo viên: ' + bai.teacher + '\n' +
      'Nhóm: ' + (parseNhomList(bai.cacNhom).join(', ') || '(chưa chọn)') + '\n' +
      'Mã đề: ' + (cfg.maDeList || []).join(', ') + '\n' +
      'Số ảnh phiếu: ' + soAnh + '\n' +
      'File đáp án trên hệ thống: ' + (hasDapAn ? 'Có' : 'Chưa có') + '\n\n' +
      'Thư mục của bài trên Drive: https://drive.google.com/drive/folders/' + bai.driveFolderId + '\n';
    MailApp.sendEmail({ to: to, subject: '[Điểm] Có ảnh phiếu mới: ' + bai.tenBai, body: body, name: 'Hệ thống điểm' });
    return true;
  } catch (e) {
    return false;
  }
}

/* ---- Dọn ảnh cũ (chạy tự động hằng ngày) ---- */

function cleanupOldImages() {
  var days = Number(getSetting('soNgayGiuAnh', 30)) || 30;
  var now = Date.now();
  sysRows('BaiKiemTra').forEach(function (r) {
    if (!r.ngayCongBo) return;
    var t = new Date(r.ngayCongBo).getTime();
    if (isNaN(t) || now - t <= days * 86400000) return;
    try {
      var it = DriveApp.getFolderById(r.driveFolderId).getFoldersByName('anh');
      while (it.hasNext()) it.next().setTrashed(true);
      CacheService.getScriptCache().remove('baimeta_' + r.maBai);
    } catch (e) {}
  });
}

/** Cài lịch dọn dẹp mỗi ngày lúc 2 giờ sáng: ảnh cũ, phiên hết hạn, nhật ký lỗi (bấm được trong Cài đặt) */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'cleanupOldImages' || f === 'donDepHangNgay') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('donDepHangNgay').timeBased().everyDays(1).atHour(2).create();
  Logger.log('Đã cài lịch dọn dẹp hằng ngày (2h sáng).');
}


/* ===================== CaiDat.gs ===================== */
/** CÀI ĐẶT HỆ THỐNG, GÓP Ý, DỌN DẸP, NHẬT KÝ LỖI, NÂNG CẤP DỮ LIỆU, KHỞI ĐỘNG */

var CAI_DAT_MAC_DINH = {
  tenHienThi: 'Vật Lý Cô Lâm Cường',
  loiNhanTraDiem: 'Nhập số báo danh để xem điểm, từng câu đúng sai và chuyên cần của con. Có thắc mắc, vui lòng liên hệ trực tiếp cô.',
  hienThiThuHang: 'false',
  hienDiemTB: 'false',
  hienDiemCaoNhat: 'true',
  hienChuyenCan: 'false',
  nguongChuyenCan: '80',
  choPhepGopY: 'true',
  nguongBoTuc: '5',
  soNgayGiuAnh: '30',
  emailNhanThongBao: ''
};
var CAI_DAT_BOOL = { hienThiThuHang: 1, hienDiemTB: 1, choPhepGopY: 1, hienDiemCaoNhat: 1, hienChuyenCan: 1 };
var CAI_DAT_SO = { nguongBoTuc: [0, 10], soNgayGiuAnh: [1, 365], nguongChuyenCan: [0, 100] };

/**
 * Tự nâng cấp dữ liệu: tạo tab còn thiếu, thêm cột mới, thêm cài đặt mới.
 * Chạy ở lượt gọi đầu tiên sau khi cập nhật (rồi nhớ 6 giờ) nên không cần chạy lại setup().
 */
function ensureSchema() {
  var c = CacheService.getScriptCache();
  if (c.get('schema_' + APP_VERSION)) return;
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return; }
  try {
    if (c.get('schema_' + APP_VERSION)) return;
    var ss = getSystemSheet();
    Object.keys(SYS_TABS).forEach(function (name) {
      var sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);
      if (sh.getLastRow() === 0) sh.appendRow(SYS_TABS[name]);
    });
    ensureSysColumns(ss);
    var have = {};
    sysTouch('CaiDat');
    sysRows('CaiDat').forEach(function (r) { have[String(r.key)] = true; });
    var add = [];
    Object.keys(CAI_DAT_MAC_DINH).forEach(function (k) {
      if (have[k]) return;
      var v = CAI_DAT_MAC_DINH[k];
      if (k === 'emailNhanThongBao') { try { v = Session.getEffectiveUser().getEmail() || ''; } catch (e) {} }
      add.push([k, v]);
    });
    var CU = { tenHienThi: ['Lớp Vật lý mẹ Cường', 'Tra điểm lớp học thêm', 'Tra điểm', 'Vật lý cô Cường'],
      loiNhanTraDiem: ['Điểm các bài kiểm tra trên lớp học thêm. Thắc mắc về điểm, hãy hỏi trực tiếp giáo viên.'] };
    var cdSh = ss.getSheetByName('CaiDat');
    sysRows('CaiDat').forEach(function (r) {
      var k = String(r.key);
      if (CU[k] && CU[k].indexOf(String(r.value)) !== -1) setTextCell(cdSh, r._row, 2, CAI_DAT_MAC_DINH[k]);
    });
    var meRow = findUserRow('me');
    if (meRow && String(userField(meRow, 'fullName')) === 'Mẹ (giáo viên Vật lý)') meRow.sheet.getRange(meRow.row, meRow.H['fullName'] + 1).setValue('Cô Lâm Cường');
    if (add.length) {
      var cd = ss.getSheetByName('CaiDat'), r = cd.getLastRow() + 1;
      cd.getRange(r, 1, add.length, 2).setNumberFormat('@');
      cd.getRange(r, 1, add.length, 2).setValues(add);
    }
    sysTouch();
    c.put('schema_' + APP_VERSION, '1', 21600);
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Cài đặt ---------------- */

function getCaiDatAll() {
  var out = {};
  Object.keys(CAI_DAT_MAC_DINH).forEach(function (k) {
    var v = getSetting(k, CAI_DAT_MAC_DINH[k]);
    out[k] = CAI_DAT_BOOL[k] ? toBool(v) : (CAI_DAT_SO[k] ? Number(v) : String(v === null || v === undefined ? '' : v));
  });
  return out;
}

function setCaiDat(p) {
  var vals = p.caiDat || {};
  if (typeof vals === 'string') vals = JSON.parse(vals);
  var sh = getSystemSheet().getSheetByName('CaiDat');
  var rows = sysRows('CaiDat'), rowOf = {};
  rows.forEach(function (r) { rowOf[String(r.key)] = r._row; });
  Object.keys(vals).forEach(function (k) {
    if (!CAI_DAT_MAC_DINH.hasOwnProperty(k)) return;
    var v = vals[k];
    if (CAI_DAT_BOOL[k]) v = toBool(v) ? 'true' : 'false';
    else if (CAI_DAT_SO[k]) {
      var n = Number(v);
      if (!isFinite(n) || n < CAI_DAT_SO[k][0] || n > CAI_DAT_SO[k][1]) throw new Error('Giá trị "' + k + '" phải từ ' + CAI_DAT_SO[k][0] + ' đến ' + CAI_DAT_SO[k][1]);
      v = String(n);
    } else {
      v = String(v || '').trim().slice(0, 500);
      if (k === 'emailNhanThongBao' && v && !/^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(v)) throw new Error('Email nhận thông báo không hợp lệ');
    }
    if (rowOf[k]) setTextCell(sh, rowOf[k], 2, v);
    else { var r = sh.getLastRow() + 1; sh.getRange(r, 1, 1, 2).setNumberFormat('@'); sh.getRange(r, 1, 1, 2).setValues([[k, v]]); rowOf[k] = r; }
  });
  sysTouch(); pubTouch();
  return getCaiDatAll();
}

/* ---------------- Góp ý (quản trị) ---------------- */

function listGopY() {
  var rows = sysRows('GopY').map(function (r) {
    return {
      id: String(r.id || ('row' + r._row)), _row: r._row, thoiGian: tStr(r.thoiGian), noiDung: String(r.noiDung || ''),
      lienHe: String(r.lienHe || ''), daXuLy: toBool(r.daXuLy), nguon: String(r.nguon || ''), nguoiGui: String(r.nguoiGui || ''), sbd: normSbd(r.sbd) || ''
    };
  });
  rows.reverse();
  return { gopY: rows, moi: rows.filter(function (r) { return !r.daXuLy; }).length };
}

function xuLyGopY(p) {
  var sh = getSystemSheet().getSheetByName('GopY');
  var H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  var r = sysRows('GopY').filter(function (x) { return String(x.id || ('row' + x._row)) === String(p.id); })[0];
  if (!r) throw new Error('Không tìm thấy góp ý (có thể đã bị xóa)');
  if (p.xoa) { deleteRowsList(sh, [r._row]); return { ok: true }; }
  sh.getRange(r._row, H['daXuLy'] + 1).setValue(p.daXuLy ? 'TRUE' : '');
  sysTouch('GopY');
  return { ok: true };
}

/* ---------------- Dọn dẹp tự động hằng ngày ---------------- */

function donDepHangNgay() {
  try { cleanupOldImages(); } catch (e) {}
  try { donPhien(); } catch (e) {}
  try { catNhatKyLoi(); } catch (e) {}
}

/** Xóa phiên đăng nhập đã hết hạn */
function donPhien() {
  var sh = getSystemSheet().getSheetByName('Phien');
  var rows = sysRows('Phien'), now = Date.now(), del = [];
  rows.forEach(function (r) { var t = new Date(r.expiry).getTime(); if (isNaN(t) || t < now) del.push(r._row); });
  if (del.length) deleteRowsList(sh, del);
}

function trangThaiTuDong() {
  var on = ScriptApp.getProjectTriggers().some(function (t) {
    var f = t.getHandlerFunction();
    return f === 'donDepHangNgay' || f === 'cleanupOldImages';
  });
  return { daBat: on };
}
function batTuDong() { installTriggers(); return trangThaiTuDong(); }

/* ---------------- Nhật ký lỗi (lỗi lập trình, để sửa) ---------------- */

function ghiLoi(action, p, err) {
  try {
    var sh = getSystemSheet().getSheetByName('NhatKyLoi');
    if (!sh) return;
    var r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, 4).setNumberFormat('@');
    sh.getRange(r, 1, 1, 4).setValues([[nowIso(), String(action), String((p && p.username) || ''),
      String((err && err.stack) || err).slice(0, 1500)]]);
  } catch (e) {}
}
function catNhatKyLoi() {
  var sh = getSystemSheet().getSheetByName('NhatKyLoi');
  if (!sh) return;
  var n = sh.getLastRow() - 1;
  if (n > 500) sh.deleteRows(2, n - 300);
}
function nhatKyLoi() {
  var rows = sysRows('NhatKyLoi').slice(-100).reverse();
  return { loi: rows.map(function (r) { return { thoiGian: tStr(r.thoiGian), action: String(r.action), user: String(r.user), message: String(r.message) }; }) };
}

/* ---------------- Khởi động: một lượt gọi lấy hết dữ liệu cần để mở trang quản lý ---------------- */

function bootstrap(p, u) {
  if (u.role === 'loptruong') {
    var lt = ltInfo(u), reg = getNhomMap()[lt.maNhom] || {};
    return { user: u, version: APP_VERSION, laLopTruong: true, maNhom: lt.maNhom, tenNhom: String(reg.tenNhom || lt.maNhom), caiDat: publicConfig(), groups: [], bai: [], gopYMoi: 0, homNay: homNay() };
  }
  var out = { user: u, version: APP_VERSION, groups: [], bai: [], caiDat: publicConfig(), gopYMoi: 0 };
  try { out.groups = listGroups(u).groups; } catch (e) { out.loiNhom = String(e.message || e); }
  out.bai = listBai(u, {}).bai;
  if (u.role === 'admin') {
    out.caiDat = getCaiDatAll();
    try { out.gopYMoi = listGopY().moi; } catch (e) {}
  }
  try { out.ddChoDuyet = ddYeuCauList(u).filter(function (y) { return y.trangThai === 'Chờ duyệt'; }).length; } catch (e) { out.ddChoDuyet = 0; }
  out.homNay = homNay();
  return out;
}


/* ===================== Bridge.gs ===================== */
/**
 * CẦU NỐI cho máy chủ mới (Cloudflare Worker). Từ bản 7, nghiệp vụ và dữ liệu chạy trên Worker + D1;
 * Apps Script chỉ còn làm những việc cần tài khoản Google: Drive (ảnh phiếu, file đáp án), email,
 * đọc/ghi file Google Sheet danh sách lớp, và nhận bản sao lưu hằng đêm về file Sheet hệ thống cũ.
 *
 * Xác thực: Worker gửi kèm mã dùng một lần, Apps Script gọi ngược về Worker để xác nhận (xem bridgeAuthOk).
 * Địa chỉ Worker đặt ở hằng WORKER_URL trong Config.gs. Không dùng phiên người dùng.
 */
function bridgeHandle(p) {
  if (!bridgeAuthOk(p)) return { ok: false, error: 'Khóa cầu nối không đúng' };
  try {
    var fn = BRIDGE_OPS[String(p.op || '')];
    if (!fn) throw new Error('Không có thao tác cầu nối: ' + p.op);
    return { ok: true, data: fn(p) };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/**
 * Hai cách xác thực:
 *  1) Mã dùng một lần "nonce:<mã>": gọi ngược về Worker (WORKER_URL?bridgeVerify=<mã>) để xác nhận mã do Worker phát ra. Không cần dán khóa ở đâu cả.
 *  2) Khóa cố định BRIDGE_SECRET trong Script Properties (nếu có đặt).
 */
function bridgeAuthOk(p) {
  var tok = String(p.bridge || '');
  if (!tok) return false;
  var secret = String(_props().getProperty('BRIDGE_SECRET') || '');
  if (secret && tok === secret) return true;
  if (tok.indexOf('nonce:') !== 0) return false;
  var base = String(_props().getProperty('WORKER_URL') || (typeof WORKER_URL !== 'undefined' ? WORKER_URL : '') || '').replace(/\/+$/, '');
  if (!/^https:\/\//.test(base)) return false;
  try {
    var res = UrlFetchApp.fetch(base + '/?bridgeVerify=' + encodeURIComponent(tok.slice(6)), { muteHttpExceptions: true, followRedirects: true });
    var j = JSON.parse(res.getContentText());
    return !!(j && j.ok && j.valid);
  } catch (e) { return false; }
}

var BRIDGE_OPS = {
  ping: function () { return { ok: true, luc: nowIso() }; },

  /* ---- đọc toàn bộ file Sheet hệ thống (để chuyển dữ liệu sang máy chủ mới) ---- */
  sysRead: function () {
    var ss = getSystemSheet();
    return { sheetId: ss.getId(), tabs: ss.getSheets().map(function (sh) { var lr = sh.getLastRow(), lc = sh.getLastColumn(); return { name: sh.getName(), values: lr && lc ? sh.getRange(1, 1, lr, lc).getDisplayValues() : [] }; }) };
  },

  /* ---- danh sách lớp: đọc cả file, ghi đè từng tab ---- */
  rosterRead: function (p) {
    var ss = SpreadsheetApp.openById(String(p.fileId));
    return { tabs: ss.getSheets().map(function (sh) { var lr = sh.getLastRow(), lc = sh.getLastColumn(); return { name: sh.getName(), values: lr && lc ? sh.getRange(1, 1, lr, lc).getDisplayValues() : [] }; }) };
  },
  rosterWrite: function (p) {
    var ss = SpreadsheetApp.openById(String(p.fileId)), tabs = p.tabs || [];
    if (typeof tabs === 'string') tabs = JSON.parse(tabs);
    tabs.forEach(function (t) {
      var sh = ss.getSheetByName(t.name) || ss.insertSheet(t.name);
      var vals = (t.values || []).map(function (r) { return r.map(function (v) { return v === null || v === undefined ? '' : v; }); });
      var nc = vals.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
      vals = vals.map(function (r) { while (r.length < nc) r.push(''); return r; });
      var lr = sh.getLastRow(), lc = sh.getLastColumn();
      if (lr && lc) sh.getRange(1, 1, lr, lc).clearContent();
      if (vals.length) { sh.getRange(1, 1, vals.length, nc).setNumberFormat('@'); sh.getRange(1, 1, vals.length, nc).setValues(vals); }
    });
    return { ok: true, tabs: tabs.length };
  },

  /* ---- Drive ---- */
  ensureBaiFolder: function (p) {
    if (p.folderId) { try { var f = DriveApp.getFolderById(String(p.folderId)); if (!f.isTrashed()) return { folderId: f.getId() }; } catch (e) {} }
    var folder = createBaiFolder(String(p.teacher || 'me'), String(p.tenBai || 'Bài'), String(p.ngay || ''));
    return { folderId: folder.getId() };
  },
  anhInfo: function (p) {
    var anh = getAnhFolder(String(p.folderId));
    return { count: countFilesIn(anh), folderUrl: 'https://drive.google.com/drive/folders/' + anh.getId() };
  },
  uploadAnh: function (p) {
    var anh = getAnhFolder(String(p.folderId)), imgs = p.images || [];
    if (typeof imgs === 'string') imgs = JSON.parse(imgs);
    var added = 0;
    imgs.forEach(function (im) {
      var data = String(im.data || ''); if (!data) throw new Error('Ảnh rỗng');
      if (data.length > 12 * 1024 * 1024) throw new Error('Ảnh quá lớn (trên 9 MB)');
      var name = String(im.name || ('anh-' + Date.now())).replace(/[^A-Za-z0-9._-]/g, '_');
      if (!/\.jpe?g$/i.test(name)) name += '.jpg';
      anh.createFile(Utilities.newBlob(Utilities.base64Decode(data), 'image/jpeg', name)); added++;
    });
    return { added: added, count: countFilesIn(anh) };
  },
  clearAnh: function (p) {
    var it = getAnhFolder(String(p.folderId)).getFiles(), n = 0;
    while (it.hasNext()) { it.next().setTrashed(true); n++; }
    return { removed: n };
  },
  saveFile: function (p) {
    var folder = DriveApp.getFolderById(String(p.folderId));
    var file = saveBlobToFolder(folder, Utilities.newBlob(Utilities.base64Decode(String(p.data)), String(p.mime || 'application/octet-stream'), String(p.name)), String(p.name));
    return { fileId: file.getId() };
  },
  genDapAnFile: function (p) {
    var header = p.header, rows = p.rows;
    if (typeof header === 'string') header = JSON.parse(header);
    if (typeof rows === 'string') rows = JSON.parse(rows);
    var tmp = SpreadsheetApp.create('tmp-dapan-' + Date.now());
    try {
      var s2 = tmp.insertSheet('DapAn'), nCol = header.length, nRow = rows.length + 1;
      s2.getRange(1, 1, 1, nCol).setNumberFormat('@');
      if (rows.length) s2.getRange(2, 2, rows.length, nCol - 1).setNumberFormat('@');
      s2.getRange(1, 1, nRow, nCol).setValues([header].concat(rows));
      SpreadsheetApp.flush();
      var blob = exportSpreadsheetAsXlsxBlob(tmp.getId(), 'dap-an.xlsx');
      var folder = DriveApp.getFolderById(String(p.folderId));
      var file = saveBlobToFolder(folder, blob, String(p.name));
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return { fileId: file.getId() };
    } finally { try { DriveApp.getFileById(tmp.getId()).setTrashed(true); } catch (e) {} }
  },
  trash: function (p) { try { DriveApp.getFolderById(String(p.id)).setTrashed(true); } catch (e) { try { DriveApp.getFileById(String(p.id)).setTrashed(true); } catch (e2) {} } return { ok: true }; },
  cleanupOldImages: function (p) {
    var folders = p.folders || []; if (typeof folders === 'string') folders = JSON.parse(folders);
    var n = 0;
    folders.forEach(function (id) { try { var it = DriveApp.getFolderById(String(id)).getFoldersByName('anh'); while (it.hasNext()) { it.next().setTrashed(true); n++; } } catch (e) {} });
    return { removed: n };
  },

  /* ---- email ---- */
  mail: function (p) {
    MailApp.sendEmail({ to: String(p.to), subject: String(p.subject || ''), body: String(p.body || ''), name: String(p.name || 'Hệ thống điểm') });
    return { ok: true };
  },

  /* ---- sao lưu từ D1 về file Sheet hệ thống cũ (ghi đè từng tab) ---- */
  backup: function (p) {
    var ss = SpreadsheetApp.openById(String(p.sheetId)), tabs = p.tabs || [];
    if (typeof tabs === 'string') tabs = JSON.parse(tabs);
    tabs.forEach(function (t) {
      var sh = ss.getSheetByName(t.name) || ss.insertSheet(t.name);
      var vals = (t.values || []).map(function (r) { return r.map(function (v) { return v === null || v === undefined ? '' : v; }); });
      var nc = vals.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
      vals = vals.map(function (r) { while (r.length < nc) r.push(''); return r; });
      var lr = sh.getLastRow(), lc = sh.getLastColumn();
      if (lr && lc) sh.getRange(1, 1, lr, lc).clearContent();
      if (vals.length) { sh.getRange(1, 1, vals.length, nc).setNumberFormat('@'); sh.getRange(1, 1, vals.length, nc).setValues(vals); }
    });
    var meta = ss.getSheetByName('_SaoLuu') || ss.insertSheet('_SaoLuu');
    meta.getRange(1, 1, 1, 2).setValues([['Sao lưu tự động từ máy chủ mới lúc', nowIso()]]);
    return { ok: true, tabs: tabs.length };
  }
};

function bridgeCell(v) {
  if (v instanceof Date) return v.toISOString();
  return v === null || v === undefined ? '' : v;
}

/** Chạy tay MỘT LẦN trong trình soạn Apps Script: sinh khóa bí mật cho cầu nối, in ra để dán vào Worker */
function taoKhoaCauNoi() {
  var p = _props(), k = p.getProperty('BRIDGE_SECRET');
  if (!k) { k = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); p.setProperty('BRIDGE_SECRET', k); }
  Logger.log('BRIDGE_SECRET = ' + k);
  return k;
}


/* ===================== Code.gs ===================== */
/** ĐIỂM VÀO WEB APP — mọi lượt gọi từ trang web đi qua đây */

function doGet(e) { return handle(e, 'GET'); }
function doPost(e) { return handle(e, 'POST'); }

/*
 * Kết quả trả về: { ok, data, ms, v } hoặc { ok:false, error, code, ms }.
 * code: AUTH (cần đăng nhập lại), FORBIDDEN (không đủ quyền), ERR (lỗi nghiệp vụ, có lời nhắn rõ).
 * Lỗi lập trình (TypeError, …) được ghi vào tab NhatKyLoi để sửa.
 */
function handle(e, method) {
  var t0 = Date.now(), p = {}, action = '';
  try {
    _ROWS_MEMO = {}; _SS_MEMO = {}; _DVER = null; _KHONG_DOI = false;     // mỗi lượt gọi bắt đầu sạch
    if (method === 'POST' && e && e.postData && e.postData.contents) p = JSON.parse(e.postData.contents);
    else if (e) p = e.parameter || {};
    if (p.bridge !== undefined) return jsonOut(bridgeHandle(p));          // cầu nối từ Worker (xác thực bằng khóa bí mật)
    action = String(p.action || '');
    if (action !== 'ping') ensureSchema();
    var data = route(action, p);
    if (!SERVER_READ[action] && !_KHONG_DOI) { try { dataTouch(); } catch (e) {} }        // thao tác ghi: đổi phiên bản dữ liệu
    return jsonOut({ ok: true, data: data, ms: Date.now() - t0, v: APP_VERSION });
  } catch (err) {
    var msg = String((err && err.message) || err);
    var code = (err && err.code) || (/Phiên đã hết hạn/.test(msg) ? 'AUTH' : (/Không đủ quyền|Không có quyền/.test(msg) ? 'FORBIDDEN' : 'ERR'));
    if (err && err.name && err.name !== 'Error') {
      ghiLoi(action, p, err);
      msg = 'Lỗi hệ thống khi xử lý (' + action + '). Đã ghi lại để sửa. Chi tiết: ' + msg;
    }
    return jsonOut({ ok: false, error: msg, code: code, ms: Date.now() - t0, v: APP_VERSION });
  }
}

/** Lệnh CHỈ ĐỌC (không đổi dữ liệu). Lệnh không có trong đây đều được coi là ghi: an toàn khi quên khai báo. */
var SERVER_READ = { ping: 1, me: 1, bootstrap: 1, listTeachers: 1, trangThaiTuDong: 1, nhatKyLoi: 1, validateRoster: 1, getBai: 1, publicConfig: 1, traDiem: 1, chiTietBai: 1, listBai: 1, getBaiFull: 1, getKetQua: 1,
  thongKeBai: 1, listGroups: 1, listStudents: 1, listUsers: 1, getCaiDat: 1, listGopY: 1, getNhatKy: 1, previewDiem: 1,
  ddTongQuan: 1, ddPhanTich: 1, ddXemBuoi: 1, ngoaiDoiChieu: 1 };

function route(action, p) {
  switch (action) {
    /* ---- công khai, không cần đăng nhập ---- */
    case 'ping':             return { time: nowIso(), v: APP_VERSION };
    case 'login':            return apiLogin(p);
    case 'quenMatKhau':      return quenMatKhau(p);
    case 'datLaiMatKhau':    return datLaiMatKhau(p);
    case 'publicConfig':     return publicConfig();
    case 'traDiem':          return traDiem(p);
    case 'chiTietBai':       return chiTietBai(p);
    case 'guiGopY':          return guiGopY(p, optionalUser(p));

    /* ---- tài khoản ---- */
    case 'logout':           return apiLogout(p);
    case 'me':               return { user: currentUser(p) };
    case 'changePassword':   return apiChangePassword(p);
    case 'bootstrap':        return requireAuth(p, function (u) { return bootstrap(p, u); });

    /* ---- quản trị ---- */
    case 'listUsers':        return requireRole(p, 'admin', function () { return listUsers(); });
    case 'listTeachers':     return requireRole(p, 'admin', function () { return listTeachers(); });
    case 'createUser':       return requireRole(p, 'admin', function () { return createUser(p); });
    case 'setTeacherRoster': return requireRole(p, 'admin', function () { return setTeacherRoster(p); });
    case 'setUserInfo':      return requireRole(p, 'admin', function () { return setUserInfo(p); });
    case 'setUserActive':    return requireRole(p, 'admin', function (u) { return setUserActive(p, u); });
    case 'resetPassword':    return requireRole(p, 'admin', function () { return adminResetPassword(p); });
    case 'getCaiDat':        return requireRole(p, 'admin', function () { return getCaiDatAll(); });
    case 'setCaiDat':        return requireRole(p, 'admin', function () { return setCaiDat(p); });
    case 'listGopY':         return requireRole(p, 'admin', function () { return listGopY(); });
    case 'xuLyGopY':         return requireRole(p, 'admin', function () { return xuLyGopY(p); });
    case 'trangThaiTuDong':  return requireRole(p, 'admin', function () { return trangThaiTuDong(); });
    case 'batTuDong':        return requireRole(p, 'admin', function () { return batTuDong(); });
    case 'nhatKyLoi':        return requireRole(p, 'admin', function () { return nhatKyLoi(); });

    /* ---- nhóm, học sinh ---- */
    case 'syncGroups':       return requireStaff(p, function (u) { return apiSyncGroups(p, u); });
    case 'listGroups':       return requireStaff(p, function (u) { return listGroups(u, !!p.fresh); });
    case 'listStudents':     return requireStaff(p, function (u) { return listStudents(p, u); });
    case 'upsertStudent':    return requireStaff(p, function (u) { return upsertStudent(p, u); });
    case 'setStudentStatus': return requireStaff(p, function (u) { return setStudentStatus(p, u); });
    case 'transferStudent':  return requireStaff(p, function (u) { return transferStudent(p, u); });
    case 'validateRoster':   return requireStaff(p, function (u) { return validateRoster(u); });

    /* ---- bài kiểm tra, đáp án ---- */
    case 'createBai':        return requireStaff(p, function (u) { return createBai(p, u); });
    case 'updateBai':        return requireStaff(p, function (u) { return updateBai(p, u); });
    case 'listBai':          return requireStaff(p, function (u) { return listBai(u, p); });
    case 'getBai':           return requireStaff(p, function (u) { return getBai(p, u); });
    case 'getBaiFull':       return requireStaff(p, function (u) { return getBaiFull(p, u); });
    case 'saveDapAn':        return requireStaff(p, function (u) { return saveDapAn(p, u); });
    case 'genDapAnFile':     return requireStaff(p, function (u) { return genDapAnFile(p, u); });
    case 'luuTruBai':        return requireStaff(p, function (u) { return luuTruBai(p, u); });
    case 'xoaBai':           return requireStaff(p, function (u) { return xoaBai(p, u); });
    case 'luuFileDrive':     return requireStaff(p, function (u) { return luuFileDrive(p, u); });

    /* ---- ảnh phiếu ---- */
    case 'anhInfo':          return requireStaff(p, function (u) { return anhInfo(p, u); });
    case 'uploadAnh':        return requireStaff(p, function (u) { return uploadAnh(p, u); });
    case 'clearAnh':         return requireStaff(p, function (u) { return clearAnh(p, u); });
    case 'finishAnh':        return requireStaff(p, function (u) { return finishAnh(p, u); });

    /* ---- điểm ---- */
    case 'previewDiem':      return requireStaff(p, function (u) { return previewDiem(p, u); });
    case 'commitDiem':       return requireStaff(p, function (u) { return commitDiem(p, u); });
    case 'getKetQua':        return requireStaff(p, function (u) { return getKetQua(p, u); });
    case 'editKetQua':       return requireStaff(p, function (u) { return editKetQua(p, u); });
    case 'xoaKetQua':        return requireStaff(p, function (u) { return xoaKetQua(p, u); });
    case 'setCongBo':        return requireStaff(p, function (u) { return setCongBo(p, u); });
    case 'getNhatKy':        return requireStaff(p, function (u) { return getNhatKy(p, u); });
    case 'giaiQuyetChoXuLy': return requireStaff(p, function (u) { return giaiQuyetChoXuLy(p, u); });
    case 'thongKeBai':       return requireStaff(p, function (u) { return thongKeBai(p, u); });

    /* ---- điểm danh (lớp trưởng dùng được: quyền kiểm tra bên trong) ---- */
    case 'ddTongQuan':       return requireAuth(p, function (u) { return ddTongQuan(p, u); });
    case 'ddMoBuoi':         return requireAuth(p, function (u) { return ddMoBuoi(p, u); });
    case 'ddXemBuoi':        return requireAuth(p, function (u) { return ddXemBuoi(p, u); });
    case 'ddLuu':            return requireAuth(p, function (u) { return ddLuu(p, u); });
    case 'ddNop':            return requireAuth(p, function (u) { return ddNop(p, u); });
    case 'ddYeuCau':         return requireAuth(p, function (u) { return ddYeuCau(p, u); });
    case 'ddChot':           return requireStaff(p, function (u) { return ddChot(p, u); });
    case 'ddXoaBuoi':        return requireStaff(p, function (u) { return ddXoaBuoi(p, u); });
    case 'ddDuyet':          return requireStaff(p, function (u) { return ddDuyet(p, u); });
    case 'ddPhanTich':       return requireStaff(p, function (u) { return ddPhanTich(p, u); });
    case 'ltTao':            return requireStaff(p, function (u) { return ltTao(p, u); });
    case 'ltSua':            return requireStaff(p, function (u) { return ltSua(p, u); });
    case 'ddBaoSai':         return ddYeuCau(p, null);
    case 'ngoaiDoiChieu':    return requireStaff(p, function (u) { return ngoaiDoiChieu(p, u); });
    case 'ngoaiLuu':         return requireStaff(p, function (u) { return ngoaiLuu(p, u); });
    case 'ngoaiSuaDiem':     return requireStaff(p, function (u) { return ngoaiSuaDiem(p, u); });

    default: throw new Error('Hành động không hợp lệ: ' + action);
  }
}

function requireAuth(p, fn) { return fn(currentUser(p)); }
/** Chức năng của giáo viên / quản trị (tài khoản lớp trưởng không dùng được) */
function requireStaff(p, fn) {
  var u = currentUser(p);
  if (u.role !== 'admin' && u.role !== 'teacher') { var e = new Error('Không đủ quyền'); e.code = 'FORBIDDEN'; throw e; }
  return fn(u);
}
function requireRole(p, role, fn) {
  var u = currentUser(p);
  if (u.role !== role) { var e = new Error('Không đủ quyền'); e.code = 'FORBIDDEN'; throw e; }
  return fn(u);
}


  const _fn = {};
  try { _fn.handle = handle; } catch (e) {}
  try { _fn.route = route; } catch (e) {}
  try { _fn.currentUser = currentUser; } catch (e) {}
  try { _fn.requireAuth = requireAuth; } catch (e) {}
  try { _fn.requireStaff = requireStaff; } catch (e) {}
  try { _fn.findBai = findBai; } catch (e) {}
  try { _fn.requireBaiAccess = requireBaiAccess; } catch (e) {}
  try { _fn.sysRows = sysRows; } catch (e) {}
  try { _fn.sysTouch = sysTouch; } catch (e) {}
  try { _fn.getSetting = getSetting; } catch (e) {}
  try { _fn.getBai = getBai; } catch (e) {}
  try { _fn.parseJsonSafe = parseJsonSafe; } catch (e) {}
  try { _fn.safeName = safeName; } catch (e) {}
  try { _fn.ngayStr = ngayStr; } catch (e) {}
  try { _fn.parseNhomList = parseNhomList; } catch (e) {}
  try { _fn.donDepHangNgay = donDepHangNgay; } catch (e) {}
  try { _fn.cleanupOldImages = cleanupOldImages; } catch (e) {}
  try { _fn.pubTouch = pubTouch; } catch (e) {}
  try { _fn.dataTouch = dataTouch; } catch (e) {}
  try { _fn.ensureSchema = ensureSchema; } catch (e) {}
  try { _fn.setTextCell = setTextCell; } catch (e) {}
  try { _fn.idx = idx; } catch (e) {}
  try { _fn.nowIso = nowIso; } catch (e) {}
  try { _fn.getSystemSheet = getSystemSheet; } catch (e) {}
  try { _fn.getNhomMap = getNhomMap; } catch (e) {}
  try { _fn.rosterTouch = rosterTouch; } catch (e) {}
  try { _fn.listGroups = listGroups; } catch (e) {}
  try { _fn.syncTeacher = syncTeacher; } catch (e) {}
  try { _fn.logSua = logSua; } catch (e) {}
  return _fn;
}
