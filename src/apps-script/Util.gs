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
