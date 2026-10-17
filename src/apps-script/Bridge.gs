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
