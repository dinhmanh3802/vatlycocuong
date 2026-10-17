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
