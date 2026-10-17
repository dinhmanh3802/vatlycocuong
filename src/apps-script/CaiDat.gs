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
