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
