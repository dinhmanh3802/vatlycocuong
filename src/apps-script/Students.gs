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
