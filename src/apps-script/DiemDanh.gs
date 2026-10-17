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
