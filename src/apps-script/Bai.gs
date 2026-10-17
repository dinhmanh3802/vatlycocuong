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
