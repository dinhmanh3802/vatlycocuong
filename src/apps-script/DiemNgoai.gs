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
