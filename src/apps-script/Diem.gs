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
