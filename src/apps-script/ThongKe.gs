/** THỐNG KÊ MỘT BÀI — để giáo viên biết chữa câu nào, bổ túc cho em nào */

function _median(arr) {
  if (!arr.length) return null;
  var a = arr.slice().sort(function (x, y) { return x - y; }), m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : round2((a[m - 1] + a[m]) / 2);
}

function _tomTat(list) {
  var n = list.length, out = { n: n, tb: null, trungVi: null, cao: null, thap: null, p1: null, p2: null, p3: null, pho: [] };
  for (var k = 0; k <= 10; k++) out.pho.push(0);
  if (!n) return out;
  var t = list.map(function (r) { return r.tong; }), s = 0, s1 = 0, s2 = 0, s3 = 0;
  list.forEach(function (r) {
    s += r.tong; s1 += r.p1; s2 += r.p2; s3 += r.p3;
    var bin = Math.min(10, Math.max(0, Math.floor(r.tong + 1e-9)));
    out.pho[bin]++;
  });
  out.tb = round2(s / n); out.trungVi = _median(t);
  out.cao = Math.max.apply(null, t); out.thap = Math.min.apply(null, t);
  out.p1 = round2(s1 / n); out.p2 = round2(s2 / n); out.p3 = round2(s3 / n);
  return out;
}

function thongKeBai(p, u) {
  var bai = findBai(p.maBai);
  requireBaiAccess(u, bai);
  if (parseJsonSafe(bai.cauHinh, {}).loai === 'ngoai') throw new Error('Bài chỉ có điểm tổng: xem phổ điểm ở trang Kết quả');
  return cachedResult('tk_' + bai.maBai + '_' + u.username, 900, function () { return thongKeBai_(p, u, bai); }, p.fresh);
}
function thongKeBai_(p, u, bai) {
  var cfg = parseJsonSafe(bai.cauHinh, {});
  var mx = maxDiemPhan(cfg);
  var res = kqLoad(bai.maBai), keys = getBaiDapAn(bai.maBai);
  var groups = loadAccessibleGroups(u, false);
  var nguong = Number(getSetting('nguongBoTuc', 5)) || 5;

  // ---- tổng quan, theo nhóm
  var byGroup = {};
  res.forEach(function (r) { var ma = r.sbd.substring(0, 3); (byGroup[ma] = byGroup[ma] || []).push(r); });
  var maList = parseNhomList(bai.cacNhom);
  Object.keys(byGroup).forEach(function (m) { if (maList.indexOf(m) === -1) maList.push(m); });
  maList.sort();
  var names = {};
  var nhomOut = maList.map(function (ma) {
    var g = groups[ma], list = byGroup[ma] || [];
    var dangHoc = g ? g.students.filter(function (s) { return s.trangThai === 'Đang học'; }) : [];
    if (g) g.students.forEach(function (s) { names[s.sbd] = (s.ho + ' ' + s.ten).trim(); });
    var co = {};
    list.forEach(function (r) { co[r.sbd] = true; });
    var t = _tomTat(list);
    t.maNhom = ma; t.tenNhom = g ? g.tenNhom : ma;
    t.vang = dangHoc.filter(function (s) { return !co[s.sbd]; }).length;
    return t;
  });

  // ---- theo câu, tách theo mã đề (các mã đã trộn nên câu 1 mã này khác câu 1 mã kia)
  var theoMa = {}, coChiTiet = 0;
  res.forEach(function (r) {
    var k = keys[r.maDe];
    if (!k || !r.tl) return;
    var sc = chamBai(cfg, r.tl, k);
    if (sc.thieuKey) return;
    coChiTiet++;
    var M = theoMa[r.maDe];
    if (!M) {
      M = theoMa[r.maDe] = { maDe: r.maDe, n: 0, p1: [], p2: [], p3: [] };
      (k.p1 || []).forEach(function (x, i) { if (cfg.p1 && cfg.p1.on && i < cfg.p1.soCau) M.p1.push({ cau: i + 1, dapAn: x, dung: 0, chon: { A: 0, B: 0, C: 0, D: 0, '': 0 } }); });
      if (cfg.p2 && cfg.p2.on) for (var j = 0; j < cfg.p2.soCau; j++) M.p2.push({ cau: j + 1, dapAn: normP2((k.p2 || [])[j]), y: [0, 0, 0, 0], du4: 0 });
      if (cfg.p3 && cfg.p3.on) for (var q = 0; q < cfg.p3.soCau; q++) M.p3.push({ cau: q + 1, dapAn: normP3((k.p3 || [])[q]), dung: 0, sai: {} });
    }
    M.n++;
    M.p1.forEach(function (c, i) {
      var a = String((r.tl.p1 || [])[i] || '').toUpperCase();
      c.chon[/^[A-D]$/.test(a) ? a : '']++;
      if (sc.dung.p1[i] === 1) c.dung++;
    });
    M.p2.forEach(function (c, i) {
      var a = normP2((r.tl.p2 || [])[i]);
      for (var y = 0; y < 4; y++) if (a.charAt(y) === c.dapAn.charAt(y)) c.y[y]++;
      if (sc.dung.p2[i] === 4) c.du4++;
    });
    M.p3.forEach(function (c, i) {
      if (sc.dung.p3[i] === 1) c.dung++;
      else { var a = normP3((r.tl.p3 || [])[i]) || '(bỏ trống)'; c.sai[a] = (c.sai[a] || 0) + 1; }
    });
  });
  var pct = function (x, n) { return n ? Math.round(x * 1000 / n) / 10 : 0; };
  var khoNhat = [];
  var maOut = Object.keys(theoMa).sort().map(function (md) {
    var M = theoMa[md];
    var o = {
      maDe: md, n: M.n,
      p1: M.p1.map(function (c) {
        var tl = pct(c.dung, M.n);
        khoNhat.push({ maDe: md, phan: 'I', cau: c.cau, tiLe: tl });
        var nhieu = ['A', 'B', 'C', 'D'].filter(function (x) { return x !== c.dapAn; }).sort(function (a, b) { return c.chon[b] - c.chon[a]; })[0];
        return { cau: c.cau, dapAn: c.dapAn, tiLe: tl, chon: c.chon, nhieuNhat: c.chon[nhieu] ? nhieu : null };
      }),
      p2: M.p2.map(function (c) {
        var tl = pct(c.du4, M.n);
        khoNhat.push({ maDe: md, phan: 'II', cau: c.cau, tiLe: tl });
        return { cau: c.cau, dapAn: c.dapAn, tiLeDu4: tl, y: c.y.map(function (x) { return pct(x, M.n); }) };
      }),
      p3: M.p3.map(function (c) {
        var tl = pct(c.dung, M.n);
        khoNhat.push({ maDe: md, phan: 'III', cau: c.cau, tiLe: tl });
        var sai = Object.keys(c.sai).map(function (k) { return { gt: k, n: c.sai[k] }; }).sort(function (a, b) { return b.n - a.n; }).slice(0, 3);
        return { cau: c.cau, dapAn: c.dapAn, tiLe: tl, saiHay: sai };
      })
    };
    return o;
  });
  khoNhat.sort(function (a, b) { return a.tiLe - b.tiLe; });

  // ---- cần bổ túc
  var boTuc = [];
  res.forEach(function (r) {
    var ly = [];
    if (r.tong < nguong) ly.push('dưới ' + nguong + ' điểm');
    ['p1', 'p2', 'p3'].forEach(function (k, i) { if (mx[i] > 0 && r[k] < mx[i] / 2) ly.push('Phần ' + ['I', 'II', 'III'][i] + ' dưới một nửa'); });
    if (ly.length) boTuc.push({ sbd: r.sbd, ten: names[r.sbd] || '', tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, lyDo: ly });
  });
  boTuc.sort(function (a, b) { return a.tong - b.tong; });

  // ---- so với các bài trước của cùng nhóm (tối đa 6 bài gần nhất)
  var cungNhom = sysRows('BaiKiemTra').filter(function (b) {
    if (u.role !== 'admin' && b.teacher !== u.username) return false;
    var ms = parseNhomList(b.cacNhom);
    return ms.some(function (m) { return maList.indexOf(m) !== -1; });
  });
  cungNhom.sort(function (a, b) { var x = ngayStr(a.ngay) || String(a.createdAt), y = ngayStr(b.ngay) || String(b.createdAt); return x < y ? -1 : (x > y ? 1 : 0); });
  var vt = cungNhom.map(function (b) { return String(b.maBai); }).indexOf(String(bai.maBai));
  var chon = cungNhom.slice(Math.max(0, vt - 5), vt + 1);
  var chonSet = {};
  chon.forEach(function (b) { chonSet[String(b.maBai)] = { maBai: String(b.maBai), tenBai: String(b.tenBai), ngay: ngayStr(b.ngay), nhom: {} }; });
  var shk = _kq(), last = shk.getLastRow();
  if (last >= 2 && chon.length > 1) {
    var H = kqHeader(shk);
    var v = shk.getRange(2, 1, last - 1, Math.max(H['maBai'], H['sbd'], H['tong']) + 1).getValues();
    v.forEach(function (row) {
      var o = chonSet[String(row[H['maBai']])];
      if (!o) return;
      var ma = normSbd(row[H['sbd']]).substring(0, 3);
      if (maList.indexOf(ma) === -1) return;
      var g = o.nhom[ma] = o.nhom[ma] || { s: 0, n: 0 };
      g.s += Number(row[H['tong']]); g.n++;
    });
  } else if (chon.length === 1) {
    res.forEach(function (r) { var ma = r.sbd.substring(0, 3); var g = chonSet[String(bai.maBai)].nhom[ma] = chonSet[String(bai.maBai)].nhom[ma] || { s: 0, n: 0 }; g.s += r.tong; g.n++; });
  }
  var xuHuong = chon.map(function (b) {
    var o = chonSet[String(b.maBai)], nh = {};
    Object.keys(o.nhom).forEach(function (m) { nh[m] = o.nhom[m].n ? round2(o.nhom[m].s / o.nhom[m].n) : null; });
    return { maBai: o.maBai, tenBai: o.tenBai, ngay: o.ngay, laBaiNay: o.maBai === String(bai.maBai), tbNhom: nh };
  }).filter(function (x) { return Object.keys(x.tbNhom).length; });

  return {
    bai: { maBai: bai.maBai, tenBai: String(bai.tenBai), ngay: ngayStr(bai.ngay), trangThai: bai.trangThai },
    mx: mx, nguong: nguong,
    tongQuan: _tomTat(res), nhom: nhomOut,
    coChiTiet: coChiTiet, theoMa: maOut, khoNhat: khoNhat.slice(0, 8),
    boTuc: boTuc, xuHuong: xuHuong
  };
}
