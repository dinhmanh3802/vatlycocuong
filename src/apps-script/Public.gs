/**
 * TRANG TRA ĐIỂM CÔNG KHAI — học sinh/phụ huynh nhập SBD để xem điểm các bài ĐÃ CÔNG BỐ.
 *
 * Tốc độ: mỗi nhóm có một "chỉ mục" gom sẵn điểm đã công bố của nhóm đó (kèm thứ hạng, điểm TB),
 * giữ trong bộ nhớ đệm 1 giờ. Một lượt tra chỉ đọc chỉ mục của đúng nhóm em đó, không đọc cả bảng điểm.
 * Mọi thay đổi ảnh hưởng tới điểm đã công bố gọi pubTouch() để chỉ mục được dựng lại ở lượt tra sau.
 */

function pubVer() { return _props().getProperty('PUB_VER') || '0'; }
function pubTouch() { _props().setProperty('PUB_VER', Date.now().toString(36) + Math.random().toString(36).slice(2, 7)); }

/** Cấu hình hiển thị cho trang công khai */
function publicConfig() {
  var c = CacheService.getScriptCache(), k = 'pubcfg_' + pubVer();
  var hit = c.get(k);
  if (hit) return JSON.parse(hit);
  var out = {
    tenHienThi: String(getSetting('tenHienThi', '') || 'Tra điểm'),
    loiNhan: String(getSetting('loiNhanTraDiem', '') || ''),
    hienThiThuHang: toBool(getSetting('hienThiThuHang', false)),
    hienDiemTB: toBool(getSetting('hienDiemTB', false)),
    choPhepGopY: toBool(getSetting('choPhepGopY', true)),
    hienDiemCaoNhat: toBool(getSetting('hienDiemCaoNhat', true)),
    hienChuyenCan: toBool(getSetting('hienChuyenCan', false))
  };
  c.put(k, JSON.stringify(out), 3600);
  return out;
}

function maxDiemPhan(cfg) {
  return [
    (cfg.p1 && cfg.p1.on) ? Number(cfg.p1.diem) || 0 : 0,
    (cfg.p2 && cfg.p2.on) ? (Number(cfg.p2.soCau) || 0) * (Number(cfg.p2.d4) || 0) : 0,
    (cfg.p3 && cfg.p3.on) ? Number(cfg.p3.diem) || 0 : 0
  ];
}

/** Chỉ mục điểm đã công bố của MỘT nhóm (có bộ nhớ đệm) */
function pubGroupIndex(ma) {
  var key = 'pubg_' + pubVer() + '_' + ma;
  var hit = cacheGetBig(key);
  if (hit) return hit;
  var built = buildPubGroupIndex(ma);
  cachePutBig(key, built, 3600);
  return built;
}

function buildPubGroupIndex(ma) {
  var out = { tenNhom: '', hs: {}, kq: {}, bai: {} };
  var info = getNhomMap()[ma];
  if (!info) return out;
  out.tenNhom = String(info.tenNhom || '');
  try {
    var g = findTab(rosterSnapshot(info.rosterSheetId, false), info.sheetName, ma);
    if (g) {
      out.tenNhom = g.tenNhom;
      g.students.forEach(function (s) { if (s.sbd) out.hs[s.sbd] = { ho: s.ho, ten: s.ten, cu: s.sbdCu || '', tt: s.trangThai }; });
    }
  } catch (e) { /* không đọc được danh sách: vẫn trả điểm, chỉ thiếu tên */ }

  var pub = {};
  sysRows('BaiKiemTra').forEach(function (b) { if (b.trangThai === 'Đã công bố') pub[String(b.maBai)] = b; });
  if (!Object.keys(pub).length) return out;

  var sh = _kq(), last = sh.getLastRow();
  if (last < 2) return out;
  var H = kqHeader(sh);
  var need = Math.max(H['maBai'], H['sbd'], H['maDe'], H['tong'], H['p1'], H['p2'], H['p3']) + 1;
  var v = sh.getRange(2, 1, last - 1, need).getValues(), byBai = {};
  for (var i = 0; i < v.length; i++) {
    var mb = String(v[i][H['maBai']]);
    if (!pub[mb]) continue;
    var sbd = normSbd(v[i][H['sbd']]);
    if (!sbd || sbd.substring(0, 3) !== ma) continue;
    (byBai[mb] = byBai[mb] || []).push({
      sbd: sbd, m: normMaDe(v[i][H['maDe']]), t: Number(v[i][H['tong']]),
      p: [Number(v[i][H['p1']]), Number(v[i][H['p2']]), Number(v[i][H['p3']])]
    });
  }
  Object.keys(byBai).forEach(function (mb) {
    var list = byBai[mb], b = pub[mb], cfg = parseJsonSafe(b.cauHinh, {});
    var sum = 0, cao = 0;
    list.forEach(function (x) { sum += x.t; if (x.t > cao) cao = x.t; });
    out.bai[mb] = {
      ten: String(b.tenBai), ngay: ngayStr(b.ngay), cb: tStr(b.ngayCongBo),
      n: list.length, tb: round2(sum / list.length), cao: round2(cao), mx: maxDiemPhan(cfg),
      loai: cfg.loai === 'ngoai' ? 'ngoai' : '', thang: Number(cfg.thang) || 10, nguon: cfg.nguon || ''
    };
    list.forEach(function (x) {
      var hang = 1;                                   // cùng điểm thì cùng hạng (1, 2, 2, 4…)
      list.forEach(function (y) { if (y.t > x.t + 1e-9) hang++; });
      (out.kq[x.sbd] = out.kq[x.sbd] || []).push({ b: mb, m: x.m, t: x.t, p: x.p, h: hang });
    });
  });
  return out;
}

function pubEntry(e, gi, cfgPub, viaSbd) {
  var b = gi.bai[e.b] || {};
  return {
    maBai: e.b, tenBai: b.ten, ngay: b.ngay, maDe: e.m,
    tong: e.t, p1: e.p[0], p2: e.p[1], p3: e.p[2], mx: b.mx || [0, 0, 0], loai: b.loai || '', thang: b.thang || 10, nguon: b.nguon || '',
    hang: cfgPub.hienThiThuHang ? e.h : null,
    soHs: cfgPub.hienThiThuHang ? b.n : null,
    tbNhom: cfgPub.hienDiemTB ? b.tb : null,
    caoNhat: cfgPub.hienDiemCaoNhat ? (b.cao !== undefined ? b.cao : null) : null,
    sbdCu: viaSbd || null
  };
}

/** Chuỗi SBD của một em: SBD hiện tại + các SBD cũ (em chuyển nhóm) */
function sbdChain(sbd) {
  var chain = [sbd], cur = sbd;
  for (var k = 0; k < 4; k++) {
    var h = pubGroupIndex(cur.substring(0, 3)).hs[cur];
    if (!h || !h.cu || chain.indexOf(h.cu) !== -1) break;
    chain.push(h.cu);
    cur = h.cu;
  }
  return chain;
}

/** Tra điểm theo SBD (không cần đăng nhập) */
function traDiem(p) {
  var raw = String(p.sbd == null ? '' : p.sbd).replace(/\s/g, '');
  if (!/^\d{6}$/.test(raw)) throw new Error('Số báo danh gồm đúng 6 chữ số');
  var sbd = raw;
  var cfgPub = publicConfig();
  var gi = pubGroupIndex(sbd.substring(0, 3));
  var hs = gi.hs[sbd];
  var chain = sbdChain(sbd), list = [];
  chain.forEach(function (s, i) {
    var g = i === 0 ? gi : pubGroupIndex(s.substring(0, 3));
    (g.kq[s] || []).forEach(function (e) { list.push(pubEntry(e, g, cfgPub, i === 0 ? null : s)); });
  });
  if (!hs && !list.length) return { found: false, sbd: sbd };
  list.sort(function (a, b) { return String(b.ngay || '') < String(a.ngay || '') ? -1 : (String(b.ngay || '') > String(a.ngay || '') ? 1 : 0); });
  return {
    found: true, sbd: sbd,
    ho: hs ? hs.ho : '', ten: hs ? hs.ten : '', tenNhom: gi.tenNhom,
    daChuyen: hs ? hs.tt === 'Đã chuyển' : false,
    hienThiThuHang: cfgPub.hienThiThuHang, hienDiemTB: cfgPub.hienDiemTB, hienDiemCaoNhat: cfgPub.hienDiemCaoNhat,
    chuyenCan: cfgPub.hienChuyenCan ? pubChuyenCan(sbd) : null,
    bai: list
  };
}

/** Chi tiết một bài của một em: từng câu chọn gì, đáp án đúng, đúng/sai */
function chiTietBai(p) {
  var raw = String(p.sbd == null ? '' : p.sbd).replace(/\s/g, '');
  if (!/^\d{6}$/.test(raw)) throw new Error('Số báo danh gồm đúng 6 chữ số');
  var sbd = raw;
  var bai = findBai(p.maBai);
  if (!bai || bai.trangThai !== 'Đã công bố') throw new Error('Bài này chưa được công bố');
  if (laBaiNgoai(bai)) throw new Error('Bài này chỉ có điểm tổng, không có chi tiết từng câu');
  var chain = sbdChain(sbd);
  var r = kqLoad(bai.maBai).filter(function (x) { return chain.indexOf(x.sbd) !== -1; })[0];
  if (!r) throw new Error('Không có kết quả của số báo danh này trong bài');
  var cfg = parseJsonSafe(bai.cauHinh, {});
  var key = getBaiDapAn(bai.maBai)[r.maDe] || null;
  var sc = (key && r.tl) ? chamBai(cfg, r.tl, key) : null;
  return {
    tenBai: String(bai.tenBai), ngay: ngayStr(bai.ngay), maDe: r.maDe,
    tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, mx: maxDiemPhan(cfg),
    cfg: {
      p1: cfg.p1 && cfg.p1.on ? { soCau: cfg.p1.soCau, diem: cfg.p1.diem } : null,
      p2: cfg.p2 && cfg.p2.on ? { soCau: cfg.p2.soCau, d: [cfg.p2.d1, cfg.p2.d2, cfg.p2.d3, cfg.p2.d4] } : null,
      p3: cfg.p3 && cfg.p3.on ? { soCau: cfg.p3.soCau, diem: cfg.p3.diem } : null
    },
    tl: r.tl || null, key: key, dung: sc ? sc.dung : null, daSua: r.daSua
  };
}

/**
 * Góp ý / báo lỗi. Công khai (u = null) hoặc từ giáo viên đã đăng nhập.
 * Chống spam: ô ẩn "website" (máy tự điền thì bỏ), tối đa 30 góp ý công khai mỗi 10 phút.
 */
function guiGopY(p, u) {
  if (p.website) return { ok: true };
  var nd = String(p.noiDung || '').trim();
  if (nd.length < 5) throw new Error('Nội dung góp ý quá ngắn (ít nhất 5 ký tự)');
  if (nd.length > 2000) nd = nd.slice(0, 2000);
  if (!u && !toBool(getSetting('choPhepGopY', true))) throw new Error('Hiện chưa nhận góp ý qua trang này');
  var c = CacheService.getScriptCache(), n = Number(c.get('gopy_rate')) || 0;
  if (!u && n >= 30) throw new Error('Đang có quá nhiều góp ý, hãy thử lại sau ít phút');
  c.put('gopy_rate', String(n + 1), 600);

  var sh = getSystemSheet().getSheetByName('GopY');
  var H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), nCol = sh.getLastColumn();
  var row = []; for (var i = 0; i < nCol; i++) row.push('');
  var set = function (k, v) { if (H[k] !== undefined) row[H[k]] = v; };
  set('id', 'G' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5));
  set('thoiGian', nowIso());
  set('noiDung', nd);
  set('lienHe', String(p.lienHe || '').trim().slice(0, 200));
  set('daXuLy', '');
  set('nguon', u ? 'giáo viên' : 'trang tra điểm');
  set('nguoiGui', u ? u.username : '');
  set('sbd', normSbd(p.sbd) || '');
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, nCol).setNumberFormat('@');
  sh.getRange(r, 1, 1, nCol).setValues([row]);
  sysTouch('GopY');

  try {
    var to = String(getSetting('emailNhanThongBao', '') || '').trim();
    if (to) MailApp.sendEmail({ to: to, subject: '[Điểm] Góp ý mới' + (u ? ' từ ' + u.username : ''), name: 'Hệ thống điểm',
      body: nd + '\n\n' + (p.lienHe ? 'Liên hệ: ' + p.lienHe + '\n' : '') + (p.sbd ? 'SBD: ' + p.sbd + '\n' : '') + 'Nguồn: ' + (u ? 'giáo viên ' + u.username : 'trang tra điểm') });
  } catch (e) {}
  return { ok: true };
}
