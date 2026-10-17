/* =====================================================================
   ĐIỂM TỪ NGUỒN KHÁC: bài chỉ có điểm tổng (học kỳ ở trường, bài khác)
   ===================================================================== */
'use strict';

function cfgOf(b) { var c = b && b.cauHinh; if (typeof c === 'string') { try { c = JSON.parse(c); } catch (e) { c = {}; } } return c || {}; }
function baiNgoai(b) { return cfgOf(b).loai === 'ngoai'; }

/* ---------- tạo bài ---------- */
function screenNgoaiMoi() {
  var v = V(), gs = App.state.groups || [], chon = {};
  swap(v, '<div class="page"><a class="back" href="#/bai">' + icon('chevL') + 'Bài kiểm tra</a>' +
    '<div class="page-head"><div class="titles"><h1>Thêm điểm từ nguồn khác</h1><div class="sub">Cho các bài chỉ có họ tên và điểm tổng: kiểm tra học kỳ ở trường, bài thi thử, bài làm ở nhà… Học sinh tra lại được bằng SBD như bài thường.</div></div></div>' +
    '<div class="panel stack" style="max-width:720px">' +
    '<div class="field"><label>Tên bài</label><input class="input" id="nTen" placeholder="vd: Kiểm tra cuối học kỳ I (ở trường)"></div>' +
    '<div class="grid grid-2"><div class="field"><label>Ngày</label><input class="input" type="date" id="nNgay"></div>' +
    '<div class="field"><label>Thang điểm</label><input class="input" type="number" id="nThang" value="10" min="1" max="100"></div></div>' +
    '<div class="field"><label>Loại điểm (hiện cho phụ huynh)</label><input class="input" id="nNguon" value="Học kỳ ở trường" list="nNguonL"><datalist id="nNguonL"><option>Học kỳ ở trường</option><option>Giữa kỳ ở trường</option><option>Thi thử</option><option>Bài làm ở nhà</option></datalist></div>' +
    '<div class="field"><label>Các nhóm có điểm</label><div class="row" id="nNhom">' + gs.map(function (g) { return '<label class="chip"><input type="checkbox" value="' + esc(g.maNhom) + '"><span>' + esc(g.tenNhom) + '</span></label>'; }).join('') + '</div><div class="hint">Web chỉ tìm tên học sinh trong các nhóm được chọn.</div></div>' +
    '<div class="row"><button class="btn btn-primary btn-lg" id="nTao">' + icon('plus') + '<span>Tạo và nhập điểm</span></button></div></div></div>');
  $('#nNgay').value = new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  $('#nTao').onclick = function () {
    var nhom = $$('#nNhom input:checked').map(function (x) { return x.value; });
    busy(this, async function () {
      if (!$('#nTen').value.trim()) throw ApiError('ERR', 'Nhập tên bài');
      if (!nhom.length) throw ApiError('ERR', 'Chọn ít nhất một nhóm');
      var r = await Api.call('createBai', { tenBai: $('#nTen').value.trim(), ngay: $('#nNgay').value, cacNhom: nhom, cauHinh: { loai: 'ngoai', thang: Number($('#nThang').value) || 10, nguon: $('#nNguon').value } });
      Router.go('/diem-ngoai/' + r.maBai + '/nhap');
    });
  };
}

/* ---------- trang của bài điểm tổng ---------- */
function ngoaiHub(v, b) {
  var c = cfgOf(b), pub = b.trangThai === 'Đã công bố', n = b.soKetQua || (b.counts && b.counts.ketQua) || 0, cho = b.soChoXuLy || (b.counts && b.counts.choXuLy) || 0;
  swap(v, '<div class="page"><a class="back" href="#/bai">' + icon('chevL') + 'Bài kiểm tra</a>' +
    '<div class="page-head"><div class="titles"><h1>' + esc(b.tenBai) + ' ' + (pub ? '<span class="badge ok">Đã công bố</span>' : '<span class="badge">Chưa công bố</span>') + '</h1>' +
    kv([['Loại', c.nguon || 'Điểm từ nguồn khác'], ['Ngày', fmtDate(b.ngay)], ['Nhóm', (b.cacNhom || []).join(', ')], ['Thang điểm', c.thang || 10]]) + '</div></div>' +
    '<div class="grid grid-2">' +
    '<div class="panel stack"><h3>1. Nhập điểm</h3><p class="small muted">Dán từ Excel hoặc Google Sheet, hoặc chọn file. Chỉ cần cột họ tên và cột điểm.</p>' +
    '<div class="row"><a class="btn btn-primary" href="#/diem-ngoai/' + esc(b.maBai) + '/nhap">' + icon('upload') + '<span>' + (n ? 'Nhập lại hoặc thêm điểm' : 'Nhập điểm') + '</span></a>' +
    '<a class="btn" href="#/bai/' + esc(b.maBai) + '/ket-qua">' + icon('eye') + '<span>Xem, sửa điểm</span></a></div></div>' +
    '<div class="panel stack"><h3>2. Công bố cho học sinh tra</h3><p class="small muted">' + (pub ? 'Học sinh, phụ huynh đang tra được điểm bài này.' : 'Công bố xong, học sinh nhập SBD là thấy điểm bài này.') + '</p>' +
    '<div class="row"><button class="btn ' + (pub ? '' : 'btn-primary') + '" id="ngPub">' + icon(pub ? 'eyeOff' : 'send') + '<span>' + (pub ? 'Ẩn điểm' : 'Công bố') + '</span></button>' +
    '<button class="btn btn-ghost" id="ngDel">' + icon('trash') + '<span>Xóa bài</span></button></div></div></div></div>');
  $('#ngPub').onclick = function () {
    var btn = this;
    busy(btn, async function () { await Api.call('setCongBo', { maBai: b.maBai, congBo: !pub }); if (!pub) celebrateAt(btn); Api.invalidate(); await screenBaiHub({ id: b.maBai }); }, { ok: pub ? 'Đã ẩn điểm' : 'Đã công bố' });
  };
  $('#ngDel').onclick = async function () {
    if (!(await Modal.confirm({ title: 'Xóa bài "' + b.tenBai + '"?', danger: true, confirmText: 'Xóa', typeToConfirm: b.tenBai, html: 'Xóa cả điểm đã nhập của bài này. Gõ lại đúng tên bài để xác nhận.' }))) return;
    busy(null, async function () { await Api.call('xoaBai', { maBai: b.maBai, xacNhan: b.tenBai }); Router.go('/bai'); }, { ok: 'Đã xóa bài' });
  };
}

/* ---------- đọc bảng dán / file ---------- */
function ngTachBang(text) {
  var lines = String(text || '').replace(/\r/g, '').split('\n').filter(function (l) { return l.trim(); });
  var sep = lines.some(function (l) { return l.indexOf('\t') !== -1; }) ? '\t' : (lines.some(function (l) { return l.indexOf(';') !== -1; }) ? ';' : ',');
  return lines.map(function (l) { return l.split(sep).map(function (x) { return x.trim().replace(/^"|"$/g, ''); }); });
}
function ngLaDiem(x, thang) { var s = String(x).trim().replace(',', '.'); return s !== '' && isFinite(Number(s)) && Number(s) >= 0 && Number(s) <= thang; }
/** Đoán cột: tên, điểm, SBD, lớp; nhận cả bảng tách Họ | Tên */
function ngDoanCot(rows, thang) {
  var n = Math.max.apply(null, rows.map(function (r) { return r.length; })), H = rows[0] || [];
  var head = H.map(function (h) { return String(h).toLowerCase(); });
  var coHead = head.some(function (h) { return /tên|ten|điểm|diem|sbd|lớp|lop|họ/.test(h); });
  var data = coHead ? rows.slice(1) : rows;
  var st = []; for (var c = 0; c < n; c++) {
    var col = data.map(function (r) { return r[c] === undefined ? '' : String(r[c]); }).filter(function (x) { return x.trim(); });
    var stt = data.every(function (r, i) { return String(r[c]).trim() === String(i + 1); });
    st.push({ c: c, n: col.length, diem: col.filter(function (x) { return ngLaDiem(x, thang); }).length, sbd: col.filter(function (x) { return /^\d{6}$/.test(x.trim()); }).length,
      chu: col.filter(function (x) { return /[a-zà-ỹđ]/i.test(x) && !/^\d/.test(x.trim()); }).length, mot: col.filter(function (x) { return x.trim().split(/\s+/).length === 1; }).length,
      lop: col.filter(function (x) { return /^\d{1,2}\s*[a-zđ]/i.test(x.trim()) && x.trim().length <= 6; }).length, stt: stt, h: head[c] || '' });
  }
  var by = function (k, ok) { var c = st.filter(ok || function () { return true; }).sort(function (a, b) { return b[k] - a[k] || b.c - a.c; })[0]; return c && c[k] > 0 ? c.c : -1; };
  var cot = { ten: -1, ten2: -1, diem: -1, sbd: -1, lop: -1, coHead: coHead };
  st.forEach(function (s) {
    if (/sbd|báo danh/.test(s.h)) cot.sbd = s.c; else if (/điểm|diem/.test(s.h)) cot.diem = s.c; else if (/^lớp|^lop/.test(s.h)) cot.lop = s.c;
    else if (/^tên$|^ten$/.test(s.h)) cot.ten2 = s.c; else if (/họ|tên|ten/.test(s.h) && cot.ten < 0) cot.ten = s.c;
  });
  if (cot.sbd < 0) cot.sbd = by('sbd', function (s) { return s.sbd >= s.n * 0.6; });
  if (cot.lop < 0) cot.lop = by('lop', function (s) { return s.lop >= s.n * 0.6 && s.c !== cot.sbd; });
  if (cot.ten < 0) cot.ten = by('chu', function (s) { return s.c !== cot.lop && s.chu >= s.n * 0.6; });
  if (cot.ten >= 0 && cot.ten2 < 0) { var nx = st[cot.ten + 1]; if (nx && nx.chu >= nx.n * 0.6 && nx.mot >= nx.n * 0.8 && st[cot.ten].mot < st[cot.ten].n * 0.5) cot.ten2 = nx.c; }
  if (cot.diem < 0) cot.diem = by('diem', function (s) { return !s.stt && s.c !== cot.sbd && s.c !== cot.ten && s.c !== cot.ten2 && s.c !== cot.lop && s.diem >= s.n * 0.6; });
  cot.soCot = n;
  return cot;
}
function ngLayDong(rows, cot) {
  var data = cot.coHead ? rows.slice(1) : rows;
  return data.map(function (r) {
    var g = function (i) { return i >= 0 && r[i] !== undefined ? String(r[i]).trim() : ''; };
    return { ten: (g(cot.ten) + ' ' + g(cot.ten2)).trim(), diem: g(cot.diem), sbd: g(cot.sbd), lop: g(cot.lop) };
  }).filter(function (x) { return x.ten || x.sbd || x.diem; });
}

/* ---------- màn nhập điểm ---------- */
async function screenNgoaiNhap(params) {
  var v = V(), id = params.id;
  swap(v, skPage());
  var b = await Api.call('getBaiFull', { maBai: id });
  var c = cfgOf(b), thang = Number(c.thang) || 10, st = { rows: null, cot: null, ten: '' };
  var draw = function () {
    var cot = st.cot, opts = function (sel, none) {
      return '<option value="-1">' + none + '</option>' + Array.from({ length: cot.soCot }, function (_, i) {
        var mau = st.rows.slice(cot.coHead ? 1 : 0, (cot.coHead ? 1 : 0) + 2).map(function (r) { return r[i] || ''; }).join(', ');
        return '<option value="' + i + '" ' + (i === sel ? 'selected' : '') + '>Cột ' + (i + 1) + (cot.coHead && st.rows[0][i] ? ' "' + esc(st.rows[0][i]) + '"' : '') + ': ' + esc(mau.slice(0, 40)) + '</option>';
      }).join('');
    };
    $('#ngCot').innerHTML = !st.rows ? '' : '<div class="panel stack" style="margin-top:14px"><h3>2. Kiểm tra cột web tự nhận</h3>' +
      '<div class="grid grid-2"><div class="field"><label>Cột họ tên</label><select class="input" data-k="ten">' + opts(cot.ten, '(không có)') + '</select></div>' +
      '<div class="field"><label>Cột tên (nếu tách riêng họ, tên)</label><select class="input" data-k="ten2">' + opts(cot.ten2, '(không tách)') + '</select></div>' +
      '<div class="field"><label>Cột điểm</label><select class="input" data-k="diem">' + opts(cot.diem, '(chưa chọn)') + '</select></div>' +
      '<div class="field"><label>Cột SBD (nếu có)</label><select class="input" data-k="sbd">' + opts(cot.sbd, '(không có)') + '</select></div>' +
      '<div class="field"><label>Cột lớp (nếu có, giúp phân biệt trùng tên)</label><select class="input" data-k="lop">' + opts(cot.lop, '(không có)') + '</select></div></div>' +
      '<label class="switch"><input type="checkbox" id="ngHead" ' + (cot.coHead ? 'checked' : '') + '><span class="track"></span><span>Dòng đầu là tiêu đề cột</span></label>' +
      '<div class="row"><button class="btn btn-primary btn-lg" id="ngDc">' + icon('check') + '<span>Đối chiếu với danh sách lớp</span></button><span class="small muted">' + ngLayDong(st.rows, cot).length + ' dòng</span></div></div>';
    if (!st.rows) return;
    $$('#ngCot select[data-k]').forEach(function (s) { s.onchange = function () { st.cot[s.dataset.k] = Number(s.value); draw(); }; });
    $('#ngHead').onchange = function () { st.cot.coHead = this.checked; draw(); };
    $('#ngDc').onclick = function () {
      if (st.cot.diem < 0 || (st.cot.ten < 0 && st.cot.sbd < 0)) return toast('Chọn cột điểm và cột họ tên (hoặc SBD)', 'error');
      var rows = ngLayDong(st.rows, st.cot);
      busy(this, async function () { var r = await Api.call('ngoaiDoiChieu', { maBai: id, rows: rows }); ngoaiSoat(v, b, r, st.ten); });
    };
  };
  var nap = function (rows, ten) {
    if (!rows.length) return toast('Không thấy dòng nào', 'error');
    st.rows = rows; st.ten = ten || 'dán từ bảng tính'; st.cot = ngDoanCot(rows, thang); draw();
    $('#ngCot').scrollIntoView && $('#ngCot').scrollIntoView({ behavior: reduceMotion ? 'auto' : 'smooth', block: 'start' });
  };
  swap(v, '<div class="page"><a class="back" href="#/bai/' + esc(id) + '">' + icon('chevL') + esc(b.tenBai) + '</a>' +
    '<div class="page-head"><div class="titles"><h1>Nhập điểm</h1>' + kv([['Bài', b.tenBai], ['Nhóm', (b.cacNhom || []).join(', ')], ['Thang điểm', thang]]) + '</div></div>' +
    '<div class="panel stack"><h3>1. Đưa bảng điểm vào</h3>' +
    '<p class="small muted">Cách nhanh nhất: mở file Excel/Google Sheet, bôi đen các cột <b>Họ tên</b> và <b>Điểm</b> (có cả dòng tiêu đề cũng được), Ctrl+C rồi dán vào ô dưới. Có cột SBD hoặc Lớp thì bôi đen luôn, web khớp chắc hơn.</p>' +
    '<textarea class="input" id="ngDan" rows="7" placeholder="Nguyễn Văn An&#9;8,5&#10;Trần Thị Bình&#9;9"></textarea>' +
    '<div class="row"><button class="btn btn-primary" id="ngDocDan">' + icon('check') + '<span>Đọc bảng vừa dán</span></button><span class="small muted">hoặc</span>' +
    '<label class="btn">' + icon('upload') + '<span>Chọn file Excel, CSV</span><input type="file" id="ngFile" accept=".xlsx,.xls,.csv" hidden></label></div></div>' +
    '<div id="ngCot"></div><div id="ngSoat"></div></div>');
  $('#ngDocDan').onclick = function () { nap(ngTachBang($('#ngDan').value), 'dán từ bảng tính'); };
  $('#ngFile').onchange = function () {
    var f = this.files[0]; if (!f) return;
    busy(null, async function () {
      var X = await needXLSX(), wb = X.read(await f.arrayBuffer(), { type: 'array' });
      var rows = X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: '' }).map(function (r) { return r.map(function (x) { return String(x).trim(); }); }).filter(function (r) { return r.some(function (x) { return x; }); });
      nap(rows, f.name);
    });
  };
}

var NG_TT = { 'khop': ['ok', 'Khớp'], 'gan-dung': ['warn', 'Gần đúng'], 'trung-ten': ['red', 'Trùng tên'], 'khong-thay': ['red', 'Không thấy'], 'loi-diem': ['red', 'Điểm sai'] };

/* ---------- bảng soát trước khi lưu ---------- */
function ngoaiSoat(v, b, r, tenFile) {
  var all = r.hocSinh, ten = {}; all.forEach(function (h) { ten[h.sbd] = h; });
  var used = {}, chon = r.rows.map(function (o) {                // mặc định: dòng chắc chắn gán luôn; dòng trùng em thì dòng sau để chờ
    if ((o.trangThai === 'khop' || o.trangThai === 'gan-dung') && !used[o.sbd]) { used[o.sbd] = 1; return o.sbd; }
    return o.trangThai === 'loi-diem' ? 'bo' : 'cho';
  });
  var dem = function () { var d = { gan: 0, cho: 0, bo: 0 }; chon.forEach(function (x) { if (x === 'cho') d.cho++; else if (x === 'bo') d.bo++; else d.gan++; }); return d; };
  var opt = function (o, i) {
    var ds = (o.ungVien || []).concat(all.filter(function (h) { return !(o.ungVien || []).some(function (u) { return u.sbd === h.sbd; }); }));
    if (o.sbd && !ds.some(function (h) { return h.sbd === o.sbd; })) ds.unshift({ sbd: o.sbd, ten: o.tenDs || '', lop: '' });
    return '<select class="input input-sm" data-i="' + i + '" style="max-width:300px">' +
      (o.trangThai === 'loi-diem' ? '' : ds.map(function (h, k) { return (k === (o.ungVien || []).length && o.ungVien && o.ungVien.length ? '<option disabled>──────────</option>' : '') + '<option value="' + esc(h.sbd) + '" ' + (chon[i] === h.sbd ? 'selected' : '') + '>' + esc(h.sbd + ' ' + h.ten + (h.lop ? ' (' + h.lop + ')' : '')) + '</option>'; }).join('')) +
      '<option value="cho" ' + (chon[i] === 'cho' ? 'selected' : '') + '>Để vào "chờ xử lý"</option><option value="bo" ' + (chon[i] === 'bo' ? 'selected' : '') + '>Bỏ dòng này</option></select>';
  };
  var veDem = function () { var d = dem(); $('#ngDem').innerHTML = kv([['Gán cho học sinh', d.gan, 'good'], ['Để chờ xử lý', d.cho, d.cho ? 'warn' : ''], ['Bỏ', d.bo]]); };
  var thuTu = { 'loi-diem': 0, 'trung-ten': 1, 'khong-thay': 2, 'gan-dung': 3, 'khop': 4 };
  var idx = r.rows.map(function (o, i) { return i; }).sort(function (a, z) { return thuTu[r.rows[a].trangThai] - thuTu[r.rows[z].trangThai] || a - z; });
  var counts = {}; r.rows.forEach(function (o) { counts[o.trangThai] = (counts[o.trangThai] || 0) + 1; });
  $('#ngSoat').innerHTML = '<div class="panel stack" style="margin-top:14px"><div class="row"><h3 class="grow">3. Soát lại rồi lưu</h3>' +
    Object.keys(NG_TT).filter(function (k) { return counts[k]; }).map(function (k) { return '<span class="badge ' + NG_TT[k][0] + '">' + NG_TT[k][1] + ': ' + counts[k] + '</span>'; }).join(' ') + '</div>' +
    '<p class="small muted">Dòng "Khớp" đã gán sẵn. "Gần đúng" là tên gõ thiếu dấu, khác hoa thường: kiểm tra lại em được chọn. "Trùng tên", "Không thấy": chọn đúng em, hoặc để vào "chờ xử lý" rồi xử lý sau.</p>' +
    '<div class="table-wrap"><table class="tbl tbl-cards"><thead><tr><th>Trong file</th><th class="r">Điểm</th><th>Tình trạng</th><th>Gán cho</th></tr></thead><tbody>' +
    idx.map(function (i) {
      var o = r.rows[i], t = NG_TT[o.trangThai];
      return '<tr><td><b style="font-weight:600">' + esc(o.ten || '(không tên)') + '</b>' + (o.lop ? '<div class="small muted">Lớp ' + esc(o.lop) + '</div>' : '') + (o.sbdFile ? '<div class="small muted sbd">' + esc(o.sbdFile) + '</div>' : '') + '</td>' +
        '<td class="r num"><b>' + (o.trangThai === 'loi-diem' ? '<span style="color:var(--red)">' + esc(o.diemGoc || '(trống)') + '</span>' : fmtScore(o.diem)) + '</b></td>' +
        '<td><span class="badge ' + t[0] + ' plain">' + t[1] + '</span>' + (o.trungDong ? '<div class="small" style="color:var(--red)">Em này có ở dòng khác</div>' : '') + (o.trangThai === 'loi-diem' ? '<div class="small muted">phải từ 0 đến ' + r.thang + '</div>' : '') + '</td>' +
        '<td>' + opt(o, i) + '</td></tr>';
    }).join('') + '</tbody></table></div>' +
    (r.daCo ? '<div class="field"><label>Bài này đã có ' + r.daCo + ' em có điểm</label><div class="seg" id="ngCheDo"><button data-m="thay" class="active">Thay toàn bộ điểm cũ</button><button data-m="gop">Gộp thêm (giữ điểm cũ)</button></div></div>' : '') +
    '<div class="row"><div class="grow" id="ngDem"></div><button class="btn btn-primary btn-lg" id="ngLuu">' + icon('check') + '<span>Lưu điểm</span></button></div></div>';
  veDem();
  var cheDo = 'thay';
  $$('#ngCheDo button').forEach(function (bt) { bt.onclick = function () { cheDo = bt.dataset.m; $$('#ngCheDo button').forEach(function (x) { x.classList.toggle('active', x === bt); }); }; });
  $$('#ngSoat select[data-i]').forEach(function (s) { s.onchange = function () { chon[Number(s.dataset.i)] = s.value; veDem(); }; });
  $('#ngSoat').scrollIntoView && $('#ngSoat').scrollIntoView({ behavior: reduceMotion ? 'auto' : 'smooth', block: 'start' });
  $('#ngLuu').onclick = function () {
    var btn = this, gan = [], cho = [], seen = {};
    for (var i = 0; i < r.rows.length; i++) {
      var o = r.rows[i], x = chon[i];
      if (x === 'bo') continue;
      if (x === 'cho') { cho.push({ ten: o.ten, diem: o.diem, lop: o.lop }); continue; }
      if (seen[x]) return toast('Em ' + (ten[x] ? ten[x].ten : x) + ' đang được gán 2 dòng điểm. Sửa lại một dòng.', 'error');
      seen[x] = 1; gan.push({ sbd: x, diem: o.diem });
    }
    busy(btn, async function () {
      var res = await Api.call('ngoaiLuu', { maBai: b.maBai, gan: gan, cho: cho, cheDo: cheDo, tenFile: tenFile });
      celebrateAt(btn);
      toast('Đã lưu điểm ' + res.daGhi + ' em' + (res.cho ? ', ' + res.cho + ' dòng chờ xử lý' : ''), 'success');
      setTimeout(function () { Router.go('/bai/' + b.maBai + (res.cho ? '/ket-qua/cho' : '')); }, 700);
    });
  };
}

/* ---------- sửa điểm một em ---------- */
function ngoaiEditModal(s, b, after) {
  var c = cfgOf(b);
  Modal.open({ title: 'Sửa điểm: ' + ((s.ho || '') + ' ' + (s.ten || '')).trim(), body: '<div class="stack">' + kv([['SBD', s.sbd], ['Điểm hiện tại', s.kq ? fmtScore(s.kq.tong) : 'chưa có']]) +
    '<div class="field"><label>Điểm mới (0 đến ' + (c.thang || 10) + ')</label><input class="input" id="neD" inputmode="decimal" value="' + (s.kq ? String(s.kq.tong).replace('.', ',') : '') + '"></div>' +
    '<div class="field"><label>Lý do</label><input class="input" id="neL" placeholder="vd: nhập nhầm"></div></div>',
    actions: [{ label: 'Hủy', kind: 'btn-ghost' }, { label: 'Lưu', kind: 'btn-primary', onClick: async function (m) {
      if (s.kq) await Api.call('ngoaiSuaDiem', { maBai: b.maBai, sbd: s.sbd, diem: m.$('#neD').value, lyDo: m.$('#neL').value });
      else await Api.call('ngoaiLuu', { maBai: b.maBai, cheDo: 'gop', gan: [{ sbd: s.sbd, diem: m.$('#neD').value }], cho: [] });
      m.close(); toast('Đã lưu', 'success'); if (after) after();
    } }] });
}

window.registerNgoaiRoutes = function () {
  Router.on('/diem-ngoai/moi', function () { return isAdmin() ? Router.go('/bai') : screenNgoaiMoi(); }, { nav: 'bai' });
  Router.on('/diem-ngoai/:id/nhap', screenNgoaiNhap, { nav: 'bai' });
};
