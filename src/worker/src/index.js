/**
 * MÁY CHỦ MỚI: Cloudflare Worker + D1. Toàn bộ nghiệp vụ là code Apps Script cũ (app.generated.js) chạy trên lớp tương thích.
 * Một lượt gọi: (khóa nếu ghi) → nạp dữ liệu → chạy logic đồng bộ → ghi thay đổi → trả lời → việc phụ (mail, Drive, đồng bộ Sheet).
 */
import { createApp } from './app.generated.js';
import { Store, makeServices } from './shim.js';
import { Bridge } from './bridge.js';

const CORS = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'POST, GET, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type', 'Access-Control-Max-Age': '86400' };
const json = (obj, status) => new Response(JSON.stringify(obj), { status: status || 200, headers: Object.assign({ 'Content-Type': 'application/json;charset=utf-8', 'Cache-Control': 'no-store' }, CORS) });
const text = (s, status) => new Response(s, { status: status || 200, headers: Object.assign({ 'Content-Type': 'application/json;charset=utf-8', 'Cache-Control': 'no-store' }, CORS) });

/* Lệnh CHỈ ĐỌC: không cần khóa, không ghi. Phải khớp SERVER_READ trong Code.gs. */
const READ = new Set(['ping', 'me', 'bootstrap', 'publicConfig', 'traDiem', 'chiTietBai', 'listBai', 'getBai', 'getBaiFull', 'getKetQua', 'thongKeBai', 'listGroups', 'listStudents', 'listUsers', 'listTeachers', 'getCaiDat', 'listGopY', 'getNhatKy', 'previewDiem', 'ddTongQuan', 'ddPhanTich', 'ddXemBuoi', 'ngoaiDoiChieu', 'trangThaiTuDong', 'nhatKyLoi', 'validateRoster']);
/* Lệnh cần Drive: Worker kiểm tra quyền rồi nhờ cầu nối làm phần Drive */
const DRIVE = new Set(['anhInfo', 'uploadAnh', 'clearAnh', 'finishAnh', 'genDapAnFile', 'luuFileDrive']);
/* Lệnh đọc lại danh sách từ Google Sheet trước khi chạy */
const ROSTER_PULL = new Set(['syncGroups', 'validateRoster']);
const ROSTER_WRITE = new Set(['upsertStudent', 'setStudentStatus', 'transferStudent', 'syncGroups']);

const LOCK_MS = 25000;
async function acquireLock(db) {
  const deadline = Date.now() + 15000, me = String(Date.now() + LOCK_MS) + '-' + Math.random().toString(36).slice(2);
  while (Date.now() < deadline) {
    const r = await db.prepare("UPDATE kv SET v=? WHERE k='__lock' AND (v='' OR CAST(substr(v,1,13) AS INTEGER) < ?)").bind(me, Date.now()).run();
    if (r.meta.changes > 0) return me;
    await new Promise(res => setTimeout(res, 60 + Math.random() * 80));
  }
  throw new Error('Hệ thống đang bận ghi dữ liệu, thử lại sau vài giây');
}
async function releaseLock(db, me) { try { await db.prepare("UPDATE kv SET v='' WHERE k='__lock' AND v=?").bind(me).run(); } catch (e) {} }

export default {
  async fetch(request, env, ctx) {
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: CORS });
    const t0 = Date.now();
    let p = {};
    try {
      if (request.method === 'POST') { const raw = await request.text(); p = raw ? JSON.parse(raw) : {}; }
      else { const u = new URL(request.url); u.searchParams.forEach((v, k) => { p[k] = v; }); }
    } catch (e) { return json({ ok: false, error: 'Dữ liệu gửi lên không đọc được', code: 'ERR' }, 400); }
    const action = String(p.action || '');
    if (action === 'ping') return json({ ok: true, data: { ok: true, worker: true }, ms: 0, v: env.APP_VERSION || '7.0.0' });
    /* Apps Script gọi về để xác nhận mã cầu nối dùng một lần */
    if (p.bridgeVerify !== undefined) return json({ ok: true, valid: await Bridge.verify(env.DB, p.bridgeVerify) });
    /* Kiểm tra cầu nối từ đầu đến cuối (Worker → Apps Script → Worker): ai cũng gọi được, không lộ gì */
    if (action === 'kiemTraCauNoi') {
      const bridge0 = new Bridge(env);
      if (!bridge0.ok) return json({ ok: false, error: 'Chưa cấu hình APPS_SCRIPT_URL', ms: Date.now() - t0 });
      try { const r = await bridge0.ping(); return json({ ok: true, data: r, ms: Date.now() - t0 }); }
      catch (e) { return json({ ok: false, error: String(e && e.message || e), ms: Date.now() - t0 }); }
    }
    if (!action) return json({ ok: false, error: 'Thiếu action', code: 'ERR' }, 400);

    const bridge = new Bridge(env), isRead = READ.has(action);
    let lock = null, store, result, status = 200;
    try {
      if (!isRead) lock = await acquireLock(env.DB);
      store = await Store.load(env.DB);
      if (ROSTER_PULL.has(action) || (p.fresh && ['listStudents', 'listGroups'].includes(action))) await pullRoster(store, bridge, p);
      const { G, side } = makeServices(store, ctx);
      G.WORKER_MODE = true;
      const app = createApp(G);
      if (action === 'nhapTuSheet') result = await importFromSheet(p, app, store, bridge);
      else if (DRIVE.has(action)) result = await driveAction(action, p, app, store, bridge);
      else {
        let folderToTrash = '';
        if (action === 'xoaBai') { try { const b = app.findBai(p.maBai); folderToTrash = b ? String(b.driveFolderId || '') : ''; } catch (e) {} }
        const out = app.handle({ postData: { contents: JSON.stringify(p) } }, 'POST');
        result = JSON.parse(out.s);
        if (result.ok && folderToTrash && bridge.ok) ctx.waitUntil(bridge.trash(folderToTrash).catch(() => {}));
      }
      const nWrites = await store.flush();
      if (side.mails.length && bridge.ok) for (const m of side.mails) ctx.waitUntil(bridge.mail(m).catch(e => console.log('mail lỗi', e.message)));
      if (ROSTER_WRITE.has(action) && result.ok && bridge.ok) ctx.waitUntil(pushRoster(store, bridge).catch(e => console.log('đồng bộ danh sách lỗi', e.message)));
      result.ms = Date.now() - t0; result.w = nWrites; result.v = env.APP_VERSION || result.v;
    } catch (e) {
      console.log('LỖI', action, e && e.stack);
      result = { ok: false, error: String(e && e.message || e), code: 'ERR', ms: Date.now() - t0 };
    } finally {
      if (lock) await releaseLock(env.DB, lock);
    }
    return json(result, status);
  },

  /** Chạy mỗi ngày: dọn phiên, nhật ký, ảnh cũ; sao lưu toàn bộ dữ liệu về file Google Sheet cũ */
  async scheduled(event, env, ctx) {
    const bridge = new Bridge(env), lock = await acquireLock(env.DB);
    try {
      const store = await Store.load(env.DB);
      const { G } = makeServices(store, ctx); G.WORKER_MODE = true;
      const app = createApp(G);
      try { app.donDepHangNgay(); } catch (e) { console.log('dọn dẹp lỗi', e.message); }
      // ảnh cũ: gom thư mục bài đã công bố quá hạn rồi nhờ cầu nối dọn
      try {
        const days = Number(app.getSetting('soNgayGiuAnh', 30)) || 30, now = Date.now(), folders = [];
        app.sysRows('BaiKiemTra').forEach(r => { if (!r.ngayCongBo || !r.driveFolderId) return; const t = new Date(r.ngayCongBo).getTime(); if (!isNaN(t) && now - t > days * 86400000) folders.push(String(r.driveFolderId)); });
        if (folders.length && bridge.ok) await bridge.cleanupOldImages(folders);
      } catch (e) { console.log('dọn ảnh lỗi', e.message); }
      await store.flush();
      if (bridge.ok && env.BACKUP_SHEET_ID) {
        const sys = store.spreadsheet('SYS'), tabs = sys.getSheets().filter(s => s.getName() !== 'Phien').map(s => ({ name: s.getName(), values: s.getDataRange().getValues() }));
        await bridge.backup(env.BACKUP_SHEET_ID, tabs);
      }
    } finally { await releaseLock(env.DB, lock); }
  }
};

/* ---------- danh sách lớp: kéo từ Google Sheet về D1 / đẩy D1 lên Google Sheet ---------- */
function rosterIds(store) {
  const ids = new Set();
  const sys = store.spreadsheet('SYS'), nd = sys.getSheetByName('NguoiDung');
  if (nd) { const v = nd.getDataRange().getValues(), H = {}; v[0].forEach((h, i) => { H[String(h).trim()] = i; }); for (let i = 1; i < v.length; i++) { const id = String(v[i][H.rosterSheetId] || '').trim(); if (id) ids.add(id); } }
  return [...ids];
}
async function pullRoster(store, bridge, p) {
  if (!bridge.ok) return;
  for (const id of rosterIds(store)) {
    const data = await bridge.rosterRead(id);
    const ss = store.spreadsheet(id);
    for (const sh of ss.getSheets()) if (!data.tabs.some(t => t.name === sh.getName())) ss.deleteSheet(sh);
    data.tabs.forEach((t, i) => {
      let sh = ss.getSheetByName(t.name);
      if (!sh) { sh = ss.insertSheet(t.name); }
      sh.pos = i; sh.isNew = true; sh.rows = (t.values || []).map(r => r.map(v => (v === null || v === undefined ? '' : v))); sh.dirty = true; sh.structure = true;
    });
  }
}
async function pushRoster(store, bridge) {
  for (const id of rosterIds(store)) {
    const ss = store.ss.get(id); if (!ss) continue;
    const tabs = ss.getSheets().filter(s => s.dirty || s.isNew).map(s => ({ name: s.getName(), values: s.rows }));
    if (tabs.length) await bridge.rosterWrite(id, tabs);
  }
}

/* ---------- CHUYỂN DỮ LIỆU: đọc nguyên file Sheet hệ thống + file danh sách qua cầu nối, ghi đè vào D1 ----------
   Chỉ được phép khi D1 còn trống (lần cài đặt đầu) hoặc người gọi là admin. Phiên đăng nhập cũ không chuyển. */
const SYS_TABS = ['NguoiDung', 'Nhom', 'Phien', 'BaiKiemTra', 'DapAn', 'KetQua', 'NhatKySua', 'GopY', 'NhatKyLoi', 'BuoiHoc', 'DiemDanhLog', 'YeuCauSua', 'ChoXuLy', 'CaiDat'];
function trimRows(values) {
  const rows = (values || []).map(r => (r || []).map(v => (v === null || v === undefined ? '' : v)));
  while (rows.length && rows[rows.length - 1].every(v => String(v).trim() === '')) rows.pop();
  return rows;
}
function overwriteTab(ss, name, pos, rows) {
  let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
  sh.pos = pos; sh.isNew = true; sh.dirty = true; sh.structure = true; sh.rows = rows; sh.dirtyRows = new Set();
  return sh;
}
async function importFromSheet(p, app, store, bridge) {
  const sys = store.spreadsheet('SYS'), nd = sys.getSheetByName('NguoiDung');
  const empty = !nd || nd.rows.length <= 1;
  if (!empty) { const u = app.currentUser(p); if (u.role !== 'admin') return { ok: false, error: 'Chỉ quản trị mới được nhập lại dữ liệu từ Sheet', code: 'FORBIDDEN' }; }
  if (!bridge.ok) return { ok: false, error: 'Chưa cấu hình cầu nối Apps Script (APPS_SCRIPT_URL)', code: 'ERR' };
  const data = await bridge.call('sysRead', {}, { timeout: 120000 });
  const report = [];
  const byName = {}; (data.tabs || []).forEach(t => { byName[t.name] = t; });
  SYS_TABS.forEach((name, i) => {
    const t = byName[name]; if (!t) { report.push([name, 0, 'không có trong Sheet']); return; }
    let rows = trimRows(t.values);
    if (name === 'Phien') rows = rows.slice(0, 1);
    overwriteTab(sys, name, i, rows);
    report.push([name, Math.max(0, rows.length - 1), name === 'Phien' ? '(phiên cũ không chuyển)' : '']);
  });
  for (const sh of sys.getSheets()) if (!SYS_TABS.includes(sh.getName())) sys.deleteSheet(sh);
  // danh sách lớp của từng giáo viên
  const ids = rosterIds(store), roster = [];
  for (const id of ids) {
    const d = await bridge.rosterRead(id); const ss = store.spreadsheet(id);
    for (const sh of ss.getSheets()) if (!d.tabs.some(t => t.name === sh.getName())) ss.deleteSheet(sh);
    d.tabs.forEach((t, i) => { const rows = trimRows(t.values); overwriteTab(ss, t.name, i, rows); roster.push([t.name, Math.max(0, rows.length - 2)]); });
  }
  const setKv = (k, v) => { store.kv.set(k, v); store.kvDirty.set(k, v); };
  setKv('SYSTEM_SHEET_ID', 'SYS'); setKv('DA_CHUYEN_LUC', new Date().toISOString()); setKv('NGUON_SHEET_ID', String(data.sheetId || ''));
  if (p.ownerEmail) setKv('OWNER_EMAIL', String(p.ownerEmail));
  return { ok: true, data: { heThong: report, danhSach: roster, luc: store.kv.get('DA_CHUYEN_LUC') } };
}

/* ---------- lệnh cần Drive: quyền và dữ liệu ở Worker, thao tác file ở cầu nối ---------- */
async function driveAction(action, p, app, store, bridge) {
  const u = app.currentUser(p);
  if (u.role !== 'admin' && u.role !== 'teacher') return { ok: false, error: 'Không đủ quyền', code: 'FORBIDDEN' };
  const bai = app.findBai(p.maBai);
  app.requireBaiAccess(u, bai);
  const sys = store.spreadsheet('SYS'), sh = sys.getSheetByName('BaiKiemTra');
  const H = app.idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  const setCol = (name, val) => { if (H[name] !== undefined) sh.getRange(bai._row, H[name] + 1).setValue(val); };
  const ensureFolder = async () => {
    if (bai.driveFolderId) return String(bai.driveFolderId);
    const r = await bridge.ensureBaiFolder({ teacher: bai.teacher, tenBai: bai.tenBai, ngay: app.ngayStr(bai.ngay), driveFolderId: '' });
    bai.driveFolderId = r.folderId; setCol('driveFolderId', r.folderId); return r.folderId;
  };
  let data;
  if (action === 'anhInfo') {
    const n = Number(bai.soAnh) || 0;
    data = { count: n, folderUrl: bai.driveFolderId ? 'https://drive.google.com/drive/folders/' + bai.driveFolderId : '' };
  } else if (action === 'uploadAnh') {
    const imgs = p.images || [{ name: p.name, data: p.data }];
    if (!imgs.length) throw new Error('Không có ảnh để tải');
    const fid = await ensureFolder();
    const r = await bridge.uploadAnh(fid, imgs);
    setCol('soAnh', (Number(bai.soAnh) || 0) + (r.added || 0));
    data = { ok: true, added: r.added };
  } else if (action === 'clearAnh') {
    const fid = await ensureFolder();
    const r = await bridge.clearAnh(fid);
    setCol('soAnh', 0); store.kv.delete('ANHMAIL_' + bai.maBai); store.kvDirty.set('ANHMAIL_' + bai.maBai, null);
    data = { ok: true, removed: r.removed };
  } else if (action === 'finishAnh') {
    const soAnh = Number(bai.soAnh) || 0, last = Number(store.kv.get('ANHMAIL_' + bai.maBai)) || 0;
    data = { soAnh, hasDapAn: !!bai.dapAnFileId, emailSent: false };
    const to = String(app.getSetting('emailNhanThongBao', '') || '').trim();
    if (soAnh > last && to && bridge.ok) {
      const cfg = app.parseJsonSafe(bai.cauHinh, {});
      const body = 'Có ảnh phiếu bài làm mới, có thể chấm bằng TNMaker.\n\nBài: ' + bai.tenBai + '\nNgày: ' + (app.ngayStr(bai.ngay) || '(chưa ghi)') + '\nGiáo viên: ' + bai.teacher +
        '\nNhóm: ' + (app.parseNhomList(bai.cacNhom).join(', ') || '(chưa chọn)') + '\nMã đề: ' + (cfg.maDeList || []).join(', ') + '\nSố ảnh phiếu: ' + soAnh +
        '\nFile đáp án trên hệ thống: ' + (bai.dapAnFileId ? 'Có' : 'Chưa có') + '\n\nThư mục của bài trên Drive: https://drive.google.com/drive/folders/' + bai.driveFolderId + '\n';
      try { await bridge.mail({ to, subject: '[Điểm] Có ảnh phiếu mới: ' + bai.tenBai, body }); data.emailSent = true; store.kv.set('ANHMAIL_' + bai.maBai, String(soAnh)); store.kvDirty.set('ANHMAIL_' + bai.maBai, String(soAnh)); } catch (e) {}
    }
  } else if (action === 'luuFileDrive') {
    const d = String(p.data || ''); if (!d) throw new Error('File rỗng'); if (d.length > 14 * 1024 * 1024) throw new Error('File quá lớn');
    const name = String(p.tenFile || 'file.xlsx').replace(/[\\/:*?"<>|]/g, '-').trim().slice(0, 120) || 'file.xlsx';
    const mime = /\.xlsx$/i.test(name) ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/octet-stream';
    const fid = await ensureFolder();
    const r = await bridge.saveFile(fid, name, d, mime);
    data = { ok: true, url: 'https://drive.google.com/file/d/' + r.fileId + '/view' };
  } else if (action === 'genDapAnFile') {
    const b = app.getBai(p, u), cfg = b.cauHinh, maDeList = b.maDeList;
    if (!maDeList.length) throw new Error('Chưa có mã đề');
    const parts = [];
    if (cfg.p1 && cfg.p1.on) parts.push(['p1', cfg.p1.soCau, 'Phần I']);
    if (cfg.p2 && cfg.p2.on) parts.push(['p2', cfg.p2.soCau, 'Phần II']);
    if (cfg.p3 && cfg.p3.on) parts.push(['p3', cfg.p3.soCau, 'Phần III']);
    const thieu = [];
    maDeList.forEach(md => { const da = b.dapAn[md] || {}; parts.forEach(pt => { const arr = da[pt[0]] || []; for (let i = 0; i < pt[1]; i++) { const v = arr[i]; if (v === undefined || v === null || String(v).trim() === '') thieu.push('Mã ' + md + ' · ' + pt[2] + ' câu ' + (i + 1)); } }); });
    if (thieu.length) throw new Error('Còn thiếu đáp án: ' + thieu.slice(0, 8).join('; ') + (thieu.length > 8 ? ' …' : ''));
    const header = ['CAU\\MA DE'].concat(maDeList), rows = []; let cau = 1;
    parts.forEach(pt => { for (let i = 0; i < pt[1]; i++) { const row = [cau]; maDeList.forEach(md => { const arr = (b.dapAn[md] || {})[pt[0]] || []; let val = arr[i] !== undefined ? arr[i] : ''; val = pt[0] === 'p3' ? String(val).replace('.', ',') : String(val); row.push(val); }); rows.push(row); cau++; } });
    const fid = await ensureFolder();
    const name = 'dap-an - ' + app.safeName(b.tenBai) + '.xlsx';
    const r = await bridge.genDapAnFile(fid, name, header, rows);
    setCol('dapAnFileId', r.fileId);
    data = { ok: true, fileId: r.fileId, viewUrl: 'https://drive.google.com/file/d/' + r.fileId + '/view', downloadUrl: 'https://drive.google.com/uc?export=download&id=' + r.fileId, soCau: rows.length, soMaDe: maDeList.length };
  }
  return { ok: true, data };
}
