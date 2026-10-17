/**
 * CẦU NỐI về Apps Script cũ: chỉ cho việc Worker không tự làm được, đều hiếm và không cần nhanh:
 * đọc/ghi file Google Sheet danh sách lớp, Drive (ảnh phiếu, file đáp án), gửi email, sao lưu về Sheet.
 *
 * Xác thực: mỗi lượt gọi Worker sinh một mã dùng một lần (nonce) lưu trong D1 và gửi kèm;
 * Apps Script gọi ngược về Worker (?bridgeVerify=mã) để xác nhận mã đó do Worker phát ra.
 * Nhờ vậy không cần dán khóa bí mật vào hai nơi. Nếu có BRIDGE_SECRET thì dùng khóa đó thay cho nonce.
 */
export class Bridge {
  constructor(env, db) { this.url = env.APPS_SCRIPT_URL || ''; this.secret = env.BRIDGE_SECRET || ''; this.db = db || env.DB; }
  get ok() { return (/^https:\/\//.test(this.url) || /^http:\/\/(localhost|127\.0\.0\.1)[:/]/.test(this.url)) && (!!this.secret || !!this.db); }
  async token() {
    if (this.secret) return this.secret;
    const n = crypto.randomUUID().replace(/-/g, '') + crypto.randomUUID().replace(/-/g, ''), now = Date.now();
    await this.db.batch([
      this.db.prepare('DELETE FROM nonce WHERE exp < ?').bind(now),
      this.db.prepare('INSERT INTO nonce (n, exp) VALUES (?, ?)').bind(n, now + 10 * 60 * 1000)
    ]);
    return 'nonce:' + n;
  }
  /** Apps Script gọi về: mã có thật và chưa hết hạn thì trả valid:true và xóa mã (dùng một lần) */
  static async verify(db, n) {
    try {
      const r = await db.prepare('DELETE FROM nonce WHERE n = ? AND exp > ? RETURNING n').bind(String(n || ''), Date.now()).all();
      return !!(r.results && r.results.length);
    } catch (e) { return false; }
  }
  async call(op, params, opts) {
    if (!this.ok) throw new Error('Chưa cấu hình cầu nối Apps Script (APPS_SCRIPT_URL)');
    const body = JSON.stringify(Object.assign({ bridge: await this.token(), op }, params || {}));
    const ctrl = new AbortController(); const to = setTimeout(() => ctrl.abort(), (opts && opts.timeout) || 60000);
    try {
      const r = await fetch(this.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body, redirect: 'follow', signal: ctrl.signal });
      const txt = await r.text(); let j;
      try { j = JSON.parse(txt); } catch (e) { throw new Error('Cầu nối trả về dữ liệu lạ: ' + txt.slice(0, 120)); }
      if (!j.ok) throw new Error(j.error || 'Cầu nối báo lỗi');
      return j.data;
    } finally { clearTimeout(to); }
  }
  rosterRead(fileId) { return this.call('rosterRead', { fileId }); }
  rosterWrite(fileId, tabs) { return this.call('rosterWrite', { fileId, tabs }); }
  ensureBaiFolder(bai) { return this.call('ensureBaiFolder', { teacher: bai.teacher, tenBai: bai.tenBai, ngay: bai.ngay, folderId: bai.driveFolderId || '' }); }
  anhInfo(folderId) { return this.call('anhInfo', { folderId }); }
  uploadAnh(folderId, images) { return this.call('uploadAnh', { folderId, images }, { timeout: 120000 }); }
  clearAnh(folderId) { return this.call('clearAnh', { folderId }); }
  saveFile(folderId, name, data, mime) { return this.call('saveFile', { folderId, name, data, mime }, { timeout: 120000 }); }
  genDapAnFile(folderId, name, header, rows) { return this.call('genDapAnFile', { folderId, name, header, rows }); }
  trash(id) { return this.call('trash', { id }); }
  mail(m) { return this.call('mail', { to: m.to, subject: m.subject, body: m.body, name: m.name || 'Hệ thống điểm' }); }
  backup(sheetId, tabs) { return this.call('backup', { sheetId, tabs }, { timeout: 120000 }); }
  cleanupOldImages(folders) { return this.call('cleanupOldImages', { folders }); }
  ping() { return this.call('ping', {}, { timeout: 30000 }); }
}
