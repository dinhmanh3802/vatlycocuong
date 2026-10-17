/**
 * LỚP TƯƠNG THÍCH: cho code máy chủ viết cho Google Apps Script chạy nguyên vẹn trên Cloudflare Worker.
 *
 * Ý tưởng: mỗi "tờ Sheet" là một bảng dòng trong D1. Đầu lượt gọi nạp TOÀN BỘ dòng vào bộ nhớ (dữ liệu
 * nhỏ: vài nghìn dòng, vài trăm KB), logic cũ chạy đồng bộ trên bộ nhớ y như trên Sheet, cuối lượt ghi
 * các tờ đã đổi xuống D1 trong một giao dịch. Một lượt = 1 truy vấn đọc + 1 giao dịch ghi.
 */
import { sha256Hex } from './sha256.js';

/* =========================== kho dữ liệu =========================== */
export class Store {
  constructor(db) { this.db = db; this.ss = new Map(); this.kv = new Map(); this.kvDirty = new Map(); this.sheetsDirty = new Set(); }

  static async load(db) {
    const st = new Store(db);
    const [rows, sheets, kv] = await Promise.all([
      db.prepare('SELECT ss, tab, rn, data FROM rows ORDER BY ss, tab, rn').all(),
      db.prepare('SELECT ss, tab, pos FROM sheets ORDER BY ss, pos').all(),
      db.prepare('SELECT k, v FROM kv').all()
    ]);
    for (const s of sheets.results) st.spreadsheet(s.ss).addSheet(s.tab, s.pos);
    for (const r of rows.results) {
      const sh = st.spreadsheet(r.ss).getSheetByName(r.tab) || st.spreadsheet(r.ss).addSheet(r.tab);
      sh.rows[r.rn - 1] = JSON.parse(r.data);
    }
    for (const s of st.ss.values()) for (const sh of s.sheets.values()) { for (let i = 0; i < sh.rows.length; i++) if (!sh.rows[i]) sh.rows[i] = []; sh.dirty = false; sh.structure = false; }
    for (const x of kv.results) st.kv.set(x.k, x.v);
    return st;
  }

  spreadsheet(id) {
    if (!this.ss.has(id)) this.ss.set(id, new Spreadsheet(this, id));
    return this.ss.get(id);
  }

  /** Các câu lệnh ghi xuống D1 cho mọi thay đổi trong lượt này */
  flushStatements() {
    const S = [];
    for (const s of this.ss.values()) {
      for (const sh of s.sheets.values()) {
        if (sh.removed) { S.push(this.db.prepare('DELETE FROM rows WHERE ss=? AND tab=?').bind(s.id, sh.name)); S.push(this.db.prepare('DELETE FROM sheets WHERE ss=? AND tab=?').bind(s.id, sh.name)); continue; }
        if (sh.isNew) S.push(this.db.prepare('INSERT OR REPLACE INTO sheets(ss,tab,pos) VALUES(?,?,?)').bind(s.id, sh.name, sh.pos));
        if (!sh.dirty && !sh.isNew) continue;
        if (sh.structure || sh.isNew) {                      // xóa/chèn dòng: ghi lại cả tờ (tờ chỉ vài trăm dòng)
          S.push(this.db.prepare('DELETE FROM rows WHERE ss=? AND tab=?').bind(s.id, sh.name));
          sh.rows.forEach((r, i) => S.push(this.db.prepare('INSERT INTO rows(ss,tab,rn,data) VALUES(?,?,?,?)').bind(s.id, sh.name, i + 1, JSON.stringify(r))));
        } else {
          for (const rn of sh.dirtyRows) {
            if (rn > sh.rows.length) S.push(this.db.prepare('DELETE FROM rows WHERE ss=? AND tab=? AND rn=?').bind(s.id, sh.name, rn));
            else S.push(this.db.prepare('INSERT OR REPLACE INTO rows(ss,tab,rn,data) VALUES(?,?,?,?)').bind(s.id, sh.name, rn, JSON.stringify(sh.rows[rn - 1])));
          }
        }
      }
    }
    for (const [k, v] of this.kvDirty) S.push(v === null ? this.db.prepare('DELETE FROM kv WHERE k=?').bind(k) : this.db.prepare('INSERT OR REPLACE INTO kv(k,v) VALUES(?,?)').bind(k, v));
    return S;
  }
  async flush() {
    const S = this.flushStatements();
    if (S.length) { for (let i = 0; i < S.length; i += 90) await this.db.batch(S.slice(i, i + 90)); }
    return S.length;
  }
  hasChanges() { return this.flushStatements().length > 0; }
}

class Spreadsheet {
  constructor(store, id) { this.store = store; this.id = id; this.sheets = new Map(); }
  addSheet(name, pos) { const sh = new Sheet(this, name, pos === undefined ? this.sheets.size : pos); this.sheets.set(name, sh); return sh; }
  getId() { return this.id; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id; }
  getName() { return this.id; }
  getSheetByName(n) { const s = this.sheets.get(n); return s && !s.removed ? s : null; }
  getSheets() { return [...this.sheets.values()].filter(s => !s.removed).sort((a, b) => a.pos - b.pos); }
  insertSheet(name) { if (this.getSheetByName(name)) throw new Error('Đã có tờ ' + name); const sh = this.addSheet(name); sh.isNew = true; sh.dirty = true; return sh; }
  deleteSheet(sh) { sh.removed = true; }
}

class Sheet {
  constructor(ss, name, pos) { this.ss = ss; this.name = name; this.pos = pos; this.rows = []; this.dirty = false; this.structure = false; this.dirtyRows = new Set(); this.isNew = false; this.removed = false; }
  getName() { return this.name; }
  setName(n) { this.ss.sheets.delete(this.name); this.name = n; this.ss.sheets.set(n, this); this.isNew = true; this.dirty = true; return this; }
  getParent() { return this.ss; }
  getSheetId() { return this.pos; }
  getLastRow() { let n = this.rows.length; while (n > 0 && this.rows[n - 1].every(v => v === '' || v === null || v === undefined)) n--; return n; }
  getLastColumn() { let m = 0; for (const r of this.rows) { let k = r.length; while (k > 0 && (r[k - 1] === '' || r[k - 1] === null || r[k - 1] === undefined)) k--; if (k > m) m = k; } return m; }
  getMaxRows() { return Math.max(1000, this.rows.length); }
  getMaxColumns() { return Math.max(26, this.getLastColumn()); }
  _touch(rn) { this.dirty = true; this.dirtyRows.add(rn); }
  _get(r, c) { const row = this.rows[r - 1]; const v = row ? row[c - 1] : undefined; return v === undefined || v === null ? '' : v; }
  _set(r, c, v) { while (this.rows.length < r) this.rows.push([]); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v === undefined || v === null ? '' : v; this._touch(r); }
  getRange(a, b, c, d) {
    if (typeof a === 'string') throw new Error('getRange(A1) không hỗ trợ: ' + a);
    return new Range(this, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
  }
  getDataRange() { const lr = Math.max(1, this.getLastRow()), lc = Math.max(1, this.getLastColumn()); return new Range(this, 1, 1, lr, lc); }
  appendRow(vals) { const r = this.getLastRow() + 1; this.rows.length = Math.max(this.rows.length, r - 1); this.rows[r - 1] = vals.map(v => v === undefined || v === null ? '' : v); this._touch(r); return this; }
  deleteRow(r) { return this.deleteRows(r, 1); }
  deleteRows(r, n) { this.rows.splice(r - 1, n || 1); this.dirty = true; this.structure = true; return this; }
  insertRowsAfter(r, n) { const a = []; for (let i = 0; i < (n || 1); i++) a.push([]); this.rows.splice(r, 0, ...a); this.dirty = true; this.structure = true; return this; }
  clear() { this.rows = []; this.dirty = true; this.structure = true; return this; }
  clearContents() { return this.clear(); }
  setFrozenRows() { return this; } autoResizeColumns() { return this; } setColumnWidth() { return this; } hideColumns() { return this; }
}

class Range {
  constructor(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh._get(this.r + i, this.c + j)); out.push(row); } return out; }
  getDisplayValues() { return this.getValues().map(r => r.map(v => String(v))); }
  getValue() { return this.sh._get(this.r, this.c); }
  setValues(vals) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh._set(this.r + i, this.c + j, vals[i] ? vals[i][j] : ''); return this; }
  setValue(v) { this.sh._set(this.r, this.c, v); return this; }
  setNumberFormat() { return this; } setNumberFormats() { return this; } setFontWeight() { return this; } setBackground() { return this; } setWrap() { return this; }
  clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh._set(this.r + i, this.c + j, ''); return this; }
  getRow() { return this.r; } getColumn() { return this.c; } getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
}

/* =========================== bộ nhớ đệm (sống theo máy ảo Worker) =========================== */
const MEM = new Map();
export const CacheService = {
  getScriptCache() {
    return {
      get(k) { const e = MEM.get(k); if (!e) return null; if (e.exp && e.exp < Date.now()) { MEM.delete(k); return null; } return e.v; },
      put(k, v, ttl) { MEM.set(k, { v: String(v), exp: Date.now() + (ttl || 600) * 1000 }); },
      remove(k) { MEM.delete(k); },
      putAll(m, ttl) { for (const k in m) this.put(k, m[k], ttl); },
      getAll(ks) { const o = {}; for (const k of ks) { const v = this.get(k); if (v !== null) o[k] = v; } return o; },
      removeAll(ks) { for (const k of ks) MEM.delete(k); }
    };
  },
  getUserCache() { return this.getScriptCache(); }, getDocumentCache() { return this.getScriptCache(); }
};

/* =========================== dựng bộ dịch vụ cho một lượt gọi =========================== */
export function makeServices(store, ctx) {
  const side = { mails: [], drive: [] };          // việc phụ chạy sau khi trả lời (gửi mail, Drive)
  const props = {
    getProperty: k => (store.kv.has(k) ? store.kv.get(k) : null),
    setProperty(k, v) { store.kv.set(k, String(v)); store.kvDirty.set(k, String(v)); return this; },
    deleteProperty(k) { store.kv.delete(k); store.kvDirty.set(k, null); return this; },
    getProperties() { const o = {}; for (const [k, v] of store.kv) o[k] = v; return o; },
    getKeys() { return [...store.kv.keys()]; }
  };
  const dummyFolder = (id) => ({
    getId: () => id || '', getUrl: () => 'https://drive.google.com/drive/folders/' + (id || ''), getName: () => '',
    getFoldersByName: () => iter([]), getFilesByName: () => iter([]), getFiles: () => iter([]), getFolders: () => iter([]),
    createFolder: () => dummyFolder(''), createFile: () => ({ getId: () => '', setName() { return this; }, setSharing() { return this; } }),
    setTrashed() { return this; }, isTrashed: () => false
  });
  const iter = arr => { let i = 0; return { hasNext: () => i < arr.length, next: () => arr[i++] }; };
  const G = {
    SpreadsheetApp: {
      openById: id => store.spreadsheet(id),
      create: () => { throw new Error('Tạo file Google Sheet mới chỉ làm được qua cầu nối Apps Script'); },
      flush() {}, getActiveSpreadsheet: () => null
    },
    PropertiesService: { getScriptProperties: () => props, getUserProperties: () => props, getDocumentProperties: () => props },
    CacheService,
    LockService: { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {}, hasLock: () => true }), getDocumentLock() { return this.getScriptLock(); }, getUserLock() { return this.getScriptLock(); } },
    Utilities: {
      formatDate(d, tz, fmt) {
        const t = new Date(d.getTime() + 7 * 3600e3);                       // Asia/Ho_Chi_Minh
        const p = n => String(n).padStart(2, '0');
        return fmt.replace('yyyy', t.getUTCFullYear()).replace('MM', p(t.getUTCMonth() + 1)).replace('dd', p(t.getUTCDate())).replace('HH', p(t.getUTCHours())).replace('mm', p(t.getUTCMinutes())).replace('ss', p(t.getUTCSeconds()));
      },
      getUuid: () => crypto.randomUUID(),
      sleep() {},
      computeDigest: (alg, s) => sha256Hex(s).match(/../g).map(h => { const b = parseInt(h, 16); return b > 127 ? b - 256 : b; }),
      DigestAlgorithm: { SHA_256: 'SHA_256' }, Charset: { UTF_8: 'UTF_8' },
      base64Encode: s => btoa(unescape(encodeURIComponent(s))), base64Decode: s => s,
      newBlob: (data, mime, name) => ({ data, mime, name, setName(n) { this.name = n; return this; }, getBytes: () => data })
    },
    ContentService: { createTextOutput: s => ({ s, setMimeType() { return this; } }), MimeType: { JSON: 'json', TEXT: 'text' } },
    DriveApp: {
      getFolderById: id => dummyFolder(id), getFileById: id => dummyFolder(id), getRootFolder: () => dummyFolder('root'), createFolder: () => dummyFolder(''),
      Access: { ANYONE_WITH_LINK: 1 }, Permission: { VIEW: 1 }
    },
    MailApp: { sendEmail: m => { side.mails.push(m); }, getRemainingDailyQuota: () => 100 },
    GmailApp: { sendEmail: (to, subject, body, opt) => { side.mails.push(Object.assign({ to, subject, body }, opt || {})); } },
    ScriptApp: { getProjectTriggers: () => [{ getHandlerFunction: () => 'donDepHangNgay' }], newTrigger: () => ({ timeBased() { return this; }, everyDays() { return this; }, atHour() { return this; }, create() {} }), deleteTrigger() {}, getOAuthToken: () => '' },
    Session: { getEffectiveUser: () => ({ getEmail: () => (props.getProperty('OWNER_EMAIL') || '') }), getActiveUser: () => ({ getEmail: () => (props.getProperty('OWNER_EMAIL') || '') }), getScriptTimeZone: () => 'Asia/Ho_Chi_Minh' },
    UrlFetchApp: { fetch: () => { throw new Error('UrlFetchApp không dùng trong Worker'); } },
    HtmlService: {},
    Logger: { log: (...a) => console.log(...a) },
    console
  };
  return { G, side, props };
}
