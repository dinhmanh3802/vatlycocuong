var __defProp = Object.defineProperty;
var __name = (target, value) => __defProp(target, "name", { value, configurable: true });

// src/app.generated.js
function createApp(G) {
  const { SpreadsheetApp, PropertiesService, CacheService: CacheService2, LockService, Utilities, ContentService, DriveApp, MailApp, GmailApp, ScriptApp, Session, UrlFetchApp, HtmlService, Logger } = G;
  const WORKER_MODE = !!G.WORKER_MODE;
  const console2 = G.console;
  const ROSTER_SHEET_ID = "11S9VfB_83K74Bv9m36QCF8FHTvqXg0s8A_tZWXRG48c";
  const DEFAULT_TEMP_PW = "Matkhau@123";
  const SESSION_TTL_HOURS = 12;
  const APP_VERSION = "7.0.0";
  var WORKER_URL = "";
  function isWorker() {
    try {
      return typeof WORKER_MODE !== "undefined" && !!WORKER_MODE;
    } catch (e) {
      return false;
    }
  }
  __name(isWorker, "isWorker");
  const TEN_HE_THONG = "V\u1EADt L\xFD C\xF4 L\xE2m C\u01B0\u1EDDng";
  function _props() {
    return PropertiesService.getScriptProperties();
  }
  __name(_props, "_props");
  var _SS_MEMO = {};
  var _SYS_ID = null;
  function getSystemSheetId() {
    if (_SYS_ID) return _SYS_ID;
    var id = _props().getProperty("SYSTEM_SHEET_ID");
    if (!id) throw new Error("Ch\u01B0a ch\u1EA1y setup(). H\xE3y ch\u1EA1y h\xE0m setup() m\u1ED9t l\u1EA7n trong tr\xECnh so\u1EA1n Apps Script.");
    _SYS_ID = id;
    return id;
  }
  __name(getSystemSheetId, "getSystemSheetId");
  function openSpreadsheetMemo(id) {
    if (!_SS_MEMO[id]) _SS_MEMO[id] = SpreadsheetApp.openById(id);
    return _SS_MEMO[id];
  }
  __name(openSpreadsheetMemo, "openSpreadsheetMemo");
  function getSystemSheet() {
    return openSpreadsheetMemo(getSystemSheetId());
  }
  __name(getSystemSheet, "getSystemSheet");
  function openRoster(rosterSheetId) {
    if (!rosterSheetId) throw new Error("Gi\xE1o vi\xEAn ch\u01B0a \u0111\u01B0\u1EE3c g\xE1n file danh s\xE1ch");
    return openSpreadsheetMemo(rosterSheetId);
  }
  __name(openRoster, "openRoster");
  function uuid() {
    return Utilities.getUuid();
  }
  __name(uuid, "uuid");
  function nowIso() {
    return (/* @__PURE__ */ new Date()).toISOString();
  }
  __name(nowIso, "nowIso");
  function sha256Hex2(s) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
    return bytes.map(function(b) {
      return ("0" + (b & 255).toString(16)).slice(-2);
    }).join("");
  }
  __name(sha256Hex2, "sha256Hex");
  function hashPassword(pw, salt) {
    var h = salt + "|" + pw;
    for (var i = 0; i < 1e3; i++) h = sha256Hex2(h);
    return h;
  }
  __name(hashPassword, "hashPassword");
  function randomSalt() {
    return uuid().replace(/-/g, "");
  }
  __name(randomSalt, "randomSalt");
  function jsonOut(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  }
  __name(jsonOut, "jsonOut");
  function idx(headerRow) {
    var m = {};
    headerRow.forEach(function(h, i) {
      m[String(h).trim()] = i;
    });
    return m;
  }
  __name(idx, "idx");
  function toBool(v) {
    if (v === true) return true;
    var s = String(v).trim().toLowerCase();
    return s === "true" || s === "1" || s === "x" || s === "c\xF3";
  }
  __name(toBool, "toBool");
  function normSbd(v) {
    if (v === null || v === void 0) return "";
    var s = String(v).trim();
    if (s === "") return "";
    if (/^\d+(\.0+)?$/.test(s)) s = String(parseInt(s, 10));
    s = s.replace(/\D/g, "");
    if (s === "") return "";
    while (s.length < 6) s = "0" + s;
    return s;
  }
  __name(normSbd, "normSbd");
  function normMa(v) {
    if (v === null || v === void 0 || String(v).trim() === "") return "";
    var n = parseInt(String(v).replace(/\D/g, ""), 10);
    if (isNaN(n)) return "";
    return ("00" + n).slice(-3);
  }
  __name(normMa, "normMa");
  function pad3(n) {
    return ("00" + parseInt(n, 10)).slice(-3);
  }
  __name(pad3, "pad3");
  function setTextCell(sheet, row, col1, val) {
    var rg = sheet.getRange(row, col1);
    rg.setNumberFormat("@");
    rg.setValue(val);
    sysTouch();
  }
  __name(setTextCell, "setTextCell");
  var _ROWS_MEMO = {};
  function sysTouch(tab) {
    if (tab) delete _ROWS_MEMO[tab];
    else _ROWS_MEMO = {};
  }
  __name(sysTouch, "sysTouch");
  function sysRows(tab) {
    if (_ROWS_MEMO[tab]) return _ROWS_MEMO[tab].slice();
    var sh = getSystemSheet().getSheetByName(tab);
    if (!sh) return [];
    var vals = sh.getDataRange().getValues();
    if (!vals.length) return [];
    var H = idx(vals[0]);
    var keys = Object.keys(H);
    var out = [];
    for (var i = 1; i < vals.length; i++) {
      var o = {};
      keys.forEach(function(k) {
        o[k] = vals[i][H[k]];
      });
      o._row = i + 1;
      out.push(o);
    }
    _ROWS_MEMO[tab] = out;
    return out.slice();
  }
  __name(sysRows, "sysRows");
  var CACHE_CHUNK = 3e4;
  function cachePutBig(key, obj, ttl) {
    var s = JSON.stringify(obj), c = CacheService2.getScriptCache();
    var n = Math.max(1, Math.ceil(s.length / CACHE_CHUNK)), m = {};
    for (var i = 0; i < n; i++) m[key + "_" + i] = s.substr(i * CACHE_CHUNK, CACHE_CHUNK);
    m[key + "_n"] = String(n);
    try {
      c.putAll(m, ttl || 600);
    } catch (e) {
    }
  }
  __name(cachePutBig, "cachePutBig");
  function cacheGetBig(key) {
    var c = CacheService2.getScriptCache(), n = Number(c.get(key + "_n"));
    if (!n) return null;
    var ks = [];
    for (var i = 0; i < n; i++) ks.push(key + "_" + i);
    var got = c.getAll(ks), s = "";
    for (var j = 0; j < n; j++) {
      if (got[ks[j]] === void 0 || got[ks[j]] === null) return null;
      s += got[ks[j]];
    }
    try {
      return JSON.parse(s);
    } catch (e) {
      return null;
    }
  }
  __name(cacheGetBig, "cacheGetBig");
  var _DVER = null;
  var _KHONG_DOI = false;
  function dataVer() {
    if (_DVER) return _DVER;
    var c = CacheService2.getScriptCache(), v = c.get("DVER");
    if (!v) {
      v = Date.now().toString(36);
      c.put("DVER", v, 21600);
    }
    return _DVER = v;
  }
  __name(dataVer, "dataVer");
  function dataTouch() {
    _DVER = Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
    CacheService2.getScriptCache().put("DVER", _DVER, 21600);
  }
  __name(dataTouch, "dataTouch");
  function cachedResult(key, ttl, fn, fresh) {
    var k = "cr_" + key + "_" + dataVer();
    if (!fresh) {
      var hit = cacheGetBig(k);
      if (hit) return hit;
    }
    var res = fn();
    try {
      cachePutBig(k, res, ttl);
    } catch (e) {
    }
    return res;
  }
  __name(cachedResult, "cachedResult");
  function cacheDel(key) {
    try {
      CacheService2.getScriptCache().remove(key + "_n");
    } catch (e) {
    }
  }
  __name(cacheDel, "cacheDel");
  function parseNhomList(v) {
    if (v === null || v === void 0 || v === "") return [];
    var raw = [], i;
    if (typeof v === "number") {
      var s = String(v);
      if (s.indexOf(".") !== -1) {
        s.split(".").forEach(function(part, k) {
          raw.push(k === 0 ? part : (part + "00").slice(0, 3));
        });
      } else {
        while (s.length % 3) s = "0" + s;
        for (i = 0; i < s.length; i += 3) raw.push(s.substr(i, 3));
      }
    } else {
      String(v).split(/[,;\s]+/).forEach(function(x) {
        if (x) raw.push(x);
      });
    }
    var seen = {}, out = [];
    raw.forEach(function(x) {
      var m = normMa(x);
      if (m && !seen[m]) {
        seen[m] = true;
        out.push(m);
      }
    });
    return out;
  }
  __name(parseNhomList, "parseNhomList");
  var SYS_TABS2 = {
    NguoiDung: ["username", "role", "fullName", "passHash", "salt", "mustChange", "active", "rosterSheetId", "rangeStart", "rangeEnd", "createdAt", "email", "maNhom", "quanLy"],
    Nhom: ["maNhom", "tenNhom", "sheetName", "teacher", "rosterSheetId", "total", "active"],
    Phien: ["token", "username", "role", "fullName", "expiry"],
    BaiKiemTra: ["maBai", "tenBai", "ngay", "teacher", "cacNhom", "cauHinh", "trangThai", "driveFolderId", "ngayCongBo", "createdAt", "luuTru", "soAnh", "dapAnFileId"],
    DapAn: ["maBai", "maDe", "dapAn"],
    KetQua: ["maBai", "sbd", "maDe", "tong", "p1", "p2", "p3", "vang", "traLoi", "capNhat", "daSua", "nguon"],
    NhatKySua: ["thoiGian", "nguoiSua", "maBai", "sbd", "truong", "giaTriCu", "giaTriMoi", "lyDo"],
    GopY: ["thoiGian", "noiDung", "lienHe", "daXuLy", "id", "nguon", "nguoiGui", "sbd"],
    NhatKyLoi: ["thoiGian", "action", "user", "message"],
    BuoiHoc: ["maBuoi", "maNhom", "ngay", "trangThai", "dd", "ghiChu", "nguoiGhi", "capNhat", "nopBoi", "nopLuc", "chotBoi", "chotLuc", "taoBoi", "taoLuc"],
    DiemDanhLog: ["thoiGian", "maBuoi", "sbd", "cu", "moi", "nguoi", "loai", "lyDo"],
    YeuCauSua: ["id", "thoiGian", "maBuoi", "maNhom", "sbd", "cu", "moi", "lyDo", "nguoiGui", "lienHe", "trangThai", "duyetBoi", "duyetLuc", "phanHoi"],
    ChoXuLy: ["id", "maBai", "sbd", "maDe", "tong", "p1", "p2", "p3", "traLoi", "lyDo", "ghiChu", "tenFile", "capNhat"],
    CaiDat: ["key", "value"]
  };
  function setup() {
    var props = _props();
    var id = props.getProperty("SYSTEM_SHEET_ID");
    var ss;
    if (id) {
      ss = SpreadsheetApp.openById(id);
    } else {
      ss = SpreadsheetApp.create("H\u1EC6 TH\u1ED0NG \u0110I\u1EC2M - D\u1EEE LI\u1EC6U (\u0111\u1EEBng ch\u1EC9nh tay)");
      id = ss.getId();
      props.setProperty("SYSTEM_SHEET_ID", id);
    }
    props.setProperty("ROSTER_SHEET_ID", ROSTER_SHEET_ID);
    Object.keys(SYS_TABS2).forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);
      if (sh.getLastRow() === 0) sh.appendRow(SYS_TABS2[name]);
    });
    var def = ss.getSheetByName("Sheet1");
    if (def) {
      try {
        ss.deleteSheet(def);
      } catch (e) {
      }
    }
    ensureSysColumns(ss);
    seedSetting(ss, "hienThiThuHang", "false");
    seedSetting(ss, "hienDiemTB", "false");
    seedSetting(ss, "soNgayGiuAnh", "30");
    seedSetting(ss, "tenHienThi", TEN_HE_THONG);
    Object.keys(CAI_DAT_MAC_DINH).forEach(function(k) {
      if (k !== "emailNhanThongBao") seedSetting(ss, k, CAI_DAT_MAC_DINH[k]);
    });
    seedSetting(ss, "emailNhanThongBao", Session.getEffectiveUser().getEmail() || "");
    seedUser(ss, "admin", "admin", "Qu\u1EA3n tr\u1ECB", "", "", "");
    seedUser(ss, "me", "teacher", "C\xF4 L\xE2m C\u01B0\u1EDDng", ROSTER_SHEET_ID, 0, 199);
    var me = findUserRow("me");
    if (me) {
      if (!me.data[me.H["rosterSheetId"]]) me.sheet.getRange(me.row, me.H["rosterSheetId"] + 1).setValue(ROSTER_SHEET_ID);
      if (me.data[me.H["rangeStart"]] === "" || me.data[me.H["rangeStart"]] === void 0) me.sheet.getRange(me.row, me.H["rangeStart"] + 1).setValue(0);
      if (me.data[me.H["rangeEnd"]] === "" || me.data[me.H["rangeEnd"]] === void 0) me.sheet.getRange(me.row, me.H["rangeEnd"] + 1).setValue(199);
    }
    var report = syncTeacher("me");
    Logger.log("XONG. Sheet h\u1EC7 th\u1ED1ng: " + ss.getUrl());
    Logger.log("\u0110\u0103ng nh\u1EADp admin:     admin / " + DEFAULT_TEMP_PW);
    Logger.log("\u0110\u0103ng nh\u1EADp gi\xE1o vi\xEAn: me / " + DEFAULT_TEMP_PW);
    Logger.log("\u0110\xE3 \u0111\u0103ng k\xFD " + report.registered.length + " nh\xF3m cho c\xF4 C\u01B0\u1EDDng (d\u1EA3i 000\u2013199).");
    if (report.conflicts.length) Logger.log("C\u1EA3nh b\xE1o: " + JSON.stringify(report.conflicts));
    return ss.getUrl();
  }
  __name(setup, "setup");
  function ensureSysColumns(ss) {
    Object.keys(SYS_TABS2).forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) return;
      var lastCol = Math.max(sh.getLastColumn(), 1);
      var header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(c) {
        return String(c).trim();
      });
      SYS_TABS2[name].forEach(function(col) {
        if (header.indexOf(col) === -1) {
          sh.getRange(1, sh.getLastColumn() + 1).setValue(col);
          header.push(col);
        }
      });
    });
  }
  __name(ensureSysColumns, "ensureSysColumns");
  function seedSetting(ss, key, value) {
    var sh = ss.getSheetByName("CaiDat");
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) if (String(rows[i][0]) === key) return;
    sh.appendRow([key, value]);
  }
  __name(seedSetting, "seedSetting");
  function seedUser(ss, username, role, fullName, rosterId, rangeStart, rangeEnd) {
    var sh = ss.getSheetByName("NguoiDung");
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) if (String(rows[i][0]) === username) return;
    var salt = randomSalt();
    var hash = hashPassword(DEFAULT_TEMP_PW, salt);
    sh.appendRow([
      username,
      role,
      fullName,
      hash,
      salt,
      true,
      true,
      rosterId || "",
      rangeStart === "" ? "" : rangeStart,
      rangeEnd === "" ? "" : rangeEnd,
      nowIso()
    ]);
  }
  __name(seedUser, "seedUser");
  function ensureRosterColumns(sheet) {
    var values = sheet.getDataRange().getValues();
    var headerRow = -1, headers = [];
    for (var r = 0; r < Math.min(values.length, 6); r++) {
      var up = values[r].map(function(c) {
        return String(c).trim().toUpperCase();
      });
      if (up.indexOf("SBD") !== -1) {
        headerRow = r;
        headers = up;
        break;
      }
    }
    if (headerRow === -1) return false;
    var nextCol = sheet.getLastColumn() + 1;
    ["TR\u1EA0NG TH\xC1I", "SBD C\u0168"].forEach(function(h) {
      if (headers.indexOf(h) === -1) {
        sheet.getRange(headerRow + 1, nextCol).setValue(h);
        nextCol++;
      }
    });
    return true;
  }
  __name(ensureRosterColumns, "ensureRosterColumns");
  function round2(x) {
    return Math.round(x * 100 + 1e-9) / 100;
  }
  __name(round2, "round2");
  function normP3(s) {
    return String(s == null ? "" : s).trim().replace(/\./g, ",").replace(/_+$/, "");
  }
  __name(normP3, "normP3");
  function normP2(s) {
    var t = String(s == null ? "" : s), out = "";
    for (var i = 0; i < 4; i++) {
      var c = (t.charAt(i) || "_").toUpperCase();
      out += c === "\u0110" || c === "S" ? c : "_";
    }
    return out;
  }
  __name(normP2, "normP2");
  function chamBai(cfg, tl, key) {
    tl = tl || {};
    key = key || {};
    var out = { p1: 0, p2: 0, p3: 0, tong: 0, thieuKey: false, dung: { p1: [], p2: [], p3: [] } };
    if (cfg.p1 && cfg.p1.on) {
      var n1 = Number(cfg.p1.soCau) || 0, per1 = n1 ? Number(cfg.p1.diem) / n1 : 0;
      for (var i = 0; i < n1; i++) {
        var k1 = key.p1 && key.p1[i] ? String(key.p1[i]).trim().toUpperCase() : "";
        if (!k1) {
          out.thieuKey = true;
          out.dung.p1.push(null);
          continue;
        }
        var a1 = tl.p1 && tl.p1[i] ? String(tl.p1[i]).trim().toUpperCase() : "";
        var ok1 = a1 === k1;
        out.dung.p1.push(ok1 ? 1 : 0);
        if (ok1) out.p1 += per1;
      }
    }
    if (cfg.p2 && cfg.p2.on) {
      var n2 = Number(cfg.p2.soCau) || 0;
      var pts = [0, Number(cfg.p2.d1) || 0, Number(cfg.p2.d2) || 0, Number(cfg.p2.d3) || 0, Number(cfg.p2.d4) || 0];
      for (var j = 0; j < n2; j++) {
        var k2 = key.p2 && key.p2[j] ? normP2(key.p2[j]) : "";
        if (!k2 || k2.indexOf("_") !== -1) {
          out.thieuKey = true;
          out.dung.p2.push(null);
          continue;
        }
        var a2 = normP2(tl.p2 && tl.p2[j]);
        var cnt = 0;
        for (var m = 0; m < 4; m++) if (a2.charAt(m) === k2.charAt(m)) cnt++;
        out.dung.p2.push(cnt);
        out.p2 += pts[cnt];
      }
    }
    if (cfg.p3 && cfg.p3.on) {
      var n3 = Number(cfg.p3.soCau) || 0, per3 = n3 ? Number(cfg.p3.diem) / n3 : 0;
      for (var q = 0; q < n3; q++) {
        var k3 = key.p3 ? normP3(key.p3[q]) : "";
        if (!k3) {
          out.thieuKey = true;
          out.dung.p3.push(null);
          continue;
        }
        var a3 = normP3(tl.p3 && tl.p3[q]);
        var ok3 = a3 !== "" && a3 === k3;
        out.dung.p3.push(ok3 ? 1 : 0);
        if (ok3) out.p3 += per3;
      }
    }
    out.p1 = round2(out.p1);
    out.p2 = round2(out.p2);
    out.p3 = round2(out.p3);
    out.tong = round2(out.p1 + out.p2 + out.p3);
    return out;
  }
  __name(chamBai, "chamBai");
  function getRootFolder() {
    var p = _props();
    var id = p.getProperty("DRIVE_ROOT_ID");
    if (id) {
      try {
        return DriveApp.getFolderById(id);
      } catch (e) {
      }
    }
    var f = DriveApp.createFolder("H\u1EC6 TH\u1ED0NG \u0110I\u1EC2M - L\u1EDBp h\u1ECDc th\xEAm");
    p.setProperty("DRIVE_ROOT_ID", f.getId());
    return f;
  }
  __name(getRootFolder, "getRootFolder");
  function getTeacherFolder(username) {
    var p = _props();
    var key = "DRIVEGV_" + username;
    var id = p.getProperty(key);
    if (id) {
      try {
        return DriveApp.getFolderById(id);
      } catch (e) {
      }
    }
    var root = getRootFolder();
    var name = "GV - " + username;
    var it = root.getFoldersByName(name);
    var f = it.hasNext() ? it.next() : root.createFolder(name);
    p.setProperty(key, f.getId());
    return f;
  }
  __name(getTeacherFolder, "getTeacherFolder");
  function safeName(s) {
    return String(s || "").replace(/[\\/:*?"<>|]/g, "-").trim().substring(0, 80) || "B\xE0i";
  }
  __name(safeName, "safeName");
  function createBaiFolder(username, tenBai, ngay) {
    var tf = getTeacherFolder(username);
    var name = (ngay ? ngay + " - " : "") + safeName(tenBai);
    return tf.createFolder(name);
  }
  __name(createBaiFolder, "createBaiFolder");
  function getAnhFolder(baiFolderId) {
    var folder = DriveApp.getFolderById(baiFolderId);
    var it = folder.getFoldersByName("anh");
    return it.hasNext() ? it.next() : folder.createFolder("anh");
  }
  __name(getAnhFolder, "getAnhFolder");
  function exportSpreadsheetAsXlsxBlob(ssId, filename) {
    var url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx";
    var res = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error("Kh\xF4ng xu\u1EA5t \u0111\u01B0\u1EE3c file xlsx (" + res.getResponseCode() + ")");
    return res.getBlob().setName(filename);
  }
  __name(exportSpreadsheetAsXlsxBlob, "exportSpreadsheetAsXlsxBlob");
  function saveBlobToFolder(folder, blob, name) {
    var it = folder.getFilesByName(name);
    while (it.hasNext()) it.next().setTrashed(true);
    return folder.createFile(blob.setName(name));
  }
  __name(saveBlobToFolder, "saveBlobToFolder");
  var LOGIN_MAX_FAIL = 5;
  var LOGIN_LOCK_MIN = 15;
  function _users() {
    return getSystemSheet().getSheetByName("NguoiDung");
  }
  __name(_users, "_users");
  function _phien() {
    return getSystemSheet().getSheetByName("Phien");
  }
  __name(_phien, "_phien");
  function findUserRow(username) {
    var sh = _users();
    var rows = sh.getDataRange().getValues();
    var H = idx(rows[0]);
    var want = String(username || "").trim().toLowerCase();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][H["username"]]).trim().toLowerCase() === want) return { row: i + 1, data: rows[i], H, sheet: sh };
    }
    return null;
  }
  __name(findUserRow, "findUserRow");
  function userField(u, k) {
    return u.H[k] === void 0 ? "" : u.data[u.H[k]];
  }
  __name(userField, "userField");
  function authError(msg) {
    var e = new Error(msg);
    e.code = "AUTH";
    return e;
  }
  __name(authError, "authError");
  function issuedAt(username) {
    var k = "ISS_" + String(username).toLowerCase(), c = CacheService2.getScriptCache(), v = c.get(k);
    if (v === null || v === void 0) {
      v = _props().getProperty(k) || "0";
      c.put(k, v, 21600);
    }
    return Number(v) || 0;
  }
  __name(issuedAt, "issuedAt");
  function setIssued(username, t) {
    var k = "ISS_" + String(username).toLowerCase();
    _props().setProperty(k, String(t));
    CacheService2.getScriptCache().put(k, String(t), 21600);
  }
  __name(setIssued, "setIssued");
  function revokeSessions(username) {
    var k = "REV_" + String(username).toLowerCase(), v = String(Math.max(Date.now(), issuedAt(username)));
    _props().setProperty(k, v);
    CacheService2.getScriptCache().put(k, v, 21600);
  }
  __name(revokeSessions, "revokeSessions");
  function revokedAt(username) {
    var k = "REV_" + String(username).toLowerCase(), c = CacheService2.getScriptCache(), v = c.get(k);
    if (v === null || v === void 0) {
      v = _props().getProperty(k) || "0";
      c.put(k, v, 21600);
    }
    return Number(v) || 0;
  }
  __name(revokedAt, "revokedAt");
  function apiLogin(p) {
    var username = String(p.username || "").trim();
    var pw = String(p.password || "");
    if (!username || !pw) throw new Error("Nh\u1EADp t\xE0i kho\u1EA3n v\xE0 m\u1EADt kh\u1EA9u");
    var cache = CacheService2.getScriptCache(), fk = "lf_" + username.toLowerCase();
    var fails = Number(cache.get(fk)) || 0;
    if (fails >= LOGIN_MAX_FAIL) throw new Error("\u0110\u0103ng nh\u1EADp sai qu\xE1 " + LOGIN_MAX_FAIL + " l\u1EA7n. Th\u1EED l\u1EA1i sau " + LOGIN_LOCK_MIN + ' ph\xFAt, ho\u1EB7c d\xF9ng "Qu\xEAn m\u1EADt kh\u1EA9u".');
    var u = findUserRow(username);
    var dungMk = u && hashPassword(pw, userField(u, "salt")) === userField(u, "passHash");
    if (!dungMk) {
      cache.put(fk, String(fails + 1), LOGIN_LOCK_MIN * 60);
      var con = LOGIN_MAX_FAIL - fails - 1;
      throw new Error("Sai t\xE0i kho\u1EA3n ho\u1EB7c m\u1EADt kh\u1EA9u" + (con > 0 && con <= 2 ? " (c\xF2n " + con + " l\u1EA7n th\u1EED)" : ""));
    }
    if (!toBool(userField(u, "active"))) throw new Error("T\xE0i kho\u1EA3n \u0111\xE3 b\u1ECB kh\xF3a. Li\xEAn h\u1EC7 qu\u1EA3n tr\u1ECB.");
    cache.remove(fk);
    var uname = String(userField(u, "username")), role = userField(u, "role"), fullName = userField(u, "fullName");
    var token = createSession(uname, role, fullName);
    return { token, user: { username: uname, role, fullName, mustChange: toBool(userField(u, "mustChange")) } };
  }
  __name(apiLogin, "apiLogin");
  function createSession(username, role, fullName) {
    var token = uuid() + uuid().replace(/-/g, "");
    var now = Math.max(Date.now(), revokedAt(username) + 1);
    setIssued(username, now);
    var exp = new Date(now + SESSION_TTL_HOURS * 3600 * 1e3).toISOString();
    var obj = { token, username, role, fullName, expiry: exp, iat: now };
    _phien().appendRow([token, username, role, fullName, exp]);
    CacheService2.getScriptCache().put("sess_" + token, JSON.stringify(obj), 6 * 3600);
    return token;
  }
  __name(createSession, "createSession");
  function getSession(token) {
    if (!token) return null;
    var cached = CacheService2.getScriptCache().get("sess_" + token);
    var s = null;
    if (cached) {
      s = JSON.parse(cached);
    } else {
      var rows = _phien().getDataRange().getValues();
      var H = idx(rows[0]);
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][H["token"]]) === token) {
          s = { token, username: rows[i][H["username"]], role: rows[i][H["role"]], fullName: rows[i][H["fullName"]], expiry: rows[i][H["expiry"]], iat: new Date(rows[i][H["expiry"]]).getTime() - SESSION_TTL_HOURS * 3600 * 1e3 };
          break;
        }
      }
    }
    if (!s) return null;
    if (new Date(s.expiry).getTime() < Date.now()) return null;
    var rev = revokedAt(s.username);
    if (rev && (s.iat || 0) <= rev) return null;
    return s;
  }
  __name(getSession, "getSession");
  function currentUser(p) {
    var s = getSession(p.token);
    if (!s) throw authError("Phi\xEAn \u0111\xE3 h\u1EBFt h\u1EA1n, h\xE3y \u0111\u0103ng nh\u1EADp l\u1EA1i");
    return { username: s.username, role: s.role, fullName: s.fullName };
  }
  __name(currentUser, "currentUser");
  function optionalUser(p) {
    try {
      return p.token ? currentUser(p) : null;
    } catch (e) {
      return null;
    }
  }
  __name(optionalUser, "optionalUser");
  function apiLogout(p) {
    if (p.token) {
      CacheService2.getScriptCache().remove("sess_" + p.token);
      var sh = _phien();
      var rows = sh.getDataRange().getValues();
      for (var i = rows.length - 1; i >= 1; i--) if (String(rows[i][0]) === p.token) sh.deleteRow(i + 1);
    }
    return { ok: true };
  }
  __name(apiLogout, "apiLogout");
  function setPassword(u, newPw, mustChange) {
    var salt = randomSalt();
    u.sheet.getRange(u.row, u.H["salt"] + 1).setValue(salt);
    u.sheet.getRange(u.row, u.H["passHash"] + 1).setValue(hashPassword(newPw, salt));
    u.sheet.getRange(u.row, u.H["mustChange"] + 1).setValue(!!mustChange);
  }
  __name(setPassword, "setPassword");
  function checkNewPassword(pw) {
    pw = String(pw || "");
    if (pw.length < 6) throw new Error("M\u1EADt kh\u1EA9u m\u1EDBi c\u1EA7n \xEDt nh\u1EA5t 6 k\xFD t\u1EF1");
    if (pw.length > 100) throw new Error("M\u1EADt kh\u1EA9u qu\xE1 d\xE0i");
    return pw;
  }
  __name(checkNewPassword, "checkNewPassword");
  function apiChangePassword(p) {
    var u0 = currentUser(p);
    var u = findUserRow(u0.username);
    if (!u) throw new Error("Kh\xF4ng t\xECm th\u1EA5y t\xE0i kho\u1EA3n");
    if (hashPassword(String(p.oldPassword || ""), userField(u, "salt")) !== userField(u, "passHash")) throw new Error("M\u1EADt kh\u1EA9u hi\u1EC7n t\u1EA1i kh\xF4ng \u0111\xFAng");
    var pw = checkNewPassword(p.newPassword);
    if (pw === String(p.oldPassword || "")) throw new Error("M\u1EADt kh\u1EA9u m\u1EDBi ph\u1EA3i kh\xE1c m\u1EADt kh\u1EA9u c\u0169");
    setPassword(u, pw, false);
    revokeSessions(u0.username);
    Utilities.sleep(5);
    var token = createSession(u0.username, u0.role, u0.fullName);
    return { ok: true, token };
  }
  __name(apiChangePassword, "apiChangePassword");
  function quenMatKhau(p) {
    var username = String(p.username || "").trim();
    if (!username) throw new Error("Nh\u1EADp t\xEAn t\xE0i kho\u1EA3n");
    var res = { ok: true, msg: "N\u1EBFu t\xE0i kho\u1EA3n c\xF3 email, m\xE3 \u0111\u1EB7t l\u1EA1i \u0111\xE3 \u0111\u01B0\u1EE3c g\u1EEDi t\u1EDBi email \u0111\xF3. M\xE3 c\xF3 hi\u1EC7u l\u1EF1c 15 ph\xFAt." };
    var cache = CacheService2.getScriptCache(), rk = "rs_" + username.toLowerCase();
    if (cache.get(rk + "_wait")) return res;
    cache.put(rk + "_wait", "1", 60);
    var u = findUserRow(username);
    if (!u || !toBool(userField(u, "active"))) return res;
    var email = String(userField(u, "email") || "").trim();
    if (!email) return res;
    var code = String(Math.floor(1e5 + Math.random() * 9e5));
    cache.put(rk, JSON.stringify({ h: sha256Hex2(code), n: 0 }), 900);
    try {
      MailApp.sendEmail({
        to: email,
        name: "H\u1EC7 th\u1ED1ng \u0111i\u1EC3m",
        subject: "[\u0110i\u1EC3m] M\xE3 \u0111\u1EB7t l\u1EA1i m\u1EADt kh\u1EA9u: " + code,
        body: 'M\xE3 \u0111\u1EB7t l\u1EA1i m\u1EADt kh\u1EA9u cho t\xE0i kho\u1EA3n "' + username + '": ' + code + "\n\nM\xE3 c\xF3 hi\u1EC7u l\u1EF1c 15 ph\xFAt. N\u1EBFu kh\xF4ng ph\u1EA3i b\u1EA1n y\xEAu c\u1EA7u, h\xE3y b\u1ECF qua email n\xE0y."
      });
    } catch (e) {
    }
    return res;
  }
  __name(quenMatKhau, "quenMatKhau");
  function datLaiMatKhau(p) {
    var username = String(p.username || "").trim();
    var cache = CacheService2.getScriptCache(), rk = "rs_" + username.toLowerCase();
    var st = cache.get(rk);
    if (!st) throw new Error('M\xE3 \u0111\xE3 h\u1EBFt h\u1EA1n ho\u1EB7c ch\u01B0a \u0111\u01B0\u1EE3c g\u1EEDi. B\u1EA5m "G\u1EEDi m\xE3" l\u1EA1i.');
    st = JSON.parse(st);
    if (st.n >= 5) {
      cache.remove(rk);
      throw new Error("Nh\u1EADp sai m\xE3 qu\xE1 nhi\u1EC1u l\u1EA7n. H\xE3y g\u1EEDi m\xE3 m\u1EDBi.");
    }
    if (sha256Hex2(String(p.code || "").trim()) !== st.h) {
      st.n++;
      cache.put(rk, JSON.stringify(st), 900);
      throw new Error("M\xE3 kh\xF4ng \u0111\xFAng");
    }
    var pw = checkNewPassword(p.newPassword);
    var u = findUserRow(username);
    if (!u) throw new Error("Kh\xF4ng t\xECm th\u1EA5y t\xE0i kho\u1EA3n");
    setPassword(u, pw, false);
    cache.remove(rk);
    cache.remove("lf_" + username.toLowerCase());
    revokeSessions(username);
    return { ok: true };
  }
  __name(datLaiMatKhau, "datLaiMatKhau");
  function rangesOverlap(aS, aE, bS, bE) {
    return aS <= bE && bS <= aE;
  }
  __name(rangesOverlap, "rangesOverlap");
  function checkRangeFree(rangeStart, rangeEnd, exceptUsername) {
    if (rangeStart === "" || rangeStart === void 0 || rangeStart === null || rangeEnd === "" || rangeEnd === void 0 || rangeEnd === null) return;
    var s = parseInt(rangeStart, 10), e = parseInt(rangeEnd, 10);
    if (isNaN(s) || isNaN(e) || s < 0 || e > 999 || s > e) throw new Error("D\u1EA3i m\xE3 kh\xF4ng h\u1EE3p l\u1EC7 (0\u2013999, \u0111\u1EA7u \u2264 cu\u1ED1i)");
    sysRows("NguoiDung").forEach(function(r) {
      if (r.role !== "teacher") return;
      if (exceptUsername && String(r.username).toLowerCase() === String(exceptUsername).toLowerCase()) return;
      if (r.rangeStart === "" || r.rangeStart === void 0) return;
      if (rangesOverlap(s, e, parseInt(r.rangeStart, 10), parseInt(r.rangeEnd, 10)))
        throw new Error("D\u1EA3i m\xE3 tr\xF9ng v\u1EDBi gi\xE1o vi\xEAn " + r.username + " (" + pad3(r.rangeStart) + "\u2013" + pad3(r.rangeEnd) + ")");
    });
  }
  __name(checkRangeFree, "checkRangeFree");
  function listUsers() {
    sysTouch("NguoiDung");
    return {
      users: sysRows("NguoiDung").map(function(r) {
        return {
          username: r.username,
          role: r.role,
          fullName: r.fullName,
          email: String(r.email || ""),
          active: toBool(r.active),
          mustChange: toBool(r.mustChange),
          rosterSheetId: r.rosterSheetId,
          rangeStart: r.rangeStart,
          rangeEnd: r.rangeEnd
        };
      })
    };
  }
  __name(listUsers, "listUsers");
  function listTeachers() {
    var users = sysRows("NguoiDung").filter(function(r) {
      return r.role === "teacher";
    });
    var nhom = sysRows("Nhom");
    return {
      teachers: users.map(function(u) {
        return {
          username: u.username,
          fullName: u.fullName,
          active: toBool(u.active),
          rosterSheetId: u.rosterSheetId,
          rangeStart: u.rangeStart,
          rangeEnd: u.rangeEnd,
          groupCount: nhom.filter(function(n) {
            return n.teacher === u.username;
          }).length
        };
      })
    };
  }
  __name(listTeachers, "listTeachers");
  function checkEmail(email) {
    email = String(email || "").trim();
    if (email && !/^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(email)) throw new Error("Email kh\xF4ng h\u1EE3p l\u1EC7");
    return email;
  }
  __name(checkEmail, "checkEmail");
  function createUser(p) {
    var username = String(p.username || "").trim();
    if (!/^[A-Za-z0-9._-]{2,40}$/.test(username)) throw new Error("T\xEAn t\xE0i kho\u1EA3n 2\u201340 k\xFD t\u1EF1, ch\u1EC9 g\u1ED3m ch\u1EEF kh\xF4ng d\u1EA5u, s\u1ED1, d\u1EA5u ch\u1EA5m, g\u1EA1ch");
    if (findUserRow(username)) throw new Error("T\xE0i kho\u1EA3n \u0111\xE3 t\u1ED3n t\u1EA1i");
    var role = p.role === "admin" ? "admin" : "teacher";
    var rangeStart = "", rangeEnd = "";
    if (role === "teacher") {
      rangeStart = p.rangeStart === void 0 || p.rangeStart === "" ? "" : parseInt(p.rangeStart, 10);
      rangeEnd = p.rangeEnd === void 0 || p.rangeEnd === "" ? "" : parseInt(p.rangeEnd, 10);
      checkRangeFree(rangeStart, rangeEnd, null);
    }
    var email = checkEmail(p.email);
    var sh = _users(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), nCol = sh.getLastColumn();
    var row = [];
    for (var i = 0; i < nCol; i++) row.push("");
    var salt = randomSalt();
    var set = /* @__PURE__ */ __name(function(k, v) {
      if (H[k] !== void 0) row[H[k]] = v;
    }, "set");
    set("username", username);
    set("role", role);
    set("fullName", String(p.fullName || "").trim());
    set("passHash", hashPassword(DEFAULT_TEMP_PW, salt));
    set("salt", salt);
    set("mustChange", true);
    set("active", true);
    set("rosterSheetId", String(p.rosterSheetId || "").trim());
    set("rangeStart", rangeStart);
    set("rangeEnd", rangeEnd);
    set("createdAt", nowIso());
    set("email", email);
    sh.appendRow(row);
    sysTouch();
    return { ok: true, tempPassword: DEFAULT_TEMP_PW };
  }
  __name(createUser, "createUser");
  function setUserInfo(p) {
    var u = findUserRow(p.username);
    if (!u) throw new Error("Kh\xF4ng t\xECm th\u1EA5y t\xE0i kho\u1EA3n");
    var isTeacher = userField(u, "role") === "teacher";
    if (isTeacher && p.rangeStart !== void 0 && p.rangeEnd !== void 0) checkRangeFree(p.rangeStart, p.rangeEnd, p.username);
    var put = /* @__PURE__ */ __name(function(k, v) {
      if (u.H[k] !== void 0) u.sheet.getRange(u.row, u.H[k] + 1).setValue(v);
    }, "put");
    if (p.fullName !== void 0) put("fullName", String(p.fullName || "").trim());
    if (p.email !== void 0) {
      if (u.H["email"] !== void 0) setTextCell(u.sheet, u.row, u.H["email"] + 1, checkEmail(p.email));
    }
    if (isTeacher && p.rosterSheetId !== void 0) put("rosterSheetId", String(p.rosterSheetId || "").trim());
    if (isTeacher && p.rangeStart !== void 0) put("rangeStart", p.rangeStart === "" ? "" : parseInt(p.rangeStart, 10));
    if (isTeacher && p.rangeEnd !== void 0) put("rangeEnd", p.rangeEnd === "" ? "" : parseInt(p.rangeEnd, 10));
    sysTouch();
    return { ok: true };
  }
  __name(setUserInfo, "setUserInfo");
  function setTeacherRoster(p) {
    return setUserInfo(p);
  }
  __name(setTeacherRoster, "setTeacherRoster");
  function setUserActive(p, me) {
    var u = findUserRow(p.username);
    if (!u) throw new Error("Kh\xF4ng t\xECm th\u1EA5y t\xE0i kho\u1EA3n");
    if (me && String(me.username).toLowerCase() === String(p.username).toLowerCase() && !p.active) throw new Error("Kh\xF4ng th\u1EC3 t\u1EF1 kh\xF3a t\xE0i kho\u1EA3n \u0111ang d\xF9ng");
    u.sheet.getRange(u.row, u.H["active"] + 1).setValue(!!p.active);
    if (!p.active) revokeSessions(userField(u, "username"));
    sysTouch();
    return { ok: true };
  }
  __name(setUserActive, "setUserActive");
  function adminResetPassword(p) {
    var u = findUserRow(p.username);
    if (!u) throw new Error("Kh\xF4ng t\xECm th\u1EA5y t\xE0i kho\u1EA3n");
    setPassword(u, DEFAULT_TEMP_PW, true);
    revokeSessions(userField(u, "username"));
    CacheService2.getScriptCache().remove("lf_" + String(p.username).toLowerCase());
    return { ok: true, tempPassword: DEFAULT_TEMP_PW };
  }
  __name(adminResetPassword, "adminResetPassword");
  function readGroupFromSheet(sheet) {
    var values = sheet.getDataRange().getValues();
    var headerRow = -1, cols = {};
    for (var r = 0; r < Math.min(values.length, 6); r++) {
      var up = values[r].map(function(c) {
        return String(c).trim().toUpperCase();
      });
      if (up.indexOf("SBD") !== -1) {
        headerRow = r;
        up.forEach(function(h, i2) {
          if (h) cols[h] = i2;
        });
        break;
      }
    }
    if (headerRow === -1) return null;
    var students = [], maNhom = "";
    for (var i = headerRow + 1; i < values.length; i++) {
      var raw = values[i];
      var sbd = normSbd(raw[cols["SBD"]]);
      var ho = cols["H\u1ECC"] !== void 0 ? String(raw[cols["H\u1ECC"]] || "").trim() : "";
      var ten = cols["T\xCAN"] !== void 0 ? String(raw[cols["T\xCAN"]] || "").trim() : "";
      if (!sbd && !ho && !ten) continue;
      if (sbd && !maNhom) maNhom = sbd.substring(0, 3);
      students.push({
        sbd,
        ho,
        ten,
        lop: cols["L\u1EDAP"] !== void 0 ? String(raw[cols["L\u1EDAP"]] || "").trim() : "",
        trangThai: (cols["TR\u1EA0NG TH\xC1I"] !== void 0 ? String(raw[cols["TR\u1EA0NG TH\xC1I"]] || "").trim() : "") || "\u0110ang h\u1ECDc",
        sbdCu: cols["SBD C\u0168"] !== void 0 ? normSbd(raw[cols["SBD C\u0168"]]) : "",
        rowIndex: i + 1
      });
    }
    return { sheetName: sheet.getName(), tenNhom: sheet.getName().trim(), maNhom, headerRow: headerRow + 1, cols, students };
  }
  __name(readGroupFromSheet, "readGroupFromSheet");
  function rosterSnapshot(fileId, fresh) {
    var key = "ros_" + fileId;
    if (!fresh) {
      var hit = cacheGetBig(key);
      if (hit) return hit;
    }
    var tabs = [];
    openRoster(fileId).getSheets().forEach(function(sh) {
      var g = readGroupFromSheet(sh);
      if (g) tabs.push(g);
    });
    cachePutBig(key, tabs, 600);
    return tabs;
  }
  __name(rosterSnapshot, "rosterSnapshot");
  function rosterTouch(fileId) {
    if (fileId) cacheDel("ros_" + fileId);
    try {
      pubTouch();
    } catch (e) {
    }
  }
  __name(rosterTouch, "rosterTouch");
  function findTab(tabs, sheetName, maNhom) {
    var byName = tabs.filter(function(t) {
      return String(t.sheetName) === String(sheetName);
    })[0];
    return byName || tabs.filter(function(t) {
      return t.maNhom === maNhom;
    })[0] || null;
  }
  __name(findTab, "findTab");
  function getNhomMap() {
    var m = {};
    sysRows("Nhom").forEach(function(r) {
      var ma = normMa(r.maNhom);
      if (ma) m[ma] = { tenNhom: r.tenNhom, teacher: r.teacher, rosterSheetId: r.rosterSheetId, sheetName: r.sheetName, _row: r._row };
    });
    return m;
  }
  __name(getNhomMap, "getNhomMap");
  function canAccessGroup(u, maNhom) {
    if (u.role === "admin") return true;
    var info = getNhomMap()[normMa(maNhom)];
    return info && info.teacher === u.username;
  }
  __name(canAccessGroup, "canAccessGroup");
  function requireGroupAccess(u, maNhom) {
    if (!canAccessGroup(u, maNhom)) throw new Error("Kh\xF4ng c\xF3 quy\u1EC1n v\u1EDBi nh\xF3m " + maNhom);
  }
  __name(requireGroupAccess, "requireGroupAccess");
  function getGroupByMaNhom(maNhom) {
    maNhom = normMa(maNhom);
    var reg = getNhomMap()[maNhom];
    if (!reg || !reg.rosterSheetId) return null;
    var file = openRoster(reg.rosterSheetId);
    var sh = file.getSheetByName(reg.sheetName);
    if (!sh) {
      var sheets = file.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        var g0 = readGroupFromSheet(sheets[i]);
        if (g0 && g0.maNhom === maNhom) {
          sh = sheets[i];
          break;
        }
      }
    }
    if (!sh) return null;
    return { sheet: sh, group: readGroupFromSheet(sh), reg };
  }
  __name(getGroupByMaNhom, "getGroupByMaNhom");
  function syncTeacher(username) {
    var user = findUserRow(username);
    if (!user) throw new Error("Kh\xF4ng t\xECm th\u1EA5y t\xE0i kho\u1EA3n " + username);
    var rosterId = user.data[user.H["rosterSheetId"]];
    if (!rosterId) throw new Error("Gi\xE1o vi\xEAn " + username + " ch\u01B0a g\xE1n file danh s\xE1ch");
    var rs = user.data[user.H["rangeStart"]] === "" ? 0 : parseInt(user.data[user.H["rangeStart"]], 10);
    var re = user.data[user.H["rangeEnd"]] === "" ? 999 : parseInt(user.data[user.H["rangeEnd"]], 10);
    var ss = getSystemSheet();
    var nhomSheet = ss.getSheetByName("Nhom");
    var Hn = idx(nhomSheet.getRange(1, 1, 1, nhomSheet.getLastColumn()).getValues()[0]);
    var reg = sysRows("Nhom");
    var registered = [], conflicts = [];
    var file = openRoster(rosterId);
    file.getSheets().forEach(function(sh) {
      ensureRosterColumns(sh);
      var g = readGroupFromSheet(sh);
      if (!g || !g.maNhom) return;
      var pfx = parseInt(g.maNhom, 10);
      if (pfx < rs || pfx > re) {
        conflicts.push({ maNhom: g.maNhom, tenNhom: g.tenNhom, lyDo: "ngo\xE0i d\u1EA3i " + pad3(rs) + "\u2013" + pad3(re) });
        return;
      }
      var otherOwner = reg.filter(function(r) {
        return normMa(r.maNhom) === g.maNhom && r.teacher !== username;
      })[0];
      if (otherOwner) {
        conflicts.push({ maNhom: g.maNhom, tenNhom: g.tenNhom, lyDo: "tr\xF9ng m\xE3 v\u1EDBi " + otherOwner.teacher });
        return;
      }
      var total = g.students.length;
      var active = g.students.filter(function(s) {
        return s.trangThai === "\u0110ang h\u1ECDc";
      }).length;
      var existing = reg.filter(function(r) {
        return normMa(r.maNhom) === g.maNhom && r.teacher === username;
      })[0];
      if (existing) {
        nhomSheet.getRange(existing._row, Hn["tenNhom"] + 1).setValue(g.tenNhom);
        nhomSheet.getRange(existing._row, Hn["sheetName"] + 1).setValue(g.sheetName);
        nhomSheet.getRange(existing._row, Hn["rosterSheetId"] + 1).setValue(rosterId);
        nhomSheet.getRange(existing._row, Hn["total"] + 1).setValue(total);
        nhomSheet.getRange(existing._row, Hn["active"] + 1).setValue(active);
      } else {
        nhomSheet.appendRow([g.maNhom, g.tenNhom, g.sheetName, username, rosterId, total, active]);
        reg.push({ maNhom: g.maNhom, teacher: username, _row: nhomSheet.getLastRow() });
      }
      registered.push({ maNhom: g.maNhom, tenNhom: g.tenNhom });
    });
    rosterTouch(rosterId);
    sysTouch();
    return { registered, conflicts };
  }
  __name(syncTeacher, "syncTeacher");
  function apiSyncGroups(p, u) {
    var target = u.role === "admin" && p.username ? p.username : u.username;
    if (u.role !== "admin" && target !== u.username) throw new Error("Kh\xF4ng \u0111\u1EE7 quy\u1EC1n");
    return syncTeacher(target);
  }
  __name(apiSyncGroups, "apiSyncGroups");
  function listGroups(u, fresh) {
    var reg = sysRows("Nhom").filter(function(r) {
      return (u.role === "admin" || r.teacher === u.username) && normMa(r.maNhom);
    });
    var byFile = {};
    reg.forEach(function(r) {
      (byFile[r.rosterSheetId] = byFile[r.rosterSheetId] || []).push(r);
    });
    var out = [];
    Object.keys(byFile).forEach(function(fileId) {
      var tabs = [];
      try {
        tabs = rosterSnapshot(fileId, !!fresh);
      } catch (e) {
      }
      byFile[fileId].forEach(function(r) {
        var ma = normMa(r.maNhom), g = findTab(tabs, r.sheetName, ma);
        out.push({
          maNhom: ma,
          tenNhom: g ? g.tenNhom : String(r.tenNhom),
          teacher: r.teacher,
          total: g ? g.students.length : Number(r.total) || 0,
          active: g ? g.students.filter(function(s) {
            return s.trangThai === "\u0110ang h\u1ECDc";
          }).length : Number(r.active) || 0
        });
      });
    });
    out.sort(function(a, b) {
      return a.maNhom < b.maNhom ? -1 : 1;
    });
    return { groups: out };
  }
  __name(listGroups, "listGroups");
  function listStudents(p, u) {
    requireGroupAccess(u, p.maNhom);
    var ma = normMa(p.maNhom), reg = getNhomMap()[ma];
    if (!reg) throw new Error("Kh\xF4ng t\xECm th\u1EA5y nh\xF3m " + p.maNhom);
    var g = findTab(rosterSnapshot(reg.rosterSheetId, !!p.fresh), reg.sheetName, ma);
    if (!g) throw new Error("Kh\xF4ng t\xECm th\u1EA5y tab c\u1EE7a nh\xF3m " + ma + " trong file danh s\xE1ch");
    return { maNhom: ma, tenNhom: g.tenNhom, students: g.students };
  }
  __name(listStudents, "listStudents");
  function nextSbd(g) {
    var max = 0;
    g.students.forEach(function(s) {
      if (s.sbd && s.sbd.substring(0, 3) === g.maNhom) {
        var n = parseInt(s.sbd.substring(3), 10);
        if (n > max) max = n;
      }
    });
    return g.maNhom + ("00" + (max + 1)).slice(-3);
  }
  __name(nextSbd, "nextSbd");
  function upsertStudent(p, u) {
    requireGroupAccess(u, p.maNhom);
    var lock = LockService.getScriptLock();
    lock.waitLock(2e4);
    try {
      var found = getGroupByMaNhom(p.maNhom);
      if (!found) throw new Error("Kh\xF4ng t\xECm th\u1EA5y nh\xF3m " + p.maNhom);
      ensureRosterColumns(found.sheet);
      found = getGroupByMaNhom(p.maNhom);
      var sh = found.sheet, g = found.group, cols = g.cols;
      var ho = (p.ho || "").trim(), ten = (p.ten || "").trim();
      if (!ho && !ten) throw new Error("Ph\u1EA3i nh\u1EADp h\u1ECD t\xEAn");
      var sbd = normSbd(p.sbd);
      var targetRow = null;
      if (sbd) {
        var ex = g.students.filter(function(s) {
          return s.sbd === sbd;
        })[0];
        if (ex) targetRow = ex.rowIndex;
      }
      var isNew = !targetRow;
      if (isNew) {
        if (sbd && g.students.filter(function(s) {
          return s.sbd === sbd;
        }).length) throw new Error("SBD \u0111\xE3 t\u1ED3n t\u1EA1i: " + sbd);
        if (!sbd) sbd = nextSbd(g);
        targetRow = (g.students.length ? g.students[g.students.length - 1].rowIndex : g.headerRow) + 1;
        if (cols["STT"] !== void 0) sh.getRange(targetRow, cols["STT"] + 1).setValue(g.students.length + 1);
        if (cols["TR\u1EA0NG TH\xC1I"] !== void 0) sh.getRange(targetRow, cols["TR\u1EA0NG TH\xC1I"] + 1).setValue("\u0110ang h\u1ECDc");
      }
      if (cols["H\u1ECC"] !== void 0) sh.getRange(targetRow, cols["H\u1ECC"] + 1).setValue(ho);
      if (cols["T\xCAN"] !== void 0) sh.getRange(targetRow, cols["T\xCAN"] + 1).setValue(ten);
      if (p.lop !== void 0 && cols["L\u1EDAP"] !== void 0) sh.getRange(targetRow, cols["L\u1EDAP"] + 1).setValue((p.lop || "").trim());
      if (cols["SBD"] !== void 0) setTextCell(sh, targetRow, cols["SBD"] + 1, sbd);
      rosterTouch(found.reg.rosterSheetId);
      return { ok: true, sbd, isNew };
    } finally {
      lock.releaseLock();
    }
  }
  __name(upsertStudent, "upsertStudent");
  function setStudentStatus(p, u) {
    requireGroupAccess(u, p.maNhom);
    var lock = LockService.getScriptLock();
    lock.waitLock(2e4);
    try {
      var found = getGroupByMaNhom(p.maNhom);
      if (!found) throw new Error("Kh\xF4ng t\xECm th\u1EA5y nh\xF3m " + p.maNhom);
      var sh = found.sheet, g = found.group;
      if (g.cols["TR\u1EA0NG TH\xC1I"] === void 0) {
        ensureRosterColumns(sh);
        found = getGroupByMaNhom(p.maNhom);
        sh = found.sheet;
        g = found.group;
      }
      var st = g.students.filter(function(s) {
        return s.sbd === normSbd(p.sbd);
      })[0];
      if (!st) throw new Error("Kh\xF4ng t\xECm th\u1EA5y SBD " + p.sbd);
      sh.getRange(st.rowIndex, g.cols["TR\u1EA0NG TH\xC1I"] + 1).setValue(p.trangThai || "\u0110ang h\u1ECDc");
      rosterTouch(found.reg.rosterSheetId);
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  }
  __name(setStudentStatus, "setStudentStatus");
  function transferStudent(p, u) {
    requireGroupAccess(u, p.fromMaNhom);
    requireGroupAccess(u, p.toMaNhom);
    var lock = LockService.getScriptLock();
    lock.waitLock(2e4);
    try {
      var from = getGroupByMaNhom(p.fromMaNhom);
      var to = getGroupByMaNhom(p.toMaNhom);
      if (!from || !to) throw new Error("Kh\xF4ng t\xECm th\u1EA5y nh\xF3m");
      ensureRosterColumns(from.sheet);
      ensureRosterColumns(to.sheet);
      from = getGroupByMaNhom(p.fromMaNhom);
      to = getGroupByMaNhom(p.toMaNhom);
      var oldSbd = normSbd(p.sbd);
      var st = from.group.students.filter(function(s) {
        return s.sbd === oldSbd;
      })[0];
      if (!st) throw new Error("Kh\xF4ng t\xECm th\u1EA5y SBD " + p.sbd + " trong nh\xF3m " + p.fromMaNhom);
      var newSbd = nextSbd(to.group);
      var tsh = to.sheet, tc = to.group.cols;
      var newRow = (to.group.students.length ? to.group.students[to.group.students.length - 1].rowIndex : to.group.headerRow) + 1;
      if (tc["STT"] !== void 0) tsh.getRange(newRow, tc["STT"] + 1).setValue(to.group.students.length + 1);
      if (tc["H\u1ECC"] !== void 0) tsh.getRange(newRow, tc["H\u1ECC"] + 1).setValue(st.ho);
      if (tc["T\xCAN"] !== void 0) tsh.getRange(newRow, tc["T\xCAN"] + 1).setValue(st.ten);
      if (tc["L\u1EDAP"] !== void 0) tsh.getRange(newRow, tc["L\u1EDAP"] + 1).setValue(st.lop);
      if (tc["SBD"] !== void 0) setTextCell(tsh, newRow, tc["SBD"] + 1, newSbd);
      if (tc["TR\u1EA0NG TH\xC1I"] !== void 0) tsh.getRange(newRow, tc["TR\u1EA0NG TH\xC1I"] + 1).setValue("\u0110ang h\u1ECDc");
      if (tc["SBD C\u0168"] !== void 0) setTextCell(tsh, newRow, tc["SBD C\u0168"] + 1, oldSbd);
      var fc = from.group.cols;
      if (fc["TR\u1EA0NG TH\xC1I"] !== void 0) from.sheet.getRange(st.rowIndex, fc["TR\u1EA0NG TH\xC1I"] + 1).setValue("\u0110\xE3 chuy\u1EC3n");
      rosterTouch(from.reg.rosterSheetId);
      rosterTouch(to.reg.rosterSheetId);
      return { ok: true, newSbd };
    } finally {
      lock.releaseLock();
    }
  }
  __name(transferStudent, "transferStudent");
  function validateRoster(u) {
    var issues = [], allSbd = {};
    var reg = sysRows("Nhom").filter(function(r) {
      return (u.role === "admin" || r.teacher === u.username) && normMa(r.maNhom);
    });
    var byFile = {};
    reg.forEach(function(r) {
      (byFile[r.rosterSheetId] = byFile[r.rosterSheetId] || []).push(r);
    });
    Object.keys(byFile).forEach(function(fileId) {
      var tabs;
      try {
        tabs = rosterSnapshot(fileId, true);
      } catch (e) {
        issues.push({ group: "(file)", type: "Kh\xF4ng m\u1EDF \u0111\u01B0\u1EE3c file", detail: String(fileId) });
        return;
      }
      byFile[fileId].forEach(function(r) {
        var g = findTab(tabs, r.sheetName, normMa(r.maNhom));
        if (!g) {
          issues.push({ group: String(r.tenNhom), type: "Thi\u1EBFu tab", detail: String(r.sheetName) });
          return;
        }
        var seen = {};
        g.students.forEach(function(s) {
          if (s.sbd) {
            (allSbd[s.sbd] = allSbd[s.sbd] || []).push(g.tenNhom);
            if (seen[s.sbd]) issues.push({ group: g.tenNhom, type: "Tr\xF9ng SBD trong nh\xF3m", detail: s.sbd });
            seen[s.sbd] = true;
            if (s.sbd.substring(0, 3) !== normMa(r.maNhom)) issues.push({ group: g.tenNhom, type: "SBD sai m\xE3 nh\xF3m", detail: s.sbd + " (nh\xF3m " + normMa(r.maNhom) + ")" });
            if (!s.ho && !s.ten) issues.push({ group: g.tenNhom, type: "C\xF3 SBD nh\u01B0ng thi\u1EBFu t\xEAn", detail: s.sbd });
          } else if (s.ho || s.ten) {
            issues.push({ group: g.tenNhom, type: "C\xF3 t\xEAn nh\u01B0ng thi\u1EBFu SBD", detail: (s.ho + " " + s.ten).trim() });
          }
        });
      });
    });
    Object.keys(allSbd).forEach(function(sbd) {
      if (allSbd[sbd].length > 1) issues.push({ group: allSbd[sbd].join(", "), type: "SBD tr\xF9ng nhi\u1EC1u nh\xF3m", detail: sbd });
    });
    return { ok: issues.length === 0, issues };
  }
  __name(validateRoster, "validateRoster");
  function _bai() {
    return getSystemSheet().getSheetByName("BaiKiemTra");
  }
  __name(_bai, "_bai");
  function _dapan() {
    return getSystemSheet().getSheetByName("DapAn");
  }
  __name(_dapan, "_dapan");
  function normMaDe(x) {
    var s = String(x == null ? "" : x).trim();
    if (s === "") return "";
    if (/^\d+$/.test(s)) return ("00" + parseInt(s, 10)).slice(-3);
    return s;
  }
  __name(normMaDe, "normMaDe");
  function ngayStr(v) {
    if (Object.prototype.toString.call(v) === "[object Date]") return Utilities.formatDate(v, "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
    return v === null || v === void 0 ? "" : String(v);
  }
  __name(ngayStr, "ngayStr");
  function findBai(maBai) {
    return sysRows("BaiKiemTra").filter(function(r) {
      return String(r.maBai) === String(maBai);
    })[0] || null;
  }
  __name(findBai, "findBai");
  function requireBaiAccess(u, baiRow) {
    if (!baiRow) throw new Error("Kh\xF4ng t\xECm th\u1EA5y b\xE0i");
    if (u.role !== "admin" && baiRow.teacher !== u.username) throw new Error("Kh\xF4ng c\xF3 quy\u1EC1n v\u1EDBi b\xE0i n\xE0y");
  }
  __name(requireBaiAccess, "requireBaiAccess");
  function validateCauHinh(cfg) {
    var total = 0;
    if (cfg.p1 && cfg.p1.on) {
      if (!(cfg.p1.soCau > 0)) throw new Error("Ph\u1EA7n I: s\u1ED1 c\xE2u ph\u1EA3i > 0");
      total += Number(cfg.p1.diem) || 0;
    }
    if (cfg.p2 && cfg.p2.on) {
      if (!(cfg.p2.soCau > 0)) throw new Error("Ph\u1EA7n II: s\u1ED1 c\xE2u ph\u1EA3i > 0");
      total += (Number(cfg.p2.soCau) || 0) * (Number(cfg.p2.d4) || 0);
    }
    if (cfg.p3 && cfg.p3.on) {
      if (!(cfg.p3.soCau > 0)) throw new Error("Ph\u1EA7n III: s\u1ED1 c\xE2u ph\u1EA3i > 0");
      total += Number(cfg.p3.diem) || 0;
    }
    if (Math.abs(total - 10) > 0.01) throw new Error("T\u1ED5ng \u0111i\u1EC3m t\u1ED1i \u0111a ph\u1EA3i b\u1EB1ng 10 (\u0111ang " + total.toFixed(2) + ")");
  }
  __name(validateCauHinh, "validateCauHinh");
  function createBai(p, u) {
    var cfg = p.cauHinh;
    if (typeof cfg === "string") cfg = JSON.parse(cfg);
    if (cfg && cfg.loai === "ngoai") cfg = chuanCfgNgoai(cfg);
    else {
      validateCauHinh(cfg);
      cfg.maDeList = (cfg.maDeList || []).map(normMaDe).filter(function(x) {
        return x;
      });
      if (!cfg.maDeList.length) throw new Error("Ph\u1EA3i c\xF3 \xEDt nh\u1EA5t 1 m\xE3 \u0111\u1EC1");
    }
    var nhom = getNhomMap();
    var cacNhom = (p.cacNhom || []).map(normMa).filter(function(x) {
      return x;
    });
    cacNhom.forEach(function(ma) {
      var info = nhom[ma];
      if (!info) throw new Error("Nh\xF3m " + ma + " ch\u01B0a \u0111\u0103ng k\xFD");
      if (u.role !== "admin" && info.teacher !== u.username) throw new Error("Nh\xF3m " + ma + " kh\xF4ng thu\u1ED9c b\u1EA1n");
    });
    var teacher = u.role === "admin" && p.teacher ? p.teacher : u.username;
    var folder = isWorker() ? { getId: /* @__PURE__ */ __name(function() {
      return "";
    }, "getId") } : createBaiFolder(teacher, p.tenBai, p.ngay);
    var maBai;
    do {
      maBai = "B" + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 4).toUpperCase();
    } while (sysRows("BaiKiemTra").some(function(r) {
      return String(r.maBai) === maBai;
    }));
    var sh = _bai(), row = sh.getLastRow() + 1;
    var vals = [maBai, (p.tenBai || "").trim(), p.ngay || "", teacher, cacNhom.join(","), JSON.stringify(cfg), "Nh\xE1p", folder.getId(), "", nowIso()];
    sh.getRange(row, 1, 1, vals.length).setNumberFormat("@");
    sh.getRange(row, 1, 1, vals.length).setValues([vals]);
    sysTouch();
    return { ok: true, maBai, driveFolderId: folder.getId() };
  }
  __name(createBai, "createBai");
  function updateBai(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    var cfg = p.cauHinh;
    if (typeof cfg === "string") cfg = JSON.parse(cfg);
    if (laBaiNgoai(bai) || cfg && cfg.loai === "ngoai") cfg = chuanCfgNgoai(cfg || {});
    else {
      validateCauHinh(cfg);
      cfg.maDeList = (cfg.maDeList || []).map(normMaDe).filter(function(x) {
        return x;
      });
      if (!cfg.maDeList.length) throw new Error("Ph\u1EA3i c\xF3 \xEDt nh\u1EA5t 1 m\xE3 \u0111\u1EC1");
    }
    var sh = _bai();
    var H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
    if (p.tenBai !== void 0) setTextCell(sh, bai._row, H["tenBai"] + 1, (p.tenBai || "").trim());
    if (p.ngay !== void 0) setTextCell(sh, bai._row, H["ngay"] + 1, p.ngay || "");
    if (p.cacNhom !== void 0) setTextCell(sh, bai._row, H["cacNhom"] + 1, (p.cacNhom || []).map(normMa).filter(Boolean).join(","));
    sh.getRange(bai._row, H["cauHinh"] + 1).setValue(JSON.stringify(cfg));
    sysTouch();
    return { ok: true };
  }
  __name(updateBai, "updateBai");
  function demTheoBai(tab, col1) {
    var sh = getSystemSheet().getSheetByName(tab), out = {};
    if (!sh || sh.getLastRow() < 2) return out;
    sh.getRange(2, col1, sh.getLastRow() - 1, 1).getValues().forEach(function(r) {
      var k = String(r[0]);
      out[k] = (out[k] || 0) + 1;
    });
    return out;
  }
  __name(demTheoBai, "demTheoBai");
  function listBai(u, p) {
    var rows = sysRows("BaiKiemTra").filter(function(r) {
      return u.role === "admin" || r.teacher === u.username;
    });
    rows.sort(function(a, b) {
      var x = ngayStr(a.ngay) || String(a.createdAt), y = ngayStr(b.ngay) || String(b.createdAt);
      return x < y ? 1 : x > y ? -1 : String(b.createdAt) < String(a.createdAt) ? -1 : 1;
    });
    var kq = demTheoBai("KetQua", 1), cx = demTheoBai("ChoXuLy", 2);
    return {
      bai: rows.map(function(r) {
        var cfg = parseJsonSafe(r.cauHinh, {});
        return {
          maBai: String(r.maBai),
          tenBai: String(r.tenBai),
          ngay: ngayStr(r.ngay),
          teacher: r.teacher,
          cacNhom: parseNhomList(r.cacNhom),
          trangThai: r.trangThai,
          soMaDe: (cfg.maDeList || []).length,
          luuTru: toBool(r.luuTru),
          soKetQua: kq[String(r.maBai)] || 0,
          soChoXuLy: cx[String(r.maBai)] || 0,
          createdAt: String(r.createdAt || "")
        };
      })
    };
  }
  __name(listBai, "listBai");
  function getBaiDapAn(maBai) {
    var map = {};
    sysRows("DapAn").filter(function(r) {
      return String(r.maBai) === String(maBai);
    }).forEach(function(r) {
      var d = {};
      try {
        d = JSON.parse(r.dapAn || "{}");
      } catch (e) {
      }
      map[normMaDe(r.maDe)] = d;
    });
    return map;
  }
  __name(getBaiDapAn, "getBaiDapAn");
  function getBai(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    var cfg = {};
    try {
      cfg = JSON.parse(bai.cauHinh || "{}");
    } catch (e) {
    }
    var hasFile = false, fileId = String(bai.dapAnFileId || "");
    if (fileId) hasFile = true;
    else if (!isWorker()) {
      try {
        var folder = DriveApp.getFolderById(bai.driveFolderId);
        var it = folder.getFilesByName("dap-an - " + safeName(bai.tenBai) + ".xlsx");
        if (it.hasNext()) {
          hasFile = true;
          fileId = it.next().getId();
        }
      } catch (e) {
      }
    }
    return {
      maBai: bai.maBai,
      tenBai: bai.tenBai,
      ngay: ngayStr(bai.ngay),
      teacher: bai.teacher,
      cacNhom: parseNhomList(bai.cacNhom),
      cauHinh: cfg,
      maDeList: cfg.maDeList || [],
      trangThai: bai.trangThai,
      driveFolderId: bai.driveFolderId,
      dapAn: getBaiDapAn(bai.maBai),
      soKetQua: countKetQua(bai.maBai),
      soChoXuLy: countChoXuLy(bai.maBai),
      luuTru: toBool(bai.luuTru),
      dapAnFile: hasFile ? { fileId, viewUrl: "https://drive.google.com/file/d/" + fileId + "/view", downloadUrl: "https://drive.google.com/uc?export=download&id=" + fileId } : null
    };
  }
  __name(getBai, "getBai");
  function saveDapAn(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    var incoming = p.dapAn;
    if (typeof incoming === "string") incoming = JSON.parse(incoming);
    incoming = incoming || {};
    var only = p.maDe ? normMaDe(p.maDe) : null;
    var lock = LockService.getScriptLock();
    lock.waitLock(2e4);
    try {
      var sh = _dapan();
      var del = sysRows("DapAn").filter(function(r) {
        return String(r.maBai) === String(p.maBai) && (!only || normMaDe(r.maDe) === only);
      }).map(function(r) {
        return r._row;
      });
      if (del.length) deleteRowsList(sh, del);
      var rows = Object.keys(incoming).filter(function(md) {
        return !only || normMaDe(md) === only;
      }).map(function(md) {
        return [String(p.maBai), normMaDe(md), JSON.stringify(incoming[md])];
      });
      if (rows.length) {
        var r0 = sh.getLastRow() + 1;
        sh.getRange(r0, 1, rows.length, 2).setNumberFormat("@");
        sh.getRange(r0, 1, rows.length, 3).setValues(rows);
      }
      sysTouch();
      return { ok: true, luc: nowIso() };
    } finally {
      lock.releaseLock();
    }
  }
  __name(saveDapAn, "saveDapAn");
  function genDapAnFile(p, u) {
    var bai = getBai(p, u);
    var cfg = bai.cauHinh, maDeList = bai.maDeList;
    if (!maDeList.length) throw new Error("Ch\u01B0a c\xF3 m\xE3 \u0111\u1EC1");
    var parts = [];
    if (cfg.p1 && cfg.p1.on) parts.push(["p1", cfg.p1.soCau, "Ph\u1EA7n I"]);
    if (cfg.p2 && cfg.p2.on) parts.push(["p2", cfg.p2.soCau, "Ph\u1EA7n II"]);
    if (cfg.p3 && cfg.p3.on) parts.push(["p3", cfg.p3.soCau, "Ph\u1EA7n III"]);
    var thieu = [];
    maDeList.forEach(function(md) {
      var da = bai.dapAn[md] || {};
      parts.forEach(function(pt) {
        var arr = da[pt[0]] || [];
        for (var i = 0; i < pt[1]; i++) {
          var v = arr[i];
          if (v === void 0 || v === null || String(v).trim() === "") thieu.push("M\xE3 " + md + " \xB7 " + pt[2] + " c\xE2u " + (i + 1));
        }
      });
    });
    if (thieu.length) throw new Error("C\xF2n thi\u1EBFu \u0111\xE1p \xE1n: " + thieu.slice(0, 8).join("; ") + (thieu.length > 8 ? " \u2026" : ""));
    var header = ["CAU\\MA DE"].concat(maDeList);
    var dataRows = [];
    var cau = 1;
    parts.forEach(function(pt) {
      for (var i = 0; i < pt[1]; i++) {
        var row = [cau];
        maDeList.forEach(function(md) {
          var arr = (bai.dapAn[md] || {})[pt[0]] || [];
          var val = arr[i] !== void 0 ? arr[i] : "";
          if (pt[0] === "p3") val = String(val).replace(".", ",");
          else val = String(val);
          row.push(val);
        });
        dataRows.push(row);
        cau++;
      }
    });
    var tmp = SpreadsheetApp.create("tmp-dapan-" + bai.maBai);
    try {
      var s2 = tmp.insertSheet("DapAn");
      var nCol = header.length, nRow = dataRows.length + 1;
      s2.getRange(1, 1, 1, nCol).setNumberFormat("@");
      if (dataRows.length) s2.getRange(2, 2, dataRows.length, nCol - 1).setNumberFormat("@");
      s2.getRange(1, 1, nRow, nCol).setValues([header].concat(dataRows));
      SpreadsheetApp.flush();
      var blob = exportSpreadsheetAsXlsxBlob(tmp.getId(), "dap-an.xlsx");
      var folder = DriveApp.getFolderById(bai.driveFolderId);
      var name = "dap-an - " + safeName(bai.tenBai) + ".xlsx";
      var file = saveBlobToFolder(folder, blob, name);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return { ok: true, fileId: file.getId(), viewUrl: "https://drive.google.com/file/d/" + file.getId() + "/view", downloadUrl: "https://drive.google.com/uc?export=download&id=" + file.getId() };
    } finally {
      try {
        DriveApp.getFileById(tmp.getId()).setTrashed(true);
      } catch (e) {
      }
    }
  }
  __name(genDapAnFile, "genDapAnFile");
  function getBaiFull(p, u) {
    var b = getBai(p, u);
    if (isWorker()) {
      var bb = findBai(p.maBai);
      b.anh = { count: Number(bb && bb.soAnh) || 0, folderUrl: bb && bb.driveFolderId ? "https://drive.google.com/drive/folders/" + bb.driveFolderId : "" };
    } else {
      try {
        b.anh = anhInfo(p, u);
      } catch (e) {
        b.anh = { count: 0, folderUrl: "", loi: String(e.message || e) };
      }
    }
    return b;
  }
  __name(getBaiFull, "getBaiFull");
  function luuTruBai(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    var sh = _bai(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
    if (H["luuTru"] === void 0) throw new Error("D\u1EEF li\u1EC7u ch\u01B0a \u0111\u01B0\u1EE3c n\xE2ng c\u1EA5p, h\xE3y t\u1EA3i l\u1EA1i trang");
    sh.getRange(bai._row, H["luuTru"] + 1).setValue(p.luuTru ? "TRUE" : "");
    sysTouch();
    return { ok: true, luuTru: !!p.luuTru };
  }
  __name(luuTruBai, "luuTruBai");
  function xoaBai(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    if (String(p.xacNhan || "").trim() !== String(bai.tenBai).trim()) throw new Error("G\xF5 l\u1EA1i \u0111\xFAng t\xEAn b\xE0i \u0111\u1EC3 x\xE1c nh\u1EADn x\xF3a");
    var daCongBo = bai.trangThai === "\u0110\xE3 c\xF4ng b\u1ED1";
    var lock = LockService.getScriptLock();
    lock.waitLock(3e4);
    try {
      var kqRows = kqIndex(bai.maBai).map(function(x) {
        return x.row;
      });
      if (kqRows.length) deleteRowsList(_kq(), kqRows);
      var cx = getSystemSheet().getSheetByName("ChoXuLy");
      if (cx) {
        var cr = cxlLoad(bai.maBai).map(function(x) {
          return x._row;
        });
        if (cr.length) deleteRowsList(cx, cr);
      }
      var da = sysRows("DapAn").filter(function(r) {
        return String(r.maBai) === String(bai.maBai);
      }).map(function(r) {
        return r._row;
      });
      if (da.length) deleteRowsList(_dapan(), da);
      var cur = findBai(bai.maBai);
      if (cur) deleteRowsList(_bai(), [cur._row]);
      try {
        DriveApp.getFolderById(bai.driveFolderId).setTrashed(true);
      } catch (e) {
      }
      CacheService2.getScriptCache().remove("baimeta_" + bai.maBai);
      logSua(u, bai.maBai, "(b\xE0i)", "x\xF3a b\xE0i", String(bai.tenBai), "(\u0111\xE3 x\xF3a)", "x\xF3a h\u1EB3n b\xE0i");
      sysTouch();
      if (daCongBo) pubTouch();
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  }
  __name(xoaBai, "xoaBai");
  function luuFileDrive(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    var data = String(p.data || "");
    if (!data) throw new Error("File r\u1ED7ng");
    if (data.length > 14 * 1024 * 1024) throw new Error("File qu\xE1 l\u1EDBn");
    var name = String(p.tenFile || "file.xlsx").replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 120) || "file.xlsx";
    var mime = /\.xlsx$/i.test(name) ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/octet-stream";
    var folder = DriveApp.getFolderById(bai.driveFolderId);
    var file = saveBlobToFolder(folder, Utilities.newBlob(Utilities.base64Decode(data), mime, name), name);
    return { ok: true, url: "https://drive.google.com/file/d/" + file.getId() + "/view" };
  }
  __name(luuFileDrive, "luuFileDrive");
  var DIEM_TOL = 0.011;
  var TRANG_THAI_CHUA_CO_DIEM = ["Nh\xE1p", "Ch\u1EDD ch\u1EA5m", "S\u1EB5n s\xE0ng ch\u1EA5m"];
  function _kq() {
    var ss = getSystemSheet();
    var cache = CacheService2.getScriptCache();
    if (!cache.get("kqcols_ok2")) {
      ensureSysColumns(ss);
      cache.put("kqcols_ok2", "1", 21600);
    }
    return ss.getSheetByName("KetQua");
  }
  __name(_kq, "_kq");
  function kqHeader(sh) {
    return idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  }
  __name(kqHeader, "kqHeader");
  function kqIndex(maBai) {
    var sh = _kq(), last = sh.getLastRow();
    if (last < 2) return [];
    var v = sh.getRange(2, 1, last - 1, 2).getValues(), out = [];
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0]) === String(maBai)) out.push({ row: i + 2, sbd: normSbd(v[i][1]) });
    }
    return out;
  }
  __name(kqIndex, "kqIndex");
  function countKetQua(maBai) {
    try {
      return kqIndex(maBai).length;
    } catch (e) {
      return 0;
    }
  }
  __name(countKetQua, "countKetQua");
  function parseJsonSafe(s, def) {
    try {
      return JSON.parse(s);
    } catch (e) {
      return def;
    }
  }
  __name(parseJsonSafe, "parseJsonSafe");
  function kqLoad(maBai) {
    var ix = kqIndex(maBai);
    if (!ix.length) return [];
    var sh = _kq(), H = kqHeader(sh), nCol = sh.getLastColumn();
    var r0 = ix[0].row, r1 = ix[ix.length - 1].row;
    var vals = sh.getRange(r0, 1, r1 - r0 + 1, nCol).getValues(), out = [];
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i];
      if (String(v[H["maBai"]]) !== String(maBai)) continue;
      out.push({
        _row: r0 + i,
        sbd: normSbd(v[H["sbd"]]),
        maDe: normMaDe(v[H["maDe"]]),
        tong: Number(v[H["tong"]]),
        p1: Number(v[H["p1"]]),
        p2: Number(v[H["p2"]]),
        p3: Number(v[H["p3"]]),
        tl: parseJsonSafe(v[H["traLoi"]], { p1: [], p2: [], p3: [] }),
        daSua: toBool(v[H["daSua"]]),
        nguon: H["nguon"] !== void 0 ? normSbd(v[H["nguon"]]) || "" : ""
        // SBD ghi trong file TNMaker, nếu đã được gán sang em khác
      });
    }
    return out;
  }
  __name(kqLoad, "kqLoad");
  function kqRowArray(H, nCol, bai, o) {
    var a = [];
    for (var i = 0; i < nCol; i++) a.push("");
    a[H["maBai"]] = bai;
    a[H["sbd"]] = o.sbd;
    a[H["maDe"]] = o.maDe;
    a[H["tong"]] = o.tong;
    a[H["p1"]] = o.p1;
    a[H["p2"]] = o.p2;
    a[H["p3"]] = o.p3;
    a[H["traLoi"]] = JSON.stringify(o.tl || {});
    a[H["capNhat"]] = nowIso();
    a[H["daSua"]] = o.daSua ? "TRUE" : "";
    if (H["nguon"] !== void 0) a[H["nguon"]] = o.nguon || "";
    return a;
  }
  __name(kqRowArray, "kqRowArray");
  function deleteRowsList(sh, rows) {
    rows = rows.slice().sort(function(a, b) {
      return b - a;
    });
    var i = 0;
    while (i < rows.length) {
      var end = rows[i], start = end, j = i + 1;
      while (j < rows.length && rows[j] === start - 1) {
        start = rows[j];
        j++;
      }
      sh.deleteRows(start, end - start + 1);
      i = j;
    }
    sysTouch();
  }
  __name(deleteRowsList, "deleteRowsList");
  function tStr(v) {
    if (Object.prototype.toString.call(v) === "[object Date]") return Utilities.formatDate(v, "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm");
    return v === null || v === void 0 ? "" : String(v);
  }
  __name(tStr, "tStr");
  function logSua(user, maBai, sbd, truong, cu, moi, lyDo) {
    var sh = getSystemSheet().getSheetByName("NhatKySua"), row = sh.getLastRow() + 1;
    var vals = [nowIso(), user.username, maBai, sbd, truong, String(cu), String(moi), lyDo || ""];
    sh.getRange(row, 1, 1, vals.length).setNumberFormat("@");
    sh.getRange(row, 1, 1, vals.length).setValues([vals]);
    sysTouch("NhatKySua");
  }
  __name(logSua, "logSua");
  function loadAccessibleGroups(u, fresh) {
    var nhom = getNhomMap(), byFile = {};
    Object.keys(nhom).forEach(function(ma) {
      var info = nhom[ma];
      if (u.role !== "admin" && info.teacher !== u.username) return;
      if (!info.rosterSheetId) return;
      (byFile[info.rosterSheetId] = byFile[info.rosterSheetId] || []).push({ ma, info });
    });
    var out = {};
    Object.keys(byFile).forEach(function(fid) {
      var tabs;
      try {
        tabs = rosterSnapshot(fid, !!fresh);
      } catch (e) {
        return;
      }
      byFile[fid].forEach(function(x) {
        var g = findTab(tabs, x.info.sheetName, x.ma);
        if (g) out[x.ma] = { maNhom: x.ma, tenNhom: g.tenNhom, teacher: x.info.teacher, students: g.students };
      });
    });
    return out;
  }
  __name(loadAccessibleGroups, "loadAccessibleGroups");
  function assertGroupAccessBySbd(u, sbd) {
    var info = getNhomMap()[sbd.substring(0, 3)];
    if (!info) throw new Error("SBD " + sbd + " thu\u1ED9c nh\xF3m ch\u01B0a \u0111\u0103ng k\xFD");
    if (u.role !== "admin" && info.teacher !== u.username) throw new Error("SBD " + sbd + " kh\xF4ng thu\u1ED9c nh\xF3m c\u1EE7a b\u1EA1n");
  }
  __name(assertGroupAccessBySbd, "assertGroupAccessBySbd");
  function keyComplete(cfg, k) {
    if (!k) return false;
    var i;
    if (cfg.p1 && cfg.p1.on) {
      for (i = 0; i < cfg.p1.soCau; i++) if (!k.p1 || !String(k.p1[i] || "").trim()) return false;
    }
    if (cfg.p2 && cfg.p2.on) {
      for (i = 0; i < cfg.p2.soCau; i++) if (!k.p2 || normP2(k.p2[i]).indexOf("_") !== -1) return false;
    }
    if (cfg.p3 && cfg.p3.on) {
      for (i = 0; i < cfg.p3.soCau; i++) if (!k.p3 || !normP3(k.p3[i])) return false;
    }
    return true;
  }
  __name(keyComplete, "keyComplete");
  function keyDiffs(cfg, a, b) {
    var d = [], i;
    if (cfg.p1 && cfg.p1.on) for (i = 0; i < cfg.p1.soCau; i++) {
      var x1 = String((a.p1 || [])[i] || "").toUpperCase(), y1 = String((b.p1 || [])[i] || "").toUpperCase();
      if (x1 !== y1) d.push("Ph\u1EA7n I c\xE2u " + (i + 1) + ": h\u1EC7 th\u1ED1ng " + x1 + ", file " + y1);
    }
    if (cfg.p2 && cfg.p2.on) for (i = 0; i < cfg.p2.soCau; i++) {
      var x2 = normP2((a.p2 || [])[i]), y2 = normP2((b.p2 || [])[i]);
      if (x2 !== y2) d.push("Ph\u1EA7n II c\xE2u " + (i + 1) + ": h\u1EC7 th\u1ED1ng " + x2 + ", file " + y2);
    }
    if (cfg.p3 && cfg.p3.on) for (i = 0; i < cfg.p3.soCau; i++) {
      var x3 = normP3((a.p3 || [])[i]), y3 = normP3((b.p3 || [])[i]);
      if (x3 !== y3) d.push("Ph\u1EA7n III c\xE2u " + (i + 1) + ": h\u1EC7 th\u1ED1ng " + x3 + ", file " + y3);
    }
    return d;
  }
  __name(keyDiffs, "keyDiffs");
  function shapeOfCfg(cfg) {
    return {
      n1: cfg.p1 && cfg.p1.on ? Number(cfg.p1.soCau) : 0,
      n2: cfg.p2 && cfg.p2.on ? Number(cfg.p2.soCau) : 0,
      n3: cfg.p3 && cfg.p3.on ? Number(cfg.p3.soCau) : 0
    };
  }
  __name(shapeOfCfg, "shapeOfCfg");
  function readRows(p) {
    var rows = p.rows;
    if (typeof rows === "string") rows = JSON.parse(rows);
    if (!rows || !rows.length) throw new Error("File kh\xF4ng c\xF3 d\xF2ng \u0111i\u1EC3m n\xE0o");
    if (rows.length > 3e3) throw new Error("File qu\xE1 l\u1EDBn (tr\xEAn 3000 d\xF2ng)");
    return rows;
  }
  __name(readRows, "readRows");
  function previewDiem(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    if (laBaiNgoai(bai)) throw new Error('B\xE0i n\xE0y ch\u1EC9 c\xF3 \u0111i\u1EC3m t\u1ED5ng: d\xF9ng m\u1EE5c "Nh\u1EADp \u0111i\u1EC3m" c\u1EE7a b\xE0i');
    var cfg = parseJsonSafe(bai.cauHinh, {});
    var rows = readRows(p);
    var keys = p.keys || {};
    if (typeof keys === "string") keys = JSON.parse(keys);
    var groups = loadAccessibleGroups(u, true);
    var bySbd = {}, byCu = {};
    Object.keys(groups).forEach(function(ma) {
      groups[ma].students.forEach(function(s) {
        if (s.sbd) (bySbd[s.sbd] = bySbd[s.sbd] || []).push({ ho: s.ho, ten: s.ten, maNhom: ma, trangThai: s.trangThai });
        if (s.sbdCu) byCu[s.sbdCu] = s.sbd;
      });
    });
    var existing = {}, existCount = 0, edited = [], byNguon = {};
    kqLoad(bai.maBai).forEach(function(r) {
      existing[r.sbd] = r;
      existCount++;
      if (r.daSua) edited.push(r.sbd);
      if (r.nguon) byNguon[r.nguon] = r;
    });
    var cacNhom = parseNhomList(bai.cacNhom);
    var maDeList = (cfg.maDeList || []).map(normMaDe);
    var count = {};
    rows.forEach(function(r) {
      var s = normSbd(r.sbd);
      count[s] = (count[s] || 0) + 1;
    });
    var out = [], present = {}, khop = 0, lech = 0, khongKiemTra = 0;
    rows.forEach(function(r, i) {
      var sbd = normSbd(r.sbd), maDe = normMaDe(r.maDe);
      var o = { i, sbd, maDe, tong: r.tong, issues: [] };
      var list = bySbd[sbd] || [];
      if (!/^\d{6}$/.test(sbd)) {
        o.issues.push({ code: "KHONG_CO_SBD", level: "error", msg: "SBD kh\xF4ng h\u1EE3p l\u1EC7: " + r.sbd });
      } else if (!list.length) {
        o.issues.push({ code: "KHONG_CO_SBD", level: "error", msg: "SBD kh\xF4ng c\xF3 trong danh s\xE1ch" });
      } else if (list.length > 1) {
        o.issues.push({ code: "TRUNG_DS", level: "error", msg: "SBD \u0111ang d\xF9ng cho " + list.length + " h\u1ECDc sinh: " + list.map(function(x) {
          return (x.ho + " " + x.ten).trim() + " (nh\xF3m " + x.maNhom + ")";
        }).join("; ") + ". H\xE3y s\u1EEDa danh s\xE1ch tr\u01B0\u1EDBc" });
      } else {
        var st = list[0];
        o.ho = st.ho;
        o.ten = st.ten;
        o.maNhom = st.maNhom;
        if (st.trangThai === "\u0110\xE3 ngh\u1EC9") o.issues.push({ code: "DA_NGHI", level: "warn", msg: "H\u1ECDc sinh n\xE0y \u0111\xE3 ngh\u1EC9" });
        if (st.trangThai === "\u0110\xE3 chuy\u1EC3n") {
          o.goiY = byCu[sbd] || "";
          o.issues.push({ code: "DA_CHUYEN", level: "warn", msg: "H\u1ECDc sinh \u0111\xE3 chuy\u1EC3n nh\xF3m" + (o.goiY ? ", SBD m\u1EDBi l\xE0 " + o.goiY : "") });
        }
        if (cacNhom.length && cacNhom.indexOf(st.maNhom) === -1) o.issues.push({ code: "NGOAI_BAI", level: "warn", msg: "Nh\xF3m " + st.maNhom + " ch\u01B0a n\u1EB1m trong danh s\xE1ch nh\xF3m l\xE0m b\xE0i (s\u1EBD \u0111\u01B0\u1EE3c th\xEAm v\xE0o)" });
      }
      var gan = byNguon[sbd];
      if (gan && o.issues.length && (o.issues[0].code === "KHONG_CO_SBD" || o.issues[0].code === "TRUNG_DS")) {
        var who = (bySbd[gan.sbd] || [])[0];
        o.issues = [{ code: "DA_GAN", level: "warn", msg: "L\u1EA7n nh\u1EADp tr\u01B0\u1EDBc \u0111\xE3 g\xE1n d\xF2ng n\xE0y cho SBD " + gan.sbd + (who ? " (" + (who.ho + " " + who.ten).trim() + ")" : "") }];
        o.daGan = gan.sbd;
        if (gan.daSua) o.issues.push({ code: "DA_SUA_TAY", level: "warn", msg: "K\u1EBFt qu\u1EA3 c\u1EE7a em n\xE0y \u0111\xE3 \u0111\u01B0\u1EE3c s\u1EEDa tay tr\u01B0\u1EDBc \u0111\xF3" });
      }
      if (count[sbd] > 1) o.issues.push({ code: "TRUNG_FILE", level: "warn", msg: "SBD xu\u1EA5t hi\u1EC7n " + count[sbd] + " l\u1EA7n trong file" });
      if (maDe && maDeList.indexOf(maDe) === -1) o.issues.push({ code: "MA_DE_LA", level: "warn", msg: "M\xE3 \u0111\u1EC1 " + maDe + " ch\u01B0a c\xF3 trong b\xE0i (s\u1EBD \u0111\u01B0\u1EE3c th\xEAm v\xE0o)" });
      if (existing[sbd]) {
        o.daCo = true;
        if (existing[sbd].daSua) o.issues.push({ code: "DA_SUA_TAY", level: "warn", msg: "K\u1EBFt qu\u1EA3 n\xE0y \u0111\xE3 \u0111\u01B0\u1EE3c s\u1EEDa tay tr\u01B0\u1EDBc \u0111\xF3" });
      }
      if (r.tl && keys[maDe]) {
        var sc = chamBai(cfg, r.tl, keys[maDe]);
        if (sc.thieuKey) khongKiemTra++;
        else if (Math.abs(sc.tong - r.tong) > DIEM_TOL || Math.abs(sc.p1 - r.p1) > DIEM_TOL || Math.abs(sc.p2 - r.p2) > DIEM_TOL || Math.abs(sc.p3 - r.p3) > DIEM_TOL) {
          lech++;
          o.issues.push({ code: "LECH_DIEM", level: "warn", msg: "Web t\xEDnh " + sc.tong + " \u0111i\u1EC3m, TNMaker " + r.tong + " (ki\u1EC3m tra c\u1EA5u h\xECnh b\xE0i tr\xEAn app TNMaker)" });
        } else khop++;
      } else khongKiemTra++;
      if (/^\d{6}$/.test(sbd)) present[sbd] = true;
      out.push(o);
    });
    var maHienCo = cacNhom.length ? cacNhom : Object.keys(out.reduce(function(acc, o) {
      if (o.maNhom) acc[o.maNhom] = 1;
      return acc;
    }, {}));
    var absent = [];
    maHienCo.forEach(function(ma) {
      var g = groups[ma];
      if (!g) return;
      g.students.forEach(function(s) {
        if (s.trangThai === "\u0110ang h\u1ECDc" && s.sbd && !present[s.sbd] && !existing[s.sbd]) absent.push({ sbd: s.sbd, ho: s.ho, ten: s.ten, maNhom: ma });
      });
    });
    absent.sort(function(a, b) {
      return a.sbd < b.sbd ? -1 : 1;
    });
    var dapAn = getBaiDapAn(bai.maBai), keyConflicts = [];
    Object.keys(keys).forEach(function(md) {
      var e = dapAn[normMaDe(md)];
      if (e && keyComplete(cfg, e)) {
        var d = keyDiffs(cfg, e, keys[md]);
        if (d.length) keyConflicts.push({ maDe: normMaDe(md), soCau: d.length, chiTiet: d.slice(0, 6) });
      }
    });
    var shapeWarn = "";
    if (p.shape) {
      var sc0 = shapeOfCfg(cfg), sh = p.shape;
      if (sc0.n1 !== sh.n1 || sc0.n2 !== sh.n2 || sc0.n3 !== sh.n3) {
        shapeWarn = "File c\xF3 Ph\u1EA7n I " + sh.n1 + " c\xE2u, Ph\u1EA7n II " + sh.n2 + " c\xE2u, Ph\u1EA7n III " + sh.n3 + " c\xE2u; c\u1EA5u h\xECnh b\xE0i tr\xEAn web l\xE0 " + sc0.n1 + " / " + sc0.n2 + " / " + sc0.n3 + ". N\xEAn s\u1EEDa c\u1EA5u h\xECnh b\xE0i cho kh\u1EDBp \u0111\u1EC3 t\xEDnh l\u1EA1i \u0111i\u1EC3m ch\xEDnh x\xE1c.";
      }
    }
    var tenNhom = {};
    Object.keys(groups).forEach(function(ma) {
      tenNhom[ma] = groups[ma].tenNhom;
    });
    return {
      rows: out,
      absent,
      tenNhom,
      existing: { count: existCount, edited },
      keyConflicts,
      shapeWarn,
      doiChieu: { khop, lech, khongKiemTra },
      tenBai: bai.tenBai,
      choXuLy: countChoXuLy(bai.maBai)
    };
  }
  __name(previewDiem, "previewDiem");
  function numOrThrow(v, name, sbd) {
    var n = Number(v);
    if (!isFinite(n) || n < 0 || n > 10) throw new Error("\u0110i\u1EC3m " + name + " kh\xF4ng h\u1EE3p l\u1EC7 (SBD " + sbd + ")");
    return n;
  }
  __name(numOrThrow, "numOrThrow");
  function commitDiem(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    if (laBaiNgoai(bai)) throw new Error('B\xE0i n\xE0y ch\u1EC9 c\xF3 \u0111i\u1EC3m t\u1ED5ng: d\xF9ng m\u1EE5c "Nh\u1EADp \u0111i\u1EC3m" c\u1EE7a b\xE0i');
    var cfg = parseJsonSafe(bai.cauHinh, {});
    var mode = p.mode === "replace" ? "replace" : "merge";
    var rows = p.rows;
    if (typeof rows === "string") rows = JSON.parse(rows);
    rows = rows || [];
    var deferred = p.deferred || [];
    if (typeof deferred === "string") deferred = JSON.parse(deferred);
    var purge = p.purge || [];
    if (typeof purge === "string") purge = JSON.parse(purge);
    if (!rows.length && !deferred.length && mode !== "replace") throw new Error("Kh\xF4ng c\xF3 d\xF2ng n\xE0o \u0111\u1EC3 ghi");
    var keys = p.keys || {};
    if (typeof keys === "string") keys = JSON.parse(keys);
    var nhom = getNhomMap(), seen = {}, clean = [], addNhom = {}, maDeNew = {}, purgeSbd = {}, maDeKeys = {};
    purge.forEach(function(x) {
      var n = normSbd(x);
      if (n) purgeSbd[n] = true;
    });
    var pendRows = deferred.map(function(r) {
      var md = normMaDe(r.maDe), sb = normSbd(r.sbd) || String(r.sbd == null ? "" : r.sbd);
      if (md) maDeKeys[md] = true;
      if (normSbd(r.sbd)) purgeSbd[normSbd(r.sbd)] = true;
      var num = /* @__PURE__ */ __name(function(v) {
        var n = Number(v);
        return isFinite(n) ? n : 0;
      }, "num");
      return {
        sbd: sb,
        maDe: md,
        tong: num(r.tong),
        p1: num(r.p1),
        p2: num(r.p2),
        p3: num(r.p3),
        tl: r.tl || null,
        lyDo: String(r.lyDo || "").slice(0, 30),
        ghiChu: String(r.ghiChu || "").slice(0, 300)
      };
    });
    rows.forEach(function(r) {
      var sbd = normSbd(r.sbd);
      if (!/^\d{6}$/.test(sbd)) throw new Error("SBD kh\xF4ng h\u1EE3p l\u1EC7: " + r.sbd);
      var ma = sbd.substring(0, 3), info = nhom[ma];
      if (!info) throw new Error("SBD " + sbd + " thu\u1ED9c nh\xF3m " + ma + " ch\u01B0a \u0111\u0103ng k\xFD");
      if (u.role !== "admin" && info.teacher !== u.username) throw new Error("SBD " + sbd + " kh\xF4ng thu\u1ED9c nh\xF3m c\u1EE7a b\u1EA1n");
      if (seen[sbd]) throw new Error("SBD " + sbd + " b\u1ECB tr\xF9ng sau khi x\u1EED l\xFD");
      seen[sbd] = true;
      var maDe = normMaDe(r.maDe);
      if (!maDe) throw new Error("Thi\u1EBFu m\xE3 \u0111\u1EC1 (SBD " + sbd + ")");
      addNhom[ma] = true;
      maDeNew[maDe] = true;
      var srcS = normSbd(r.srcSbd);
      purgeSbd[srcS || sbd] = true;
      clean.push({ nguon: srcS && srcS !== sbd ? srcS : "", sbd, maDe, tong: numOrThrow(r.tong, "t\u1ED5ng", sbd), p1: numOrThrow(r.p1, "P1", sbd), p2: numOrThrow(r.p2, "P2", sbd), p3: numOrThrow(r.p3, "P3", sbd), tl: r.tl || null, daSua: false });
    });
    var lock = LockService.getScriptLock();
    lock.waitLock(3e4);
    try {
      var sh = _kq(), H = kqHeader(sh), nCol = sh.getLastColumn();
      var ix = kqIndex(bai.maBai), del = [];
      if (mode === "replace") ix.forEach(function(x) {
        del.push(x.row);
      });
      else ix.forEach(function(x) {
        if (seen[x.sbd]) del.push(x.row);
      });
      if (del.length) deleteRowsList(sh, del);
      kqAppendRows(bai.maBai, clean);
      Object.keys(maDeNew).forEach(function(m) {
        maDeKeys[m] = true;
      });
      var da = _dapan(), mdKeys = Object.keys(keys).map(normMaDe).filter(function(m) {
        return maDeKeys[m];
      });
      if (mdKeys.length) {
        var daDel = [];
        sysRows("DapAn").forEach(function(r) {
          if (String(r.maBai) === String(bai.maBai) && mdKeys.indexOf(normMaDe(r.maDe)) !== -1) daDel.push(r._row);
        });
        if (daDel.length) deleteRowsList(da, daDel);
        var daRows = [];
        Object.keys(keys).forEach(function(md) {
          var m = normMaDe(md);
          if (!maDeKeys[m]) return;
          daRows.push([bai.maBai, m, JSON.stringify(keys[md])]);
        });
        if (daRows.length) da.getRange(da.getLastRow() + 1, 1, daRows.length, 3).setValues(daRows);
      }
      cxlDelete(bai.maBai, function(r) {
        return mode === "replace" || purgeSbd[r.sbd];
      });
      cxlAppend(bai.maBai, pendRows, p.tenFile);
      var tt = clean.length || mode === "replace" ? capNhatBaiSauKhiGhi(bai, cfg, maDeNew, addNhom) : findBai(bai.maBai).trangThai;
      logSua(
        u,
        bai.maBai,
        "(file)",
        "nh\u1EADp file \u0111i\u1EC3m",
        mode === "replace" ? "x\xF3a h\u1EBFt k\u1EBFt qu\u1EA3 c\u0169" : del.length + " d\xF2ng c\u0169 b\u1ECB thay",
        clean.length + " d\xF2ng" + (pendRows.length ? ", " + pendRows.length + " d\xF2ng \u0111\u1EC3 x\u1EED l\xFD sau" : "") + (p.tenFile ? " t\u1EEB " + p.tenFile : ""),
        ""
      );
      sysTouch();
      if (tt === "\u0110\xE3 c\xF4ng b\u1ED1") pubTouch();
      return { ok: true, soDong: clean.length, soChoXuLy: countChoXuLy(bai.maBai), trangThai: tt };
    } finally {
      lock.releaseLock();
    }
  }
  __name(commitDiem, "commitDiem");
  function getKetQua(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    return cachedResult("kq_" + bai.maBai + "_" + u.username, 900, function() {
      return getKetQua_(p, u, bai);
    }, p.fresh);
  }
  __name(getKetQua, "getKetQua");
  function getKetQua_(p, u, bai) {
    var cfg = parseJsonSafe(bai.cauHinh, {});
    var res = kqLoad(bai.maBai), groups = loadAccessibleGroups(u, !!p.fresh);
    var cacNhom = parseNhomList(bai.cacNhom);
    var maSet = {};
    cacNhom.forEach(function(m) {
      maSet[m] = true;
    });
    res.forEach(function(r) {
      maSet[r.sbd.substring(0, 3)] = true;
    });
    var bySbd = {};
    res.forEach(function(r) {
      bySbd[r.sbd] = r;
    });
    var used = {}, outGroups = [];
    Object.keys(maSet).sort().forEach(function(ma) {
      var g = groups[ma];
      if (!g) return;
      var list = [];
      g.students.forEach(function(s) {
        var r = bySbd[s.sbd];
        if (s.trangThai !== "\u0110ang h\u1ECDc" && !r) return;
        if (r) used[s.sbd] = true;
        list.push({
          sbd: s.sbd,
          ho: s.ho,
          ten: s.ten,
          lop: s.lop,
          trangThai: s.trangThai,
          sbdCu: s.sbdCu,
          kq: r ? { maDe: r.maDe, tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, daSua: r.daSua, tl: r.tl } : null
        });
      });
      list.sort(function(a, b) {
        return a.sbd < b.sbd ? -1 : 1;
      });
      outGroups.push({ maNhom: ma, tenNhom: g.tenNhom, students: list });
    });
    var bySbdAll = {};
    Object.keys(groups).forEach(function(ma) {
      groups[ma].students.forEach(function(s) {
        if (s.sbd) (bySbdAll[s.sbd] = bySbdAll[s.sbd] || []).push({ ho: s.ho, ten: s.ten, maNhom: ma, trangThai: s.trangThai });
      });
    });
    var choXuLy = cxlLoad(bai.maBai).map(function(r) {
      var o = { id: r.id, sbd: r.sbd, maDe: r.maDe, tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, lyDo: r.lyDo, ghiChu: r.ghiChu, tenFile: r.tenFile };
      var m = bySbdAll[r.sbd] || [];
      if (m.length === 1) o.match = { ho: m[0].ho, ten: m[0].ten, maNhom: m[0].maNhom, trangThai: m[0].trangThai, coKetQua: !!bySbd[r.sbd] };
      else if (m.length > 1) o.trungDS = m.length;
      return o;
    });
    var orphan = [];
    res.forEach(function(r) {
      if (!used[r.sbd] && groups[r.sbd.substring(0, 3)]) orphan.push({ sbd: r.sbd, kq: { maDe: r.maDe, tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, daSua: r.daSua, tl: r.tl } });
    });
    return {
      bai: { maBai: bai.maBai, tenBai: bai.tenBai, ngay: ngayStr(bai.ngay), trangThai: bai.trangThai, ngayCongBo: tStr(bai.ngayCongBo), cauHinh: cfg },
      groups: outGroups,
      orphan,
      choXuLy,
      keys: getBaiDapAn(bai.maBai)
    };
  }
  __name(getKetQua_, "getKetQua_");
  function normTraLoi(cfg, tl) {
    var out = { p1: [], p2: [], p3: [] }, i;
    tl = tl || {};
    var n1 = cfg.p1 && cfg.p1.on ? Number(cfg.p1.soCau) : 0;
    var n2 = cfg.p2 && cfg.p2.on ? Number(cfg.p2.soCau) : 0;
    var n3 = cfg.p3 && cfg.p3.on ? Number(cfg.p3.soCau) : 0;
    for (i = 0; i < n1; i++) {
      var a = String((tl.p1 || [])[i] || "").trim().toUpperCase();
      out.p1.push(/^[A-D]$/.test(a) ? a : "");
    }
    for (i = 0; i < n2; i++) out.p2.push(normP2((tl.p2 || [])[i]));
    for (i = 0; i < n3; i++) out.p3.push(normP3((tl.p3 || [])[i]));
    return out;
  }
  __name(normTraLoi, "normTraLoi");
  function editKetQua(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    var sbd = normSbd(p.sbd);
    if (!/^\d{6}$/.test(sbd)) throw new Error("SBD kh\xF4ng h\u1EE3p l\u1EC7");
    assertGroupAccessBySbd(u, sbd);
    var lyDo = String(p.lyDo || "").trim();
    if (!lyDo) throw new Error("Ph\u1EA3i ghi l\xFD do s\u1EEDa");
    var cfg = parseJsonSafe(bai.cauHinh, {});
    var lock = LockService.getScriptLock();
    lock.waitLock(2e4);
    try {
      var cur = kqLoad(bai.maBai).filter(function(r) {
        return r.sbd === sbd;
      })[0] || null;
      var maDe = normMaDe(p.maDe || cur && cur.maDe);
      if (!maDe) throw new Error("Thi\u1EBFu m\xE3 \u0111\u1EC1");
      var next = cur ? { maDe: cur.maDe, tong: cur.tong, p1: cur.p1, p2: cur.p2, p3: cur.p3, tl: cur.tl } : { maDe, tong: 0, p1: 0, p2: 0, p3: 0, tl: { p1: [], p2: [], p3: [] } };
      next.maDe = maDe;
      if (p.traLoi) {
        var key = getBaiDapAn(bai.maBai)[maDe];
        if (!key) throw new Error("Ch\u01B0a c\xF3 \u0111\xE1p \xE1n m\xE3 \u0111\u1EC1 " + maDe + ", kh\xF4ng t\xEDnh l\u1EA1i \u0111\u01B0\u1EE3c. H\xE3y s\u1EEDa \u0111i\u1EC3m tr\u1EF1c ti\u1EBFp.");
        var tl = normTraLoi(cfg, p.traLoi);
        var sc = chamBai(cfg, tl, key);
        if (sc.thieuKey) throw new Error("\u0110\xE1p \xE1n m\xE3 \u0111\u1EC1 " + maDe + " ch\u01B0a \u0111\u1EE7, kh\xF4ng t\xEDnh l\u1EA1i \u0111\u01B0\u1EE3c. H\xE3y s\u1EEDa \u0111i\u1EC3m tr\u1EF1c ti\u1EBFp.");
        next.tl = tl;
        next.p1 = sc.p1;
        next.p2 = sc.p2;
        next.p3 = sc.p3;
        next.tong = sc.tong;
      } else {
        if (p.p1 !== void 0 && p.p1 !== "") next.p1 = numOrThrow(p.p1, "P1", sbd);
        if (p.p2 !== void 0 && p.p2 !== "") next.p2 = numOrThrow(p.p2, "P2", sbd);
        if (p.p3 !== void 0 && p.p3 !== "") next.p3 = numOrThrow(p.p3, "P3", sbd);
        if (p.tong !== void 0 && p.tong !== "") next.tong = numOrThrow(p.tong, "t\u1ED5ng", sbd);
        else next.tong = round2(next.p1 + next.p2 + next.p3);
      }
      var changes = [];
      if (!cur || cur.maDe !== next.maDe) changes.push(["m\xE3 \u0111\u1EC1", cur ? cur.maDe : "(ch\u01B0a c\xF3)", next.maDe]);
      ["tong", "p1", "p2", "p3"].forEach(function(f) {
        var a = cur ? cur[f] : "", b = next[f];
        if (!cur || Math.abs(Number(a) - Number(b)) > 1e-9) changes.push([f === "tong" ? "t\u1ED5ng \u0111i\u1EC3m" : f.toUpperCase(), cur ? a : "(ch\u01B0a c\xF3)", b]);
      });
      if (p.traLoi && (!cur || JSON.stringify(cur.tl) !== JSON.stringify(next.tl))) changes.push(["c\xE2u tr\u1EA3 l\u1EDDi", cur ? "b\u1EA3n c\u0169" : "(ch\u01B0a c\xF3)", "\u0111\xE3 s\u1EEDa"]);
      if (!changes.length) throw new Error("Kh\xF4ng c\xF3 g\xEC thay \u0111\u1ED5i");
      var sh = _kq(), H = kqHeader(sh), nCol = sh.getLastColumn();
      var rowArr = kqRowArray(H, nCol, bai.maBai, { nguon: cur ? cur.nguon : "", sbd, maDe: next.maDe, tong: next.tong, p1: next.p1, p2: next.p2, p3: next.p3, tl: next.tl, daSua: true });
      if (cur) {
        sh.getRange(cur._row, 1, 1, nCol).setValues([rowArr]);
      } else {
        var start = sh.getLastRow() + 1;
        sh.getRange(start, H["sbd"] + 1, 1, 2).setNumberFormat("@");
        sh.getRange(start, 1, 1, nCol).setValues([rowArr]);
      }
      changes.forEach(function(c) {
        logSua(u, bai.maBai, sbd, c[0], c[1], c[2], lyDo);
      });
      var bs = _bai(), HB = idx(bs.getRange(1, 1, 1, bs.getLastColumn()).getValues()[0]);
      if (TRANG_THAI_CHUA_CO_DIEM.indexOf(bai.trangThai) !== -1) bs.getRange(bai._row, HB["trangThai"] + 1).setValue("\u0110\xE3 nh\u1EADp \u0111i\u1EC3m");
      sysTouch();
      if (bai.trangThai === "\u0110\xE3 c\xF4ng b\u1ED1") pubTouch();
      return { ok: true, kq: { maDe: next.maDe, tong: next.tong, p1: next.p1, p2: next.p2, p3: next.p3, daSua: true, tl: next.tl } };
    } finally {
      lock.releaseLock();
    }
  }
  __name(editKetQua, "editKetQua");
  function xoaKetQua(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    var sbd = normSbd(p.sbd);
    assertGroupAccessBySbd(u, sbd);
    var lyDo = String(p.lyDo || "").trim();
    if (!lyDo) throw new Error("Ph\u1EA3i ghi l\xFD do x\xF3a");
    var lock = LockService.getScriptLock();
    lock.waitLock(2e4);
    try {
      var cur = kqLoad(bai.maBai).filter(function(r) {
        return r.sbd === sbd;
      })[0];
      if (!cur) throw new Error("Em n\xE0y ch\u01B0a c\xF3 k\u1EBFt qu\u1EA3");
      deleteRowsList(_kq(), [cur._row]);
      logSua(u, bai.maBai, sbd, "x\xF3a k\u1EBFt qu\u1EA3", cur.tong + " \u0111i\u1EC3m (m\xE3 " + cur.maDe + ")", "(\u0111\xE3 x\xF3a)", lyDo);
      if (bai.trangThai === "\u0110\xE3 c\xF4ng b\u1ED1") pubTouch();
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  }
  __name(xoaKetQua, "xoaKetQua");
  function setCongBo(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    var bs = _bai(), HB = idx(bs.getRange(1, 1, 1, bs.getLastColumn()).getValues()[0]);
    var tt;
    if (p.congBo) {
      if (!countKetQua(bai.maBai)) throw new Error("B\xE0i ch\u01B0a c\xF3 \u0111i\u1EC3m \u0111\u1EC3 c\xF4ng b\u1ED1");
      tt = "\u0110\xE3 c\xF4ng b\u1ED1";
      bs.getRange(bai._row, HB["trangThai"] + 1).setValue(tt);
      bs.getRange(bai._row, HB["ngayCongBo"] + 1).setValue(nowIso());
    } else {
      tt = bai.trangThai === "\u0110\xE3 c\xF4ng b\u1ED1" ? "\u0110\xE3 nh\u1EADp \u0111i\u1EC3m" : bai.trangThai;
      bs.getRange(bai._row, HB["trangThai"] + 1).setValue(tt);
      bs.getRange(bai._row, HB["ngayCongBo"] + 1).setValue("");
    }
    logSua(u, bai.maBai, "(b\xE0i)", p.congBo ? "c\xF4ng b\u1ED1" : "\u1EA9n \u0111i\u1EC3m", bai.trangThai, tt, "");
    sysTouch();
    pubTouch();
    return { ok: true, trangThai: tt };
  }
  __name(setCongBo, "setCongBo");
  function getNhatKy(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    var rows = sysRows("NhatKySua").filter(function(r) {
      return String(r.maBai) === String(bai.maBai);
    });
    rows.reverse();
    return {
      nhatKy: rows.slice(0, 200).map(function(r) {
        return { thoiGian: tStr(r.thoiGian), nguoiSua: r.nguoiSua, sbd: normSbd(r.sbd) || String(r.sbd), truong: r.truong, cu: String(r.giaTriCu), moi: String(r.giaTriMoi), lyDo: r.lyDo };
      })
    };
  }
  __name(getNhatKy, "getNhatKy");
  function kqAppendRows(maBai, clean) {
    if (!clean.length) return;
    var sh = _kq(), H = kqHeader(sh), nCol = sh.getLastColumn();
    var arr = clean.map(function(o) {
      return kqRowArray(H, nCol, maBai, o);
    });
    var start = sh.getLastRow() + 1;
    sh.getRange(start, H["sbd"] + 1, arr.length, 2).setNumberFormat("@");
    sh.getRange(start, 1, arr.length, nCol).setValues(arr);
  }
  __name(kqAppendRows, "kqAppendRows");
  function capNhatBaiSauKhiGhi(bai, cfg, maDeNew, addNhom) {
    var bs = _bai(), HB = idx(bs.getRange(1, 1, 1, bs.getLastColumn()).getValues()[0]);
    var curBai = findBai(bai.maBai);
    var maDeList = (cfg.maDeList || []).map(normMaDe);
    Object.keys(maDeNew).forEach(function(m) {
      if (maDeList.indexOf(m) === -1) maDeList.push(m);
    });
    maDeList.sort();
    cfg.maDeList = maDeList;
    bs.getRange(curBai._row, HB["cauHinh"] + 1).setValue(JSON.stringify(cfg));
    var cacNhom = parseNhomList(curBai.cacNhom);
    Object.keys(addNhom).forEach(function(m) {
      if (cacNhom.indexOf(m) === -1) cacNhom.push(m);
    });
    cacNhom.sort();
    setTextCell(bs, curBai._row, HB["cacNhom"] + 1, cacNhom.join(","));
    var tt = curBai.trangThai;
    if (tt !== "\u0110\xE3 c\xF4ng b\u1ED1") {
      tt = "\u0110\xE3 nh\u1EADp \u0111i\u1EC3m";
      bs.getRange(curBai._row, HB["trangThai"] + 1).setValue(tt);
    }
    sysTouch();
    return tt;
  }
  __name(capNhatBaiSauKhiGhi, "capNhatBaiSauKhiGhi");
  function _cxl() {
    var ss = getSystemSheet(), sh = ss.getSheetByName("ChoXuLy");
    if (!sh) {
      sh = ss.insertSheet("ChoXuLy");
      sh.appendRow(SYS_TABS2.ChoXuLy);
    }
    return sh;
  }
  __name(_cxl, "_cxl");
  function cxlLoad(maBai) {
    var sh = getSystemSheet().getSheetByName("ChoXuLy");
    if (!sh) return [];
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return [];
    var H = idx(vals[0]), out = [];
    for (var i = 1; i < vals.length; i++) {
      var v = vals[i];
      if (String(v[H["maBai"]]) !== String(maBai)) continue;
      var raw = String(v[H["sbd"]]);
      out.push({
        _row: i + 1,
        id: String(v[H["id"]]),
        sbd: normSbd(raw) || raw,
        maDe: normMaDe(v[H["maDe"]]),
        tong: Number(v[H["tong"]]),
        p1: Number(v[H["p1"]]),
        p2: Number(v[H["p2"]]),
        p3: Number(v[H["p3"]]),
        tl: parseJsonSafe(v[H["traLoi"]], null),
        lyDo: String(v[H["lyDo"]]),
        ghiChu: String(v[H["ghiChu"]]),
        tenFile: String(v[H["tenFile"]])
      });
    }
    return out;
  }
  __name(cxlLoad, "cxlLoad");
  function countChoXuLy(maBai) {
    try {
      return cxlLoad(maBai).length;
    } catch (e) {
      return 0;
    }
  }
  __name(countChoXuLy, "countChoXuLy");
  function cxlAppend(maBai, rows, tenFile) {
    if (!rows.length) return;
    var sh = _cxl(), n = SYS_TABS2.ChoXuLy.length, now = nowIso(), base = Date.now().toString(36);
    var arr = rows.map(function(r, i) {
      return [
        "X" + base + "_" + i + "_" + Math.random().toString(36).slice(2, 6),
        maBai,
        r.sbd,
        r.maDe,
        r.tong,
        r.p1,
        r.p2,
        r.p3,
        JSON.stringify(r.tl || null),
        r.lyDo || "",
        r.ghiChu || "",
        tenFile || "",
        now
      ];
    });
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, arr.length, 4).setNumberFormat("@");
    sh.getRange(start, 1, arr.length, n).setValues(arr);
  }
  __name(cxlAppend, "cxlAppend");
  function cxlDelete(maBai, fn) {
    var sh = getSystemSheet().getSheetByName("ChoXuLy");
    if (!sh) return 0;
    var rows = cxlLoad(maBai).filter(fn).map(function(r) {
      return r._row;
    });
    if (rows.length) deleteRowsList(sh, rows);
    return rows.length;
  }
  __name(cxlDelete, "cxlDelete");
  function giaiQuyetChoXuLy(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    var cfg = parseJsonSafe(bai.cauHinh, {});
    if (cfg.loai === "ngoai") return ngoaiGiaiQuyet(p, u, bai);
    var lock = LockService.getScriptLock();
    lock.waitLock(3e4);
    try {
      var pend = cxlLoad(bai.maBai);
      if (!pend.length) return { ok: true, resolved: 0, discarded: 0, errors: [], conLai: 0 };
      var groups = loadAccessibleGroups(u, true), bySbd = {};
      Object.keys(groups).forEach(function(ma) {
        groups[ma].students.forEach(function(s) {
          if (s.sbd) (bySbd[s.sbd] = bySbd[s.sbd] || []).push(s);
        });
      });
      var existing = {};
      kqLoad(bai.maBai).forEach(function(r) {
        existing[r.sbd] = r;
      });
      var items = p.items || [];
      if (typeof items === "string") items = JSON.parse(items);
      if (p.auto) {
        items = [];
        pend.forEach(function(r) {
          var m = bySbd[r.sbd] || [];
          if (m.length === 1 && m[0].trangThai === "\u0110ang h\u1ECDc" && !existing[r.sbd] && r.maDe) items.push({ id: r.id, action: "assign", to: r.sbd });
        });
      }
      var byId = {};
      pend.forEach(function(r) {
        byId[r.id] = r;
      });
      var errors = [], delRows = [], write = [], addNhom = {}, maDeNew = {}, used = {}, resolved = 0, discarded = 0;
      items.forEach(function(it) {
        var r = byId[it.id];
        if (!r) return;
        if (it.action === "discard") {
          delRows.push(r._row);
          discarded++;
          logSua(u, bai.maBai, r.sbd, "b\u1ECF d\xF2ng ch\u1EDD", r.tong + " \u0111i\u1EC3m (m\xE3 " + r.maDe + ")", "(\u0111\xE3 b\u1ECF)", "b\u1ECF d\xF2ng ch\u1EDD x\u1EED l\xFD");
          return;
        }
        var to = normSbd(it.to);
        if (!/^\d{6}$/.test(to)) {
          errors.push({ id: r.id, msg: "SBD kh\xF4ng h\u1EE3p l\u1EC7" });
          return;
        }
        try {
          assertGroupAccessBySbd(u, to);
        } catch (e) {
          errors.push({ id: r.id, msg: e.message });
          return;
        }
        var m = bySbd[to] || [];
        if (m.length !== 1) {
          errors.push({ id: r.id, msg: m.length ? "SBD " + to + " \u0111ang d\xF9ng cho nhi\u1EC1u h\u1ECDc sinh" : "SBD " + to + " ch\u01B0a c\xF3 trong danh s\xE1ch l\u1EDBp" });
          return;
        }
        if (existing[to] || used[to]) {
          errors.push({ id: r.id, msg: (m[0].ho + " " + m[0].ten).trim() + " \u0111\xE3 c\xF3 k\u1EBFt qu\u1EA3 b\xE0i n\xE0y" });
          return;
        }
        if (!r.maDe) {
          errors.push({ id: r.id, msg: "D\xF2ng n\xE0y thi\u1EBFu m\xE3 \u0111\u1EC1" });
          return;
        }
        used[to] = true;
        write.push({ nguon: r.sbd && r.sbd !== to ? r.sbd : "", sbd: to, maDe: r.maDe, tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, tl: r.tl, daSua: false });
        addNhom[to.substring(0, 3)] = true;
        maDeNew[r.maDe] = true;
        delRows.push(r._row);
        resolved++;
        logSua(u, bai.maBai, to, "x\u1EED l\xFD d\xF2ng ch\u1EDD", 'd\xF2ng SBD "' + r.sbd + '" trong file', r.tong + " \u0111i\u1EC3m (m\xE3 " + r.maDe + ")", "gh\xE9p d\xF2ng ch\u1EDD x\u1EED l\xFD");
      });
      kqAppendRows(bai.maBai, write);
      var sh = getSystemSheet().getSheetByName("ChoXuLy");
      if (delRows.length) deleteRowsList(sh, delRows);
      var tt = write.length ? capNhatBaiSauKhiGhi(bai, cfg, maDeNew, addNhom) : bai.trangThai;
      sysTouch();
      if (tt === "\u0110\xE3 c\xF4ng b\u1ED1" && write.length) pubTouch();
      return { ok: true, resolved, discarded, errors, conLai: countChoXuLy(bai.maBai), trangThai: tt };
    } finally {
      lock.releaseLock();
    }
  }
  __name(giaiQuyetChoXuLy, "giaiQuyetChoXuLy");
  function chuanTen(s) {
    return String(s || "").normalize("NFC").toLowerCase().replace(/\s+/g, " ").trim();
  }
  __name(chuanTen, "chuanTen");
  function boDau(s) {
    return chuanTen(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d");
  }
  __name(boDau, "boDau");
  function laBaiNgoai(bai) {
    return parseJsonSafe(bai.cauHinh, {}).loai === "ngoai";
  }
  __name(laBaiNgoai, "laBaiNgoai");
  function chuanCfgNgoai(cfg) {
    var t = Number(cfg.thang) || 10;
    if (t < 1 || t > 100) throw new Error("Thang \u0111i\u1EC3m t\u1EEB 1 \u0111\u1EBFn 100");
    return { loai: "ngoai", thang: t, nguon: String(cfg.nguon || "").trim().slice(0, 60) };
  }
  __name(chuanCfgNgoai, "chuanCfgNgoai");
  function ngoaiBai(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    var cfg = parseJsonSafe(bai.cauHinh, {});
    if (cfg.loai !== "ngoai") throw new Error("B\xE0i n\xE0y l\xE0 b\xE0i tr\u1EAFc nghi\u1EC7m, nh\u1EADp \u0111i\u1EC3m b\u1EB1ng file TNMaker");
    return { bai, cfg, thang: Number(cfg.thang) || 10 };
  }
  __name(ngoaiBai, "ngoaiBai");
  function ngoaiHocSinh(bai) {
    var nhom = getNhomMap(), out = [];
    parseNhomList(bai.cacNhom).forEach(function(ma) {
      var reg = nhom[ma];
      if (!reg) return;
      var g = findTab(rosterSnapshot(reg.rosterSheetId, false), reg.sheetName, ma);
      if (!g) return;
      g.students.forEach(function(s) {
        if (s.sbd) out.push({ sbd: s.sbd, ten: ((s.ho || "") + " " + (s.ten || "")).trim(), lop: String(s.lop || ""), trangThai: s.trangThai, maNhom: ma });
      });
    });
    return out;
  }
  __name(ngoaiHocSinh, "ngoaiHocSinh");
  function ngoaiDiem(v, thang) {
    if (v === "" || v === null || v === void 0) return null;
    var n = Number(String(v).trim().replace(",", "."));
    if (!isFinite(n) || n < 0 || n > thang) return NaN;
    return Math.round(n * 100) / 100;
  }
  __name(ngoaiDiem, "ngoaiDiem");
  function ngoaiPick(h) {
    return { sbd: h.sbd, ten: h.ten, lop: h.lop };
  }
  __name(ngoaiPick, "ngoaiPick");
  function ngoaiDoiChieu(p, u) {
    var x = ngoaiBai(p, u), hs = ngoaiHocSinh(x.bai);
    var bySbd = {}, byTen = {}, byKd = {};
    hs.forEach(function(h) {
      bySbd[h.sbd] = h;
      (byTen[chuanTen(h.ten)] = byTen[chuanTen(h.ten)] || []).push(h);
      (byKd[boDau(h.ten)] = byKd[boDau(h.ten)] || []).push(h);
    });
    var rows = p.rows;
    if (typeof rows === "string") rows = JSON.parse(rows);
    if (!rows || !rows.length) throw new Error("Ch\u01B0a c\xF3 d\xF2ng \u0111i\u1EC3m n\xE0o");
    if (rows.length > 600) throw new Error("T\u1ED1i \u0111a 600 d\xF2ng m\u1ED7i l\u1EA7n");
    var hep = /* @__PURE__ */ __name(function(c, lop) {
      if (c.length > 1 && lop) {
        var cl = c.filter(function(h) {
          return chuanTen(h.lop) === chuanTen(lop);
        });
        if (cl.length) c = cl;
      }
      if (c.length > 1) {
        var dh = c.filter(function(h) {
          return h.trangThai === "\u0110ang h\u1ECDc";
        });
        if (dh.length) c = dh;
      }
      return c;
    }, "hep");
    var out = rows.map(function(r, i) {
      var o = { i, ten: String(r.ten || "").trim(), lop: String(r.lop || "").trim(), sbdFile: normSbd(r.sbd || ""), diemGoc: String(r.diem === void 0 ? "" : r.diem).trim(), ungVien: [] };
      o.diem = ngoaiDiem(r.diem, x.thang);
      if (o.diem === null || isNaN(o.diem)) {
        o.trangThai = "loi-diem";
        return o;
      }
      if (o.sbdFile && bySbd[o.sbdFile]) {
        o.trangThai = "khop";
        o.sbd = o.sbdFile;
        o.tenDs = bySbd[o.sbdFile].ten;
        return o;
      }
      var c = o.ten ? hep(byTen[chuanTen(o.ten)] || [], o.lop) : [];
      if (c.length === 1) {
        o.trangThai = "khop";
        o.sbd = c[0].sbd;
        o.tenDs = c[0].ten;
        return o;
      }
      if (c.length > 1) {
        o.trangThai = "trung-ten";
        o.ungVien = c.map(ngoaiPick);
        return o;
      }
      var g = o.ten ? hep(byKd[boDau(o.ten)] || [], o.lop) : [];
      if (g.length === 1) {
        o.trangThai = "gan-dung";
        o.sbd = g[0].sbd;
        o.tenDs = g[0].ten;
        o.ungVien = g.map(ngoaiPick);
        return o;
      }
      if (g.length > 1) {
        o.trangThai = "trung-ten";
        o.ungVien = g.map(ngoaiPick);
        return o;
      }
      var tenGoi = boDau(o.ten).split(" ").pop();
      o.ungVien = tenGoi ? hs.filter(function(h) {
        return boDau(h.ten).split(" ").pop() === tenGoi;
      }).slice(0, 8).map(ngoaiPick) : [];
      o.trangThai = "khong-thay";
      return o;
    });
    var dem = {};
    out.forEach(function(o) {
      if (o.sbd) dem[o.sbd] = (dem[o.sbd] || 0) + 1;
    });
    out.forEach(function(o) {
      if (o.sbd && dem[o.sbd] > 1) o.trungDong = true;
    });
    return { thang: x.thang, rows: out, hocSinh: hs.filter(function(h) {
      return h.trangThai === "\u0110ang h\u1ECDc";
    }).map(ngoaiPick), daCo: countKetQua(x.bai.maBai) };
  }
  __name(ngoaiDoiChieu, "ngoaiDoiChieu");
  function ngoaiLuu(p, u) {
    var x = ngoaiBai(p, u), by = {};
    ngoaiHocSinh(x.bai).forEach(function(h) {
      by[h.sbd] = h;
    });
    var gan = p.gan, cho = p.cho;
    if (typeof gan === "string") gan = JSON.parse(gan);
    if (typeof cho === "string") cho = JSON.parse(cho);
    gan = gan || [];
    cho = cho || [];
    var seen = {}, clean = [];
    gan.forEach(function(g) {
      var sbd = normSbd(g.sbd);
      if (!by[sbd]) throw new Error("SBD " + g.sbd + " kh\xF4ng thu\u1ED9c c\xE1c nh\xF3m c\u1EE7a b\xE0i");
      if (seen[sbd]) throw new Error("Em " + by[sbd].ten + " (" + sbd + ") \u0111ang c\xF3 2 d\xF2ng \u0111i\u1EC3m, h\xE3y b\u1ECF b\u1EDBt 1 d\xF2ng");
      seen[sbd] = 1;
      var d = ngoaiDiem(g.diem, x.thang);
      if (d === null || isNaN(d)) throw new Error("\u0110i\u1EC3m c\u1EE7a em " + sbd + " kh\xF4ng h\u1EE3p l\u1EC7 (0 \u0111\u1EBFn " + x.thang + ")");
      clean.push({ sbd, maDe: "", tong: d, p1: "", p2: "", p3: "", tl: {}, daSua: false, nguon: "ngoai" });
    });
    if (!clean.length && !cho.length) throw new Error("Ch\u01B0a c\xF3 d\xF2ng n\xE0o \u0111\u1EC3 l\u01B0u");
    var lock = LockService.getScriptLock();
    lock.waitLock(3e4);
    try {
      var thay = p.cheDo !== "gop", old = kqIndex(x.bai.maBai);
      if (thay) {
        deleteRowsList(_kq(), old.map(function(o) {
          return o.row;
        }));
        cxlDelete(x.bai.maBai, function() {
          return true;
        });
      } else deleteRowsList(_kq(), old.filter(function(o) {
        return seen[o.sbd];
      }).map(function(o) {
        return o.row;
      }));
      kqAppendRows(x.bai.maBai, clean);
      cxlAppend(x.bai.maBai, cho.map(function(c) {
        var d = ngoaiDiem(c.diem, x.thang);
        return {
          sbd: "",
          maDe: "",
          tong: d === null || isNaN(d) ? "" : d,
          p1: "",
          p2: "",
          p3: "",
          tl: null,
          lyDo: "Kh\xF4ng kh\u1EDBp t\xEAn trong danh s\xE1ch",
          ghiChu: "T\xEAn trong file: " + String(c.ten || "").slice(0, 80) + (c.lop ? " (l\u1EDBp " + String(c.lop).slice(0, 20) + ")" : "")
        };
      }), String(p.tenFile || "d\xE1n t\u1EEB b\u1EA3ng t\xEDnh").slice(0, 120));
      logSua(u, x.bai.maBai, "(c\u1EA3 b\xE0i)", "nh\u1EADp \u0111i\u1EC3m t\u1ED5ng", thay ? "thay to\xE0n b\u1ED9" : "g\u1ED9p th\xEAm", clean.length + " em, " + cho.length + " d\xF2ng ch\u1EDD", "nh\u1EADp \u0111i\u1EC3m t\u1EEB ngu\u1ED3n kh\xE1c");
      var tt = capNhatBaiSauKhiGhi(x.bai, x.cfg, {}, {});
      if (tt === "\u0110\xE3 c\xF4ng b\u1ED1") {
        try {
          pubTouch();
        } catch (e) {
        }
      }
      return { ok: true, daGhi: clean.length, cho: cho.length, trangThai: tt };
    } finally {
      lock.releaseLock();
    }
  }
  __name(ngoaiLuu, "ngoaiLuu");
  function ngoaiSuaDiem(p, u) {
    var x = ngoaiBai(p, u), sbd = normSbd(p.sbd), d = ngoaiDiem(p.diem, x.thang);
    if (d === null || isNaN(d)) throw new Error("\u0110i\u1EC3m ph\u1EA3i t\u1EEB 0 \u0111\u1EBFn " + x.thang);
    if (!String(p.lyDo || "").trim()) throw new Error("Ghi l\xFD do s\u1EEDa");
    var ix = kqIndex(x.bai.maBai).filter(function(o) {
      return o.sbd === sbd;
    })[0];
    if (!ix) throw new Error("Ch\u01B0a c\xF3 \u0111i\u1EC3m c\u1EE7a em " + sbd);
    var sh = _kq(), H = kqHeader(sh), cu = sh.getRange(ix.row, H["tong"] + 1).getValues()[0][0];
    sh.getRange(ix.row, H["tong"] + 1).setValue(d);
    sh.getRange(ix.row, H["daSua"] + 1).setValue("TRUE");
    sh.getRange(ix.row, H["capNhat"] + 1).setValue(nowIso());
    logSua(u, x.bai.maBai, sbd, "tong", cu, d, p.lyDo);
    if (x.bai.trangThai === "\u0110\xE3 c\xF4ng b\u1ED1") {
      try {
        pubTouch();
      } catch (e) {
      }
    }
    return { ok: true };
  }
  __name(ngoaiSuaDiem, "ngoaiSuaDiem");
  function ngoaiGiaiQuyet(p, u, bai) {
    var lock = LockService.getScriptLock();
    lock.waitLock(3e4);
    try {
      var pend = cxlLoad(bai.maBai), items = p.items || [];
      if (typeof items === "string") items = JSON.parse(items);
      var by = {};
      ngoaiHocSinh(bai).forEach(function(h) {
        by[h.sbd] = h;
      });
      var co = {};
      kqIndex(bai.maBai).forEach(function(o) {
        co[o.sbd] = 1;
      });
      var errors = [], add = [], del = {}, resolved = 0, discarded = 0;
      items.forEach(function(it) {
        var r = pend.filter(function(x) {
          return x.id === String(it.id);
        })[0];
        if (!r) {
          errors.push("Kh\xF4ng t\xECm th\u1EA5y d\xF2ng ch\u1EDD");
          return;
        }
        if (it.action === "discard") {
          del[r.id] = 1;
          discarded++;
          logSua(u, bai.maBai, "", "b\u1ECF d\xF2ng ch\u1EDD", r.ghiChu + ", " + r.tong + " \u0111i\u1EC3m", "(\u0111\xE3 b\u1ECF)", "b\u1ECF d\xF2ng ch\u1EDD x\u1EED l\xFD");
          return;
        }
        var to = normSbd(it.to);
        if (!by[to]) {
          errors.push("SBD " + it.to + " kh\xF4ng thu\u1ED9c nh\xF3m c\u1EE7a b\xE0i");
          return;
        }
        if (co[to]) {
          errors.push("Em " + by[to].ten + " (" + to + ") \u0111\xE3 c\xF3 \u0111i\u1EC3m");
          return;
        }
        if (isNaN(r.tong) || String(r.tong) === "") {
          errors.push('D\xF2ng "' + r.ghiChu + '" kh\xF4ng c\xF3 \u0111i\u1EC3m h\u1EE3p l\u1EC7');
          return;
        }
        co[to] = 1;
        del[r.id] = 1;
        resolved++;
        add.push({ sbd: to, maDe: "", tong: Number(r.tong), p1: "", p2: "", p3: "", tl: {}, daSua: false, nguon: "ngoai" });
      });
      kqAppendRows(bai.maBai, add);
      cxlDelete(bai.maBai, function(x) {
        return del[x.id];
      });
      if (bai.trangThai === "\u0110\xE3 c\xF4ng b\u1ED1") {
        try {
          pubTouch();
        } catch (e) {
        }
      }
      return { ok: true, resolved, discarded, errors, conLai: pend.length - resolved - discarded };
    } finally {
      lock.releaseLock();
    }
  }
  __name(ngoaiGiaiQuyet, "ngoaiGiaiQuyet");
  function pubVer() {
    return _props().getProperty("PUB_VER") || "0";
  }
  __name(pubVer, "pubVer");
  function pubTouch() {
    _props().setProperty("PUB_VER", Date.now().toString(36) + Math.random().toString(36).slice(2, 7));
  }
  __name(pubTouch, "pubTouch");
  function publicConfig() {
    var c = CacheService2.getScriptCache(), k = "pubcfg_" + pubVer();
    var hit = c.get(k);
    if (hit) return JSON.parse(hit);
    var out = {
      tenHienThi: String(getSetting("tenHienThi", "") || "Tra \u0111i\u1EC3m"),
      loiNhan: String(getSetting("loiNhanTraDiem", "") || ""),
      hienThiThuHang: toBool(getSetting("hienThiThuHang", false)),
      hienDiemTB: toBool(getSetting("hienDiemTB", false)),
      choPhepGopY: toBool(getSetting("choPhepGopY", true)),
      hienDiemCaoNhat: toBool(getSetting("hienDiemCaoNhat", true)),
      hienChuyenCan: toBool(getSetting("hienChuyenCan", false))
    };
    c.put(k, JSON.stringify(out), 3600);
    return out;
  }
  __name(publicConfig, "publicConfig");
  function maxDiemPhan(cfg) {
    return [
      cfg.p1 && cfg.p1.on ? Number(cfg.p1.diem) || 0 : 0,
      cfg.p2 && cfg.p2.on ? (Number(cfg.p2.soCau) || 0) * (Number(cfg.p2.d4) || 0) : 0,
      cfg.p3 && cfg.p3.on ? Number(cfg.p3.diem) || 0 : 0
    ];
  }
  __name(maxDiemPhan, "maxDiemPhan");
  function pubGroupIndex(ma) {
    var key = "pubg_" + pubVer() + "_" + ma;
    var hit = cacheGetBig(key);
    if (hit) return hit;
    var built = buildPubGroupIndex(ma);
    cachePutBig(key, built, 3600);
    return built;
  }
  __name(pubGroupIndex, "pubGroupIndex");
  function buildPubGroupIndex(ma) {
    var out = { tenNhom: "", hs: {}, kq: {}, bai: {} };
    var info = getNhomMap()[ma];
    if (!info) return out;
    out.tenNhom = String(info.tenNhom || "");
    try {
      var g = findTab(rosterSnapshot(info.rosterSheetId, false), info.sheetName, ma);
      if (g) {
        out.tenNhom = g.tenNhom;
        g.students.forEach(function(s) {
          if (s.sbd) out.hs[s.sbd] = { ho: s.ho, ten: s.ten, cu: s.sbdCu || "", tt: s.trangThai };
        });
      }
    } catch (e) {
    }
    var pub = {};
    sysRows("BaiKiemTra").forEach(function(b) {
      if (b.trangThai === "\u0110\xE3 c\xF4ng b\u1ED1") pub[String(b.maBai)] = b;
    });
    if (!Object.keys(pub).length) return out;
    var sh = _kq(), last = sh.getLastRow();
    if (last < 2) return out;
    var H = kqHeader(sh);
    var need = Math.max(H["maBai"], H["sbd"], H["maDe"], H["tong"], H["p1"], H["p2"], H["p3"]) + 1;
    var v = sh.getRange(2, 1, last - 1, need).getValues(), byBai = {};
    for (var i = 0; i < v.length; i++) {
      var mb = String(v[i][H["maBai"]]);
      if (!pub[mb]) continue;
      var sbd = normSbd(v[i][H["sbd"]]);
      if (!sbd || sbd.substring(0, 3) !== ma) continue;
      (byBai[mb] = byBai[mb] || []).push({
        sbd,
        m: normMaDe(v[i][H["maDe"]]),
        t: Number(v[i][H["tong"]]),
        p: [Number(v[i][H["p1"]]), Number(v[i][H["p2"]]), Number(v[i][H["p3"]])]
      });
    }
    Object.keys(byBai).forEach(function(mb2) {
      var list = byBai[mb2], b = pub[mb2], cfg = parseJsonSafe(b.cauHinh, {});
      var sum = 0, cao = 0;
      list.forEach(function(x) {
        sum += x.t;
        if (x.t > cao) cao = x.t;
      });
      out.bai[mb2] = {
        ten: String(b.tenBai),
        ngay: ngayStr(b.ngay),
        cb: tStr(b.ngayCongBo),
        n: list.length,
        tb: round2(sum / list.length),
        cao: round2(cao),
        mx: maxDiemPhan(cfg),
        loai: cfg.loai === "ngoai" ? "ngoai" : "",
        thang: Number(cfg.thang) || 10,
        nguon: cfg.nguon || ""
      };
      list.forEach(function(x) {
        var hang = 1;
        list.forEach(function(y) {
          if (y.t > x.t + 1e-9) hang++;
        });
        (out.kq[x.sbd] = out.kq[x.sbd] || []).push({ b: mb2, m: x.m, t: x.t, p: x.p, h: hang });
      });
    });
    return out;
  }
  __name(buildPubGroupIndex, "buildPubGroupIndex");
  function pubEntry(e, gi, cfgPub, viaSbd) {
    var b = gi.bai[e.b] || {};
    return {
      maBai: e.b,
      tenBai: b.ten,
      ngay: b.ngay,
      maDe: e.m,
      tong: e.t,
      p1: e.p[0],
      p2: e.p[1],
      p3: e.p[2],
      mx: b.mx || [0, 0, 0],
      loai: b.loai || "",
      thang: b.thang || 10,
      nguon: b.nguon || "",
      hang: cfgPub.hienThiThuHang ? e.h : null,
      soHs: cfgPub.hienThiThuHang ? b.n : null,
      tbNhom: cfgPub.hienDiemTB ? b.tb : null,
      caoNhat: cfgPub.hienDiemCaoNhat ? b.cao !== void 0 ? b.cao : null : null,
      sbdCu: viaSbd || null
    };
  }
  __name(pubEntry, "pubEntry");
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
  __name(sbdChain, "sbdChain");
  function traDiem(p) {
    var raw = String(p.sbd == null ? "" : p.sbd).replace(/\s/g, "");
    if (!/^\d{6}$/.test(raw)) throw new Error("S\u1ED1 b\xE1o danh g\u1ED3m \u0111\xFAng 6 ch\u1EEF s\u1ED1");
    var sbd = raw;
    var cfgPub = publicConfig();
    var gi = pubGroupIndex(sbd.substring(0, 3));
    var hs = gi.hs[sbd];
    var chain = sbdChain(sbd), list = [];
    chain.forEach(function(s, i) {
      var g = i === 0 ? gi : pubGroupIndex(s.substring(0, 3));
      (g.kq[s] || []).forEach(function(e) {
        list.push(pubEntry(e, g, cfgPub, i === 0 ? null : s));
      });
    });
    if (!hs && !list.length) return { found: false, sbd };
    list.sort(function(a, b) {
      return String(b.ngay || "") < String(a.ngay || "") ? -1 : String(b.ngay || "") > String(a.ngay || "") ? 1 : 0;
    });
    return {
      found: true,
      sbd,
      ho: hs ? hs.ho : "",
      ten: hs ? hs.ten : "",
      tenNhom: gi.tenNhom,
      daChuyen: hs ? hs.tt === "\u0110\xE3 chuy\u1EC3n" : false,
      hienThiThuHang: cfgPub.hienThiThuHang,
      hienDiemTB: cfgPub.hienDiemTB,
      hienDiemCaoNhat: cfgPub.hienDiemCaoNhat,
      chuyenCan: cfgPub.hienChuyenCan ? pubChuyenCan(sbd) : null,
      bai: list
    };
  }
  __name(traDiem, "traDiem");
  function chiTietBai(p) {
    var raw = String(p.sbd == null ? "" : p.sbd).replace(/\s/g, "");
    if (!/^\d{6}$/.test(raw)) throw new Error("S\u1ED1 b\xE1o danh g\u1ED3m \u0111\xFAng 6 ch\u1EEF s\u1ED1");
    var sbd = raw;
    var bai = findBai(p.maBai);
    if (!bai || bai.trangThai !== "\u0110\xE3 c\xF4ng b\u1ED1") throw new Error("B\xE0i n\xE0y ch\u01B0a \u0111\u01B0\u1EE3c c\xF4ng b\u1ED1");
    if (laBaiNgoai(bai)) throw new Error("B\xE0i n\xE0y ch\u1EC9 c\xF3 \u0111i\u1EC3m t\u1ED5ng, kh\xF4ng c\xF3 chi ti\u1EBFt t\u1EEBng c\xE2u");
    var chain = sbdChain(sbd);
    var r = kqLoad(bai.maBai).filter(function(x) {
      return chain.indexOf(x.sbd) !== -1;
    })[0];
    if (!r) throw new Error("Kh\xF4ng c\xF3 k\u1EBFt qu\u1EA3 c\u1EE7a s\u1ED1 b\xE1o danh n\xE0y trong b\xE0i");
    var cfg = parseJsonSafe(bai.cauHinh, {});
    var key = getBaiDapAn(bai.maBai)[r.maDe] || null;
    var sc = key && r.tl ? chamBai(cfg, r.tl, key) : null;
    return {
      tenBai: String(bai.tenBai),
      ngay: ngayStr(bai.ngay),
      maDe: r.maDe,
      tong: r.tong,
      p1: r.p1,
      p2: r.p2,
      p3: r.p3,
      mx: maxDiemPhan(cfg),
      cfg: {
        p1: cfg.p1 && cfg.p1.on ? { soCau: cfg.p1.soCau, diem: cfg.p1.diem } : null,
        p2: cfg.p2 && cfg.p2.on ? { soCau: cfg.p2.soCau, d: [cfg.p2.d1, cfg.p2.d2, cfg.p2.d3, cfg.p2.d4] } : null,
        p3: cfg.p3 && cfg.p3.on ? { soCau: cfg.p3.soCau, diem: cfg.p3.diem } : null
      },
      tl: r.tl || null,
      key,
      dung: sc ? sc.dung : null,
      daSua: r.daSua
    };
  }
  __name(chiTietBai, "chiTietBai");
  function guiGopY(p, u) {
    if (p.website) return { ok: true };
    var nd = String(p.noiDung || "").trim();
    if (nd.length < 5) throw new Error("N\u1ED9i dung g\xF3p \xFD qu\xE1 ng\u1EAFn (\xEDt nh\u1EA5t 5 k\xFD t\u1EF1)");
    if (nd.length > 2e3) nd = nd.slice(0, 2e3);
    if (!u && !toBool(getSetting("choPhepGopY", true))) throw new Error("Hi\u1EC7n ch\u01B0a nh\u1EADn g\xF3p \xFD qua trang n\xE0y");
    var c = CacheService2.getScriptCache(), n = Number(c.get("gopy_rate")) || 0;
    if (!u && n >= 30) throw new Error("\u0110ang c\xF3 qu\xE1 nhi\u1EC1u g\xF3p \xFD, h\xE3y th\u1EED l\u1EA1i sau \xEDt ph\xFAt");
    c.put("gopy_rate", String(n + 1), 600);
    var sh = getSystemSheet().getSheetByName("GopY");
    var H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), nCol = sh.getLastColumn();
    var row = [];
    for (var i = 0; i < nCol; i++) row.push("");
    var set = /* @__PURE__ */ __name(function(k, v) {
      if (H[k] !== void 0) row[H[k]] = v;
    }, "set");
    set("id", "G" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5));
    set("thoiGian", nowIso());
    set("noiDung", nd);
    set("lienHe", String(p.lienHe || "").trim().slice(0, 200));
    set("daXuLy", "");
    set("nguon", u ? "gi\xE1o vi\xEAn" : "trang tra \u0111i\u1EC3m");
    set("nguoiGui", u ? u.username : "");
    set("sbd", normSbd(p.sbd) || "");
    var r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, nCol).setNumberFormat("@");
    sh.getRange(r, 1, 1, nCol).setValues([row]);
    sysTouch("GopY");
    try {
      var to = String(getSetting("emailNhanThongBao", "") || "").trim();
      if (to) MailApp.sendEmail({
        to,
        subject: "[\u0110i\u1EC3m] G\xF3p \xFD m\u1EDBi" + (u ? " t\u1EEB " + u.username : ""),
        name: "H\u1EC7 th\u1ED1ng \u0111i\u1EC3m",
        body: nd + "\n\n" + (p.lienHe ? "Li\xEAn h\u1EC7: " + p.lienHe + "\n" : "") + (p.sbd ? "SBD: " + p.sbd + "\n" : "") + "Ngu\u1ED3n: " + (u ? "gi\xE1o vi\xEAn " + u.username : "trang tra \u0111i\u1EC3m")
      });
    } catch (e) {
    }
    return { ok: true };
  }
  __name(guiGopY, "guiGopY");
  function _median(arr) {
    if (!arr.length) return null;
    var a = arr.slice().sort(function(x, y) {
      return x - y;
    }), m = Math.floor(a.length / 2);
    return a.length % 2 ? a[m] : round2((a[m - 1] + a[m]) / 2);
  }
  __name(_median, "_median");
  function _tomTat(list) {
    var n = list.length, out = { n, tb: null, trungVi: null, cao: null, thap: null, p1: null, p2: null, p3: null, pho: [] };
    for (var k = 0; k <= 10; k++) out.pho.push(0);
    if (!n) return out;
    var t = list.map(function(r) {
      return r.tong;
    }), s = 0, s1 = 0, s2 = 0, s3 = 0;
    list.forEach(function(r) {
      s += r.tong;
      s1 += r.p1;
      s2 += r.p2;
      s3 += r.p3;
      var bin = Math.min(10, Math.max(0, Math.floor(r.tong + 1e-9)));
      out.pho[bin]++;
    });
    out.tb = round2(s / n);
    out.trungVi = _median(t);
    out.cao = Math.max.apply(null, t);
    out.thap = Math.min.apply(null, t);
    out.p1 = round2(s1 / n);
    out.p2 = round2(s2 / n);
    out.p3 = round2(s3 / n);
    return out;
  }
  __name(_tomTat, "_tomTat");
  function thongKeBai(p, u) {
    var bai = findBai(p.maBai);
    requireBaiAccess(u, bai);
    if (parseJsonSafe(bai.cauHinh, {}).loai === "ngoai") throw new Error("B\xE0i ch\u1EC9 c\xF3 \u0111i\u1EC3m t\u1ED5ng: xem ph\u1ED5 \u0111i\u1EC3m \u1EDF trang K\u1EBFt qu\u1EA3");
    return cachedResult("tk_" + bai.maBai + "_" + u.username, 900, function() {
      return thongKeBai_(p, u, bai);
    }, p.fresh);
  }
  __name(thongKeBai, "thongKeBai");
  function thongKeBai_(p, u, bai) {
    var cfg = parseJsonSafe(bai.cauHinh, {});
    var mx = maxDiemPhan(cfg);
    var res = kqLoad(bai.maBai), keys = getBaiDapAn(bai.maBai);
    var groups = loadAccessibleGroups(u, false);
    var nguong = Number(getSetting("nguongBoTuc", 5)) || 5;
    var byGroup = {};
    res.forEach(function(r) {
      var ma = r.sbd.substring(0, 3);
      (byGroup[ma] = byGroup[ma] || []).push(r);
    });
    var maList = parseNhomList(bai.cacNhom);
    Object.keys(byGroup).forEach(function(m) {
      if (maList.indexOf(m) === -1) maList.push(m);
    });
    maList.sort();
    var names = {};
    var nhomOut = maList.map(function(ma) {
      var g = groups[ma], list = byGroup[ma] || [];
      var dangHoc = g ? g.students.filter(function(s) {
        return s.trangThai === "\u0110ang h\u1ECDc";
      }) : [];
      if (g) g.students.forEach(function(s) {
        names[s.sbd] = (s.ho + " " + s.ten).trim();
      });
      var co = {};
      list.forEach(function(r) {
        co[r.sbd] = true;
      });
      var t = _tomTat(list);
      t.maNhom = ma;
      t.tenNhom = g ? g.tenNhom : ma;
      t.vang = dangHoc.filter(function(s) {
        return !co[s.sbd];
      }).length;
      return t;
    });
    var theoMa = {}, coChiTiet = 0;
    res.forEach(function(r) {
      var k = keys[r.maDe];
      if (!k || !r.tl) return;
      var sc = chamBai(cfg, r.tl, k);
      if (sc.thieuKey) return;
      coChiTiet++;
      var M = theoMa[r.maDe];
      if (!M) {
        M = theoMa[r.maDe] = { maDe: r.maDe, n: 0, p1: [], p2: [], p3: [] };
        (k.p1 || []).forEach(function(x, i) {
          if (cfg.p1 && cfg.p1.on && i < cfg.p1.soCau) M.p1.push({ cau: i + 1, dapAn: x, dung: 0, chon: { A: 0, B: 0, C: 0, D: 0, "": 0 } });
        });
        if (cfg.p2 && cfg.p2.on) for (var j = 0; j < cfg.p2.soCau; j++) M.p2.push({ cau: j + 1, dapAn: normP2((k.p2 || [])[j]), y: [0, 0, 0, 0], du4: 0 });
        if (cfg.p3 && cfg.p3.on) for (var q = 0; q < cfg.p3.soCau; q++) M.p3.push({ cau: q + 1, dapAn: normP3((k.p3 || [])[q]), dung: 0, sai: {} });
      }
      M.n++;
      M.p1.forEach(function(c, i) {
        var a = String((r.tl.p1 || [])[i] || "").toUpperCase();
        c.chon[/^[A-D]$/.test(a) ? a : ""]++;
        if (sc.dung.p1[i] === 1) c.dung++;
      });
      M.p2.forEach(function(c, i) {
        var a = normP2((r.tl.p2 || [])[i]);
        for (var y = 0; y < 4; y++) if (a.charAt(y) === c.dapAn.charAt(y)) c.y[y]++;
        if (sc.dung.p2[i] === 4) c.du4++;
      });
      M.p3.forEach(function(c, i) {
        if (sc.dung.p3[i] === 1) c.dung++;
        else {
          var a = normP3((r.tl.p3 || [])[i]) || "(b\u1ECF tr\u1ED1ng)";
          c.sai[a] = (c.sai[a] || 0) + 1;
        }
      });
    });
    var pct = /* @__PURE__ */ __name(function(x, n) {
      return n ? Math.round(x * 1e3 / n) / 10 : 0;
    }, "pct");
    var khoNhat = [];
    var maOut = Object.keys(theoMa).sort().map(function(md) {
      var M = theoMa[md];
      var o = {
        maDe: md,
        n: M.n,
        p1: M.p1.map(function(c) {
          var tl = pct(c.dung, M.n);
          khoNhat.push({ maDe: md, phan: "I", cau: c.cau, tiLe: tl });
          var nhieu = ["A", "B", "C", "D"].filter(function(x) {
            return x !== c.dapAn;
          }).sort(function(a, b) {
            return c.chon[b] - c.chon[a];
          })[0];
          return { cau: c.cau, dapAn: c.dapAn, tiLe: tl, chon: c.chon, nhieuNhat: c.chon[nhieu] ? nhieu : null };
        }),
        p2: M.p2.map(function(c) {
          var tl = pct(c.du4, M.n);
          khoNhat.push({ maDe: md, phan: "II", cau: c.cau, tiLe: tl });
          return { cau: c.cau, dapAn: c.dapAn, tiLeDu4: tl, y: c.y.map(function(x) {
            return pct(x, M.n);
          }) };
        }),
        p3: M.p3.map(function(c) {
          var tl = pct(c.dung, M.n);
          khoNhat.push({ maDe: md, phan: "III", cau: c.cau, tiLe: tl });
          var sai = Object.keys(c.sai).map(function(k) {
            return { gt: k, n: c.sai[k] };
          }).sort(function(a, b) {
            return b.n - a.n;
          }).slice(0, 3);
          return { cau: c.cau, dapAn: c.dapAn, tiLe: tl, saiHay: sai };
        })
      };
      return o;
    });
    khoNhat.sort(function(a, b) {
      return a.tiLe - b.tiLe;
    });
    var boTuc = [];
    res.forEach(function(r) {
      var ly = [];
      if (r.tong < nguong) ly.push("d\u01B0\u1EDBi " + nguong + " \u0111i\u1EC3m");
      ["p1", "p2", "p3"].forEach(function(k, i) {
        if (mx[i] > 0 && r[k] < mx[i] / 2) ly.push("Ph\u1EA7n " + ["I", "II", "III"][i] + " d\u01B0\u1EDBi m\u1ED9t n\u1EEDa");
      });
      if (ly.length) boTuc.push({ sbd: r.sbd, ten: names[r.sbd] || "", tong: r.tong, p1: r.p1, p2: r.p2, p3: r.p3, lyDo: ly });
    });
    boTuc.sort(function(a, b) {
      return a.tong - b.tong;
    });
    var cungNhom = sysRows("BaiKiemTra").filter(function(b) {
      if (u.role !== "admin" && b.teacher !== u.username) return false;
      var ms = parseNhomList(b.cacNhom);
      return ms.some(function(m) {
        return maList.indexOf(m) !== -1;
      });
    });
    cungNhom.sort(function(a, b) {
      var x = ngayStr(a.ngay) || String(a.createdAt), y = ngayStr(b.ngay) || String(b.createdAt);
      return x < y ? -1 : x > y ? 1 : 0;
    });
    var vt = cungNhom.map(function(b) {
      return String(b.maBai);
    }).indexOf(String(bai.maBai));
    var chon = cungNhom.slice(Math.max(0, vt - 5), vt + 1);
    var chonSet = {};
    chon.forEach(function(b) {
      chonSet[String(b.maBai)] = { maBai: String(b.maBai), tenBai: String(b.tenBai), ngay: ngayStr(b.ngay), nhom: {} };
    });
    var shk = _kq(), last = shk.getLastRow();
    if (last >= 2 && chon.length > 1) {
      var H = kqHeader(shk);
      var v = shk.getRange(2, 1, last - 1, Math.max(H["maBai"], H["sbd"], H["tong"]) + 1).getValues();
      v.forEach(function(row) {
        var o = chonSet[String(row[H["maBai"]])];
        if (!o) return;
        var ma = normSbd(row[H["sbd"]]).substring(0, 3);
        if (maList.indexOf(ma) === -1) return;
        var g = o.nhom[ma] = o.nhom[ma] || { s: 0, n: 0 };
        g.s += Number(row[H["tong"]]);
        g.n++;
      });
    } else if (chon.length === 1) {
      res.forEach(function(r) {
        var ma = r.sbd.substring(0, 3);
        var g = chonSet[String(bai.maBai)].nhom[ma] = chonSet[String(bai.maBai)].nhom[ma] || { s: 0, n: 0 };
        g.s += r.tong;
        g.n++;
      });
    }
    var xuHuong = chon.map(function(b) {
      var o = chonSet[String(b.maBai)], nh = {};
      Object.keys(o.nhom).forEach(function(m) {
        nh[m] = o.nhom[m].n ? round2(o.nhom[m].s / o.nhom[m].n) : null;
      });
      return { maBai: o.maBai, tenBai: o.tenBai, ngay: o.ngay, laBaiNay: o.maBai === String(bai.maBai), tbNhom: nh };
    }).filter(function(x) {
      return Object.keys(x.tbNhom).length;
    });
    return {
      bai: { maBai: bai.maBai, tenBai: String(bai.tenBai), ngay: ngayStr(bai.ngay), trangThai: bai.trangThai },
      mx,
      nguong,
      tongQuan: _tomTat(res),
      nhom: nhomOut,
      coChiTiet,
      theoMa: maOut,
      khoNhat: khoNhat.slice(0, 8),
      boTuc,
      xuHuong
    };
  }
  __name(thongKeBai_, "thongKeBai_");
  var DD_CODES = { C: "C\xF3 m\u1EB7t", M: "\u0110i mu\u1ED9n", P: "V\u1EAFng c\xF3 ph\xE9p", K: "V\u1EAFng kh\xF4ng ph\xE9p" };
  function homNay() {
    return Utilities.formatDate(/* @__PURE__ */ new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
  }
  __name(homNay, "homNay");
  function ddSheet() {
    return getSystemSheet().getSheetByName("BuoiHoc");
  }
  __name(ddSheet, "ddSheet");
  function ddRows(maNhom) {
    return sysRows("BuoiHoc").filter(function(r) {
      return !maNhom || normMa(r.maNhom) === normMa(maNhom);
    }).map(ddObj);
  }
  __name(ddRows, "ddRows");
  function ddObj(r) {
    return {
      _row: r._row,
      maBuoi: String(r.maBuoi),
      maNhom: normMa(r.maNhom),
      ngay: ngayStr(r.ngay),
      trangThai: String(r.trangThai || "\u0110ang m\u1EDF"),
      dd: parseJsonSafe(r.dd, {}) || {},
      ghiChu: parseJsonSafe(r.ghiChu, {}) || {},
      nguoiGhi: String(r.nguoiGhi || ""),
      capNhat: tStr(r.capNhat),
      nopBoi: String(r.nopBoi || ""),
      nopLuc: tStr(r.nopLuc),
      chotBoi: String(r.chotBoi || ""),
      chotLuc: tStr(r.chotLuc),
      taoBoi: String(r.taoBoi || "")
    };
  }
  __name(ddObj, "ddObj");
  function ddFind(maBuoi) {
    return ddRows().filter(function(b) {
      return b.maBuoi === String(maBuoi);
    })[0] || null;
  }
  __name(ddFind, "ddFind");
  function ddWrite(b) {
    var sh = ddSheet(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), n = sh.getLastColumn();
    var row = [];
    for (var i = 0; i < n; i++) row.push("");
    var set = /* @__PURE__ */ __name(function(k, v) {
      if (H[k] !== void 0) row[H[k]] = v;
    }, "set");
    set("maBuoi", b.maBuoi);
    set("maNhom", b.maNhom);
    set("ngay", b.ngay);
    set("trangThai", b.trangThai);
    set("dd", JSON.stringify(b.dd || {}));
    set("ghiChu", JSON.stringify(b.ghiChu || {}));
    set("nguoiGhi", b.nguoiGhi || "");
    set("capNhat", nowIso());
    set("nopBoi", b.nopBoi || "");
    set("nopLuc", b.nopLuc || "");
    set("chotBoi", b.chotBoi || "");
    set("chotLuc", b.chotLuc || "");
    set("taoBoi", b.taoBoi || "");
    set("taoLuc", b.taoLuc || nowIso());
    var r = b._row || sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, n).setNumberFormat("@");
    sh.getRange(r, 1, 1, n).setValues([row]);
    b._row = r;
    sysTouch("BuoiHoc");
    try {
      pubTouch();
    } catch (e) {
    }
    return b;
  }
  __name(ddWrite, "ddWrite");
  function ddLog(entries) {
    if (!entries.length) return;
    var sh = getSystemSheet().getSheetByName("DiemDanhLog"), n = 8, r0 = sh.getLastRow() + 1;
    sh.getRange(r0, 1, entries.length, n).setNumberFormat("@");
    sh.getRange(r0, 1, entries.length, n).setValues(entries.map(function(e) {
      return [nowIso(), e.maBuoi, e.sbd, e.cu || "", e.moi || "", e.nguoi, e.loai || "ghi", e.lyDo || ""];
    }));
    sysTouch("DiemDanhLog");
  }
  __name(ddLog, "ddLog");
  function ltInfo(u) {
    var r = findUserRow(u.username);
    if (!r) throw new Error("Kh\xF4ng t\xECm th\u1EA5y t\xE0i kho\u1EA3n");
    var ma = normMa(userField(r, "maNhom"));
    if (!ma) throw new Error("T\xE0i kho\u1EA3n l\u1EDBp tr\u01B0\u1EDFng ch\u01B0a \u0111\u01B0\u1EE3c g\xE1n nh\xF3m");
    return { maNhom: ma, quanLy: String(userField(r, "quanLy") || "") };
  }
  __name(ltInfo, "ltInfo");
  function isStaff(u) {
    return u.role === "admin" || u.role === "teacher";
  }
  __name(isStaff, "isStaff");
  function ddAccess(u, maNhom) {
    maNhom = normMa(maNhom);
    if (u.role === "loptruong") {
      if (ltInfo(u).maNhom !== maNhom) throw new Error("Kh\xF4ng c\xF3 quy\u1EC1n v\u1EDBi nh\xF3m n\xE0y");
      return { laLopTruong: true };
    }
    requireGroupAccess(u, maNhom);
    return { laLopTruong: false };
  }
  __name(ddAccess, "ddAccess");
  function ddStudents(maNhom, fresh) {
    var reg = getNhomMap()[normMa(maNhom)];
    if (!reg) throw new Error("Nh\xF3m " + maNhom + " ch\u01B0a \u0111\u0103ng k\xFD");
    var g = findTab(rosterSnapshot(reg.rosterSheetId, !!fresh), reg.sheetName, normMa(maNhom));
    if (!g) throw new Error("Kh\xF4ng \u0111\u1ECDc \u0111\u01B0\u1EE3c danh s\xE1ch nh\xF3m " + maNhom);
    return { tenNhom: g.tenNhom, students: g.students.filter(function(s) {
      return s.sbd && s.trangThai === "\u0110ang h\u1ECDc";
    }).map(function(s) {
      return { sbd: s.sbd, ho: s.ho, ten: s.ten };
    }) };
  }
  __name(ddStudents, "ddStudents");
  function ddCanEdit(u, b) {
    if (isStaff(u)) return { ok: true, sauChot: b.trangThai === "\u0110\xE3 ch\u1ED1t" || b.ngay < homNay() };
    if (b.ngay !== homNay()) return { ok: false, lyDo: "Ch\u1EC9 \u0111i\u1EC3m danh \u0111\u01B0\u1EE3c bu\u1ED5i c\u1EE7a h\xF4m nay. Mu\u1ED1n s\u1EEDa bu\u1ED5i c\u0169, h\xE3y g\u1EEDi y\xEAu c\u1EA7u s\u1EEDa cho c\xF4." };
    if (b.trangThai === "\u0110\xE3 ch\u1ED1t") return { ok: false, lyDo: "C\xF4 \u0111\xE3 ch\u1ED1t bu\u1ED5i n\xE0y. Mu\u1ED1n s\u1EEDa, h\xE3y g\u1EEDi y\xEAu c\u1EA7u s\u1EEDa cho c\xF4." };
    return { ok: true, sauChot: false };
  }
  __name(ddCanEdit, "ddCanEdit");
  function ddMoBuoi(p, u) {
    var ma = u.role === "loptruong" ? ltInfo(u).maNhom : normMa(p.maNhom);
    ddAccess(u, ma);
    var ngay = u.role === "loptruong" ? homNay() : String(p.ngay || "") || homNay();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ngay)) throw new Error("Ng\xE0y kh\xF4ng h\u1EE3p l\u1EC7");
    if (ngay > homNay()) throw new Error("Kh\xF4ng \u0111i\u1EC3m danh tr\u01B0\u1EDBc cho ng\xE0y ch\u01B0a t\u1EDBi");
    var lock = LockService.getScriptLock();
    lock.waitLock(2e4);
    try {
      var id = p.maBuoi ? String(p.maBuoi) : ma + "-" + ngay;
      var b = ddFind(id);
      if (!b) {
        if (p.maBuoi && !isStaff(u)) throw new Error("Kh\xF4ng t\xECm th\u1EA5y bu\u1ED5i h\u1ECDc");
        b = ddWrite({ maBuoi: id, maNhom: ma, ngay, trangThai: "\u0110ang m\u1EDF", dd: {}, ghiChu: {}, taoBoi: u.username, taoLuc: nowIso() });
      } else {
        ddAccess(u, b.maNhom);
        _KHONG_DOI = true;
      }
      return ddView(b, u, false);
    } finally {
      lock.releaseLock();
    }
  }
  __name(ddMoBuoi, "ddMoBuoi");
  function ddView(b, u, withLog) {
    var st = ddStudents(b.maNhom, false);
    var edit = ddCanEdit(u, b);
    var out = {
      buoi: { maBuoi: b.maBuoi, maNhom: b.maNhom, ngay: b.ngay, trangThai: b.trangThai, nguoiGhi: b.nguoiGhi, capNhat: b.capNhat, nopBoi: b.nopBoi, nopLuc: b.nopLuc, chotBoi: b.chotBoi, chotLuc: b.chotLuc, homNay: b.ngay === homNay() },
      tenNhom: st.tenNhom,
      students: st.students,
      dd: b.dd,
      ghiChu: b.ghiChu,
      suaDuoc: edit.ok,
      lyDoKhoa: edit.lyDo || "",
      laLopTruong: u.role === "loptruong"
    };
    if (withLog && isStaff(u)) out.nhatKy = sysRows("DiemDanhLog").filter(function(r) {
      return String(r.maBuoi) === b.maBuoi;
    }).reverse().slice(0, 300).map(function(r) {
      return { luc: tStr(r.thoiGian), sbd: normSbd(r.sbd) || String(r.sbd), cu: String(r.cu), moi: String(r.moi), nguoi: String(r.nguoi), loai: String(r.loai), lyDo: String(r.lyDo || "") };
    });
    return out;
  }
  __name(ddView, "ddView");
  function ddXemBuoi(p, u) {
    var b = ddFind(p.maBuoi);
    if (!b) throw new Error("Kh\xF4ng t\xECm th\u1EA5y bu\u1ED5i h\u1ECDc");
    ddAccess(u, b.maNhom);
    return ddView(b, u, true);
  }
  __name(ddXemBuoi, "ddXemBuoi");
  function ddLuu(p, u) {
    var lock = LockService.getScriptLock();
    lock.waitLock(2e4);
    try {
      var b = ddFind(p.maBuoi);
      if (!b) throw new Error("Kh\xF4ng t\xECm th\u1EA5y bu\u1ED5i h\u1ECDc");
      ddAccess(u, b.maNhom);
      var edit = ddCanEdit(u, b);
      if (!edit.ok) throw new Error(edit.lyDo);
      var lyDo = String(p.lyDo || "").trim();
      if (edit.sauChot && !lyDo) throw new Error("Bu\u1ED5i n\xE0y \u0111\xE3 kh\xF3a: c\u1EA7n ghi l\xFD do khi s\u1EEDa");
      var hop = {};
      ddStudents(b.maNhom, false).students.forEach(function(s) {
        hop[s.sbd] = true;
      });
      var ghi = p.ghi;
      if (typeof ghi === "string") ghi = JSON.parse(ghi);
      var logs = [], loai = edit.sauChot ? "s\u1EEDa sau ch\u1ED1t" : b.trangThai === "\u0110\xE3 n\u1ED9p" && u.role === "loptruong" ? "s\u1EEDa sau khi g\u1EEDi" : "ghi";
      (ghi || []).forEach(function(g) {
        var sbd = normSbd(g.sbd), tt = String(g.tt || "").toUpperCase();
        if (!hop[sbd] && !b.dd[sbd]) return;
        if (!DD_CODES[tt]) return;
        var cu = b.dd[sbd] || "C";
        if (g.ghiChu !== void 0) {
          var nt = String(g.ghiChu || "").trim().slice(0, 120);
          if (nt) b.ghiChu[sbd] = nt;
          else delete b.ghiChu[sbd];
        }
        if (cu === tt && b.dd[sbd]) return;
        b.dd[sbd] = tt;
        if (cu !== tt) logs.push({ maBuoi: b.maBuoi, sbd, cu, moi: tt, nguoi: u.username, loai, lyDo });
      });
      Object.keys(hop).forEach(function(s) {
        if (!b.dd[s]) b.dd[s] = "C";
      });
      b.nguoiGhi = u.username;
      ddWrite(b);
      ddLog(logs);
      return { ok: true, doi: logs.length, trangThai: b.trangThai, capNhat: nowIso() };
    } finally {
      lock.releaseLock();
    }
  }
  __name(ddLuu, "ddLuu");
  function ddNop(p, u) {
    var b = ddFind(p.maBuoi);
    if (!b) throw new Error("Kh\xF4ng t\xECm th\u1EA5y bu\u1ED5i h\u1ECDc");
    ddAccess(u, b.maNhom);
    var edit = ddCanEdit(u, b);
    if (!edit.ok) throw new Error(edit.lyDo);
    if (b.trangThai !== "\u0110\xE3 ch\u1ED1t") b.trangThai = "\u0110\xE3 n\u1ED9p";
    b.nopBoi = u.username;
    b.nopLuc = nowIso();
    ddWrite(b);
    ddLog([{ maBuoi: b.maBuoi, sbd: "(bu\u1ED5i)", cu: "", moi: "\u0111\xE3 g\u1EEDi", nguoi: u.username, loai: "g\u1EEDi" }]);
    return { ok: true, trangThai: b.trangThai };
  }
  __name(ddNop, "ddNop");
  function ddChot(p, u) {
    if (!isStaff(u)) throw new Error("Kh\xF4ng c\xF3 quy\u1EC1n");
    var b = ddFind(p.maBuoi);
    if (!b) throw new Error("Kh\xF4ng t\xECm th\u1EA5y bu\u1ED5i h\u1ECDc");
    ddAccess(u, b.maNhom);
    b.trangThai = p.chot ? "\u0110\xE3 ch\u1ED1t" : b.nopBoi ? "\u0110\xE3 n\u1ED9p" : "\u0110ang m\u1EDF";
    b.chotBoi = p.chot ? u.username : "";
    b.chotLuc = p.chot ? nowIso() : "";
    ddWrite(b);
    ddLog([{ maBuoi: b.maBuoi, sbd: "(bu\u1ED5i)", cu: "", moi: p.chot ? "ch\u1ED1t" : "m\u1EDF l\u1EA1i", nguoi: u.username, loai: p.chot ? "ch\u1ED1t" : "m\u1EDF l\u1EA1i" }]);
    return { ok: true, trangThai: b.trangThai };
  }
  __name(ddChot, "ddChot");
  function ddXoaBuoi(p, u) {
    if (!isStaff(u)) throw new Error("Kh\xF4ng c\xF3 quy\u1EC1n");
    var b = ddFind(p.maBuoi);
    if (!b) throw new Error("Kh\xF4ng t\xECm th\u1EA5y bu\u1ED5i h\u1ECDc");
    ddAccess(u, b.maNhom);
    if (!String(p.lyDo || "").trim()) throw new Error("Ghi l\xFD do x\xF3a bu\u1ED5i");
    deleteRowsList(ddSheet(), [b._row]);
    ddLog([{ maBuoi: b.maBuoi, sbd: "(bu\u1ED5i)", cu: JSON.stringify(b.dd).slice(0, 200), moi: "(\u0111\xE3 x\xF3a)", nguoi: u.username, loai: "x\xF3a bu\u1ED5i", lyDo: p.lyDo }]);
    try {
      pubTouch();
    } catch (e) {
    }
    return { ok: true };
  }
  __name(ddXoaBuoi, "ddXoaBuoi");
  function ddYcSheet() {
    return getSystemSheet().getSheetByName("YeuCauSua");
  }
  __name(ddYcSheet, "ddYcSheet");
  function ddYeuCau(p, u) {
    if (!u && p.website) return { ok: true };
    var b = ddFind(p.maBuoi);
    if (!b) throw new Error("Kh\xF4ng t\xECm th\u1EA5y bu\u1ED5i h\u1ECDc");
    var sbd = normSbd(p.sbd), moi = String(p.moi || "").toUpperCase(), lyDo = String(p.lyDo || "").trim();
    if (!DD_CODES[moi]) throw new Error("Ch\u1ECDn tr\u1EA1ng th\xE1i mu\u1ED1n s\u1EEDa th\xE0nh");
    if (lyDo.length < 5) throw new Error("Ghi r\xF5 l\xFD do (\xEDt nh\u1EA5t 5 k\xFD t\u1EF1)");
    if (u) ddAccess(u, b.maNhom);
    else {
      if (!toBool(getSetting("hienChuyenCan", false))) throw new Error("Ch\u01B0a m\u1EDF t\xEDnh n\u0103ng n\xE0y");
      var c = CacheService2.getScriptCache(), n = Number(c.get("ddyc_rate")) || 0;
      if (n >= 30) throw new Error("\u0110ang c\xF3 qu\xE1 nhi\u1EC1u y\xEAu c\u1EA7u, th\u1EED l\u1EA1i sau \xEDt ph\xFAt");
      c.put("ddyc_rate", String(n + 1), 600);
    }
    if (!b.dd[sbd] && sbd.substring(0, 3) !== b.maNhom) throw new Error("S\u1ED1 b\xE1o danh kh\xF4ng thu\u1ED9c bu\u1ED5i h\u1ECDc n\xE0y");
    var cu = b.dd[sbd] || "C";
    if (cu === moi) throw new Error('Tr\u1EA1ng th\xE1i hi\u1EC7n t\u1EA1i \u0111\xE3 l\xE0 "' + DD_CODES[moi] + '"');
    var sh = ddYcSheet(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), nC = sh.getLastColumn();
    var row = [];
    for (var i = 0; i < nC; i++) row.push("");
    var set = /* @__PURE__ */ __name(function(k, v) {
      if (H[k] !== void 0) row[H[k]] = v;
    }, "set");
    set("id", "Y" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5));
    set("thoiGian", nowIso());
    set("maBuoi", b.maBuoi);
    set("maNhom", b.maNhom);
    set("sbd", sbd);
    set("cu", cu);
    set("moi", moi);
    set("lyDo", lyDo.slice(0, 500));
    set("nguoiGui", u ? u.username : "ph\u1EE5 huynh/h\u1ECDc sinh");
    set("lienHe", String(p.lienHe || "").slice(0, 120));
    set("trangThai", "Ch\u1EDD duy\u1EC7t");
    var r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, nC).setNumberFormat("@");
    sh.getRange(r, 1, 1, nC).setValues([row]);
    sysTouch("YeuCauSua");
    return { ok: true };
  }
  __name(ddYeuCau, "ddYeuCau");
  function ddYeuCauList(u, maNhom) {
    var nhom = getNhomMap();
    return sysRows("YeuCauSua").filter(function(r) {
      var ma = normMa(r.maNhom), info = nhom[ma];
      if (maNhom && ma !== normMa(maNhom)) return false;
      return u.role === "admin" || info && info.teacher === u.username;
    }).map(function(r) {
      return {
        id: String(r.id),
        luc: tStr(r.thoiGian),
        maBuoi: String(r.maBuoi),
        maNhom: normMa(r.maNhom),
        sbd: normSbd(r.sbd),
        cu: String(r.cu),
        moi: String(r.moi),
        lyDo: String(r.lyDo || ""),
        nguoiGui: String(r.nguoiGui || ""),
        lienHe: String(r.lienHe || ""),
        trangThai: String(r.trangThai || ""),
        duyetBoi: String(r.duyetBoi || ""),
        _row: r._row
      };
    }).reverse();
  }
  __name(ddYeuCauList, "ddYeuCauList");
  function ddDuyet(p, u) {
    if (!isStaff(u)) throw new Error("Kh\xF4ng c\xF3 quy\u1EC1n");
    var lock = LockService.getScriptLock();
    lock.waitLock(2e4);
    try {
      var y = ddYeuCauList(u).filter(function(x) {
        return x.id === String(p.id);
      })[0];
      if (!y) throw new Error("Kh\xF4ng t\xECm th\u1EA5y y\xEAu c\u1EA7u");
      if (y.trangThai !== "Ch\u1EDD duy\u1EC7t") throw new Error("Y\xEAu c\u1EA7u n\xE0y \u0111\xE3 \u0111\u01B0\u1EE3c x\u1EED l\xFD");
      var sh = ddYcSheet(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
      if (p.dongY) {
        var b = ddFind(y.maBuoi);
        if (!b) throw new Error("Bu\u1ED5i h\u1ECDc \u0111\xE3 b\u1ECB x\xF3a");
        var cu = b.dd[y.sbd] || "C";
        b.dd[y.sbd] = y.moi;
        ddWrite(b);
        ddLog([{ maBuoi: b.maBuoi, sbd: y.sbd, cu, moi: y.moi, nguoi: u.username, loai: "duy\u1EC7t y\xEAu c\u1EA7u", lyDo: y.lyDo + " (g\u1EEDi b\u1EDFi " + y.nguoiGui + ")" }]);
      }
      setTextCell(sh, y._row, H["trangThai"] + 1, p.dongY ? "\u0110\xE3 duy\u1EC7t" : "T\u1EEB ch\u1ED1i");
      setTextCell(sh, y._row, H["duyetBoi"] + 1, u.username);
      setTextCell(sh, y._row, H["duyetLuc"] + 1, nowIso());
      if (H["phanHoi"] !== void 0) setTextCell(sh, y._row, H["phanHoi"] + 1, String(p.phanHoi || ""));
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  }
  __name(ddDuyet, "ddDuyet");
  function ddDem(dd, list) {
    var c = { C: 0, M: 0, P: 0, K: 0 };
    list.forEach(function(s) {
      c[dd[s] || "C"]++;
    });
    return c;
  }
  __name(ddDem, "ddDem");
  function ddTongQuan(p, u) {
    if (u.role === "loptruong") {
      var me = ltInfo(u), st = ddStudents(me.maNhom, false), rows = ddRows(me.maNhom);
      rows.sort(function(a, b) {
        return a.ngay < b.ngay ? 1 : -1;
      });
      return {
        laLopTruong: true,
        maNhom: me.maNhom,
        tenNhom: st.tenNhom,
        homNay: homNay(),
        ganDay: rows.slice(0, 12).map(function(b) {
          var c = ddDem(b.dd, st.students.map(function(s) {
            return s.sbd;
          }));
          return { maBuoi: b.maBuoi, ngay: b.ngay, trangThai: b.trangThai, dem: c };
        })
      };
    }
    if (!isStaff(u)) throw new Error("Kh\xF4ng c\xF3 quy\u1EC1n");
    return cachedResult("ddtq_" + u.username + "_" + homNay(), 900, function() {
      return ddTongQuan_(u);
    }, p.fresh);
  }
  __name(ddTongQuan, "ddTongQuan");
  function ddTongQuan_(u) {
    var groups = listGroups(u).groups, today = homNay(), rowsAll = ddRows();
    var yc = ddYeuCauList(u).filter(function(y) {
      return y.trangThai === "Ch\u1EDD duy\u1EC7t";
    });
    return {
      homNay: today,
      nhom: groups.map(function(g) {
        var b = rowsAll.filter(function(x) {
          return x.maNhom === g.maNhom && x.ngay === today;
        })[0];
        return {
          maNhom: g.maNhom,
          tenNhom: g.tenNhom,
          siSo: g.active,
          homNay: b ? { maBuoi: b.maBuoi, trangThai: b.trangThai, nopBoi: b.nopBoi, nopLuc: b.nopLuc, dem: ddDem(b.dd, Object.keys(b.dd)) } : null,
          soBuoi: rowsAll.filter(function(x) {
            return x.maNhom === g.maNhom;
          }).length,
          choDuyet: yc.filter(function(y) {
            return y.maNhom === g.maNhom;
          }).length
        };
      }),
      choDuyet: yc.length
    };
  }
  __name(ddTongQuan_, "ddTongQuan_");
  function ddPhanTich(p, u) {
    if (!isStaff(u)) throw new Error("Kh\xF4ng c\xF3 quy\u1EC1n");
    var ma = normMa(p.maNhom);
    ddAccess(u, ma);
    return cachedResult("ddpt_" + ma + "_" + u.username + "_" + homNay(), 900, function() {
      return ddPhanTich_(ma, u);
    }, p.fresh);
  }
  __name(ddPhanTich, "ddPhanTich");
  function ddPhanTich_(ma, u) {
    var nguong = Number(getSetting("nguongChuyenCan", 80)) || 80;
    var st = ddStudents(ma, false), list = st.students;
    var rows = ddRows(ma).sort(function(a, b) {
      return a.ngay < b.ngay ? -1 : 1;
    });
    var buoi = rows.map(function(b) {
      var c = ddDem(b.dd, list.map(function(s) {
        return s.sbd;
      })), n = list.length || 1;
      return { maBuoi: b.maBuoi, ngay: b.ngay, trangThai: b.trangThai, dem: c, tiLe: Math.round((c.C + c.M) * 1e3 / n) / 10, nopBoi: b.nopBoi, nguoiGhi: b.nguoiGhi };
    });
    var hs = list.map(function(s) {
      var c = { C: 0, M: 0, P: 0, K: 0 }, soBuoi = 0;
      rows.forEach(function(b) {
        if (b.dd[s.sbd]) {
          c[b.dd[s.sbd]]++;
          soBuoi++;
        }
      });
      var tiLe = soBuoi ? Math.round((c.C + c.M) * 1e3 / soBuoi) / 10 : null;
      return {
        sbd: s.sbd,
        ten: (s.ho + " " + s.ten).trim(),
        soBuoi,
        dem: c,
        tiLe,
        duoiNguong: tiLe !== null && tiLe < nguong,
        gan: rows.slice(-12).map(function(b) {
          return b.dd[s.sbd] || "";
        })
      };
    });
    hs.sort(function(a, b) {
      return (a.tiLe === null ? 101 : a.tiLe) - (b.tiLe === null ? 101 : b.tiLe);
    });
    var co = [];
    var baiTheoNgay = {};
    sysRows("BaiKiemTra").forEach(function(bk) {
      if (parseNhomList(bk.cacNhom).indexOf(ma) === -1) return;
      var ng = ngayStr(bk.ngay);
      if (ng) (baiTheoNgay[ng] = baiTheoNgay[ng] || []).push(bk);
    });
    var tenCua = {};
    list.forEach(function(s) {
      tenCua[s.sbd] = (s.ho + " " + s.ten).trim();
    });
    rows.forEach(function(b) {
      (baiTheoNgay[b.ngay] || []).forEach(function(bk) {
        var coBai = {};
        kqIndex(bk.maBai).forEach(function(r) {
          if (r.sbd.substring(0, 3) === ma) coBai[r.sbd] = true;
        });
        var soCo = Object.keys(coBai).length;
        if (!soCo) return;
        list.forEach(function(s) {
          var tt = b.dd[s.sbd] || "C";
          if ((tt === "P" || tt === "K") && coBai[s.sbd]) co.push({ loai: "vang-co-bai", ngay: b.ngay, maBuoi: b.maBuoi, sbd: s.sbd, ten: tenCua[s.sbd], msg: '\u0110i\u1EC3m danh v\u1EAFng nh\u01B0ng c\xF3 b\xE0i l\xE0m "' + bk.tenBai + '"' });
          if ((tt === "C" || tt === "M") && !coBai[s.sbd] && soCo >= list.length * 0.5) co.push({ loai: "comat-khong-bai", ngay: b.ngay, maBuoi: b.maBuoi, sbd: s.sbd, ten: tenCua[s.sbd], msg: '\u0110i\u1EC3m danh c\xF3 m\u1EB7t nh\u01B0ng kh\xF4ng c\xF3 b\xE0i l\xE0m "' + bk.tenBai + '"' });
        });
      });
    });
    var setBuoi = {};
    rows.forEach(function(b) {
      setBuoi[b.maBuoi] = b;
    });
    sysRows("DiemDanhLog").forEach(function(l) {
      if (!setBuoi[String(l.maBuoi)] || String(l.loai) !== "s\u1EEDa sau khi g\u1EEDi") return;
      co.push({
        loai: "sua-sau-gui",
        ngay: setBuoi[String(l.maBuoi)].ngay,
        maBuoi: String(l.maBuoi),
        sbd: normSbd(l.sbd),
        ten: tenCua[normSbd(l.sbd)] || "",
        msg: String(l.nguoi) + " s\u1EEDa " + (DD_CODES[String(l.cu)] || l.cu) + " \u2192 " + (DD_CODES[String(l.moi)] || l.moi) + " sau khi \u0111\xE3 g\u1EEDi \u0111i\u1EC3m danh (" + tStr(l.thoiGian) + ")"
      });
    });
    co.sort(function(a, b) {
      return a.ngay < b.ngay ? 1 : -1;
    });
    var tong = { C: 0, M: 0, P: 0, K: 0 };
    buoi.forEach(function(b) {
      ["C", "M", "P", "K"].forEach(function(k) {
        tong[k] += b.dem[k];
      });
    });
    var tongLuot = tong.C + tong.M + tong.P + tong.K;
    return {
      maNhom: ma,
      tenNhom: st.tenNhom,
      nguong,
      siSo: list.length,
      tongQuan: { soBuoi: rows.length, tiLe: tongLuot ? Math.round((tong.C + tong.M) * 1e3 / tongLuot) / 10 : null, dem: tong, duoiNguong: hs.filter(function(h) {
        return h.duoiNguong;
      }).length },
      buoi: buoi.slice(-30),
      ngayGan: rows.slice(-12).map(function(b) {
        return b.ngay;
      }),
      hocSinh: hs,
      coCanhBao: co.slice(0, 60),
      yeuCau: ddYeuCauList(u, ma).filter(function(y) {
        return y.trangThai === "Ch\u1EDD duy\u1EC7t";
      }),
      lopTruong: ltList(u, ma)
    };
  }
  __name(ddPhanTich_, "ddPhanTich_");
  function ltList(u, maNhom) {
    return sysRows("NguoiDung").filter(function(r) {
      if (r.role !== "loptruong") return false;
      if (maNhom && normMa(r.maNhom) !== normMa(maNhom)) return false;
      return u.role === "admin" || String(r.quanLy) === u.username;
    }).map(function(r) {
      return { username: String(r.username), ten: String(r.fullName), maNhom: normMa(r.maNhom), active: toBool(r.active), mustChange: toBool(r.mustChange) };
    });
  }
  __name(ltList, "ltList");
  function ltTao(p, u) {
    if (!isStaff(u)) throw new Error("Kh\xF4ng c\xF3 quy\u1EC1n");
    var ma = normMa(p.maNhom);
    ddAccess(u, ma);
    var sbd = normSbd(p.sbd);
    var hs = ddStudents(ma, true).students.filter(function(s) {
      return s.sbd === sbd;
    })[0];
    if (!hs) throw new Error("Kh\xF4ng t\xECm th\u1EA5y h\u1ECDc sinh " + p.sbd + " trong nh\xF3m \u0111ang h\u1ECDc");
    var username = "lt" + sbd;
    if (findUserRow(username)) throw new Error("Em n\xE0y \u0111\xE3 c\xF3 t\xE0i kho\u1EA3n l\u1EDBp tr\u01B0\u1EDFng: " + username);
    var pw = String(Math.floor(1e5 + Math.random() * 9e5));
    var sh = _users(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), n = sh.getLastColumn();
    var row = [];
    for (var i = 0; i < n; i++) row.push("");
    var salt = randomSalt();
    var set = /* @__PURE__ */ __name(function(k, v) {
      if (H[k] !== void 0) row[H[k]] = v;
    }, "set");
    set("username", username);
    set("role", "loptruong");
    set("fullName", (hs.ho + " " + hs.ten).trim());
    set("passHash", hashPassword(pw, salt));
    set("salt", salt);
    set("mustChange", true);
    set("active", true);
    set("createdAt", nowIso());
    set("maNhom", ma);
    set("quanLy", u.role === "admin" ? (getNhomMap()[ma] || {}).teacher || u.username : u.username);
    sh.appendRow(row);
    var r = sh.getLastRow();
    if (H["maNhom"] !== void 0) setTextCell(sh, r, H["maNhom"] + 1, ma);
    sysTouch();
    return { ok: true, username, matKhau: pw };
  }
  __name(ltTao, "ltTao");
  function ltSua(p, u) {
    if (!isStaff(u)) throw new Error("Kh\xF4ng c\xF3 quy\u1EC1n");
    var lt = ltList(u).filter(function(x) {
      return x.username === String(p.username);
    })[0];
    if (!lt) throw new Error("Kh\xF4ng t\xECm th\u1EA5y t\xE0i kho\u1EA3n l\u1EDBp tr\u01B0\u1EDFng");
    if (p.xoa) {
      var r = findUserRow(lt.username);
      deleteRowsList(r.sheet, [r.row]);
      revokeSessions(lt.username);
      return { ok: true };
    }
    if (p.datLai) {
      var pw = String(Math.floor(1e5 + Math.random() * 9e5));
      var ur = findUserRow(lt.username);
      setPassword(ur, pw, true);
      revokeSessions(lt.username);
      return { ok: true, matKhau: pw };
    }
    return setUserActive({ username: lt.username, active: !!p.active }, u);
  }
  __name(ltSua, "ltSua");
  function pubChuyenCan(sbd) {
    var ma = sbd.substring(0, 3);
    var key = "pubdd_" + pubVer() + "_" + ma, idxDd = cacheGetBig(key);
    if (!idxDd) {
      idxDd = ddRows(ma).sort(function(a, b) {
        return a.ngay < b.ngay ? -1 : 1;
      }).map(function(b) {
        return { m: b.maBuoi, n: b.ngay, d: b.dd };
      });
      cachePutBig(key, idxDd, 3600);
    }
    var c = { C: 0, M: 0, P: 0, K: 0 }, gan = [];
    idxDd.forEach(function(b) {
      var t = b.d[sbd];
      if (!t) return;
      c[t]++;
      gan.push({ maBuoi: b.m, ngay: b.n, tt: t });
    });
    var so = c.C + c.M + c.P + c.K;
    return { soBuoi: so, dem: c, tiLe: so ? Math.round((c.C + c.M) * 1e3 / so) / 10 : null, gan: gan.slice(-15).reverse() };
  }
  __name(pubChuyenCan, "pubChuyenCan");
  function authorize() {
    MailApp.getRemainingDailyQuota();
    DriveApp.getRootFolder();
    Logger.log("\u0110\xE3 c\u1EA5p quy\u1EC1n xong. Gi\u1EDD c\xF3 th\u1EC3 Deploy phi\xEAn b\u1EA3n m\u1EDBi.");
  }
  __name(authorize, "authorize");
  function getSetting(key, def) {
    var rows = sysRows("CaiDat");
    for (var i = 0; i < rows.length; i++) if (String(rows[i].key) === key) return rows[i].value;
    return def;
  }
  __name(getSetting, "getSetting");
  function countFilesIn(folder) {
    var n = 0, it = folder.getFiles();
    while (it.hasNext()) {
      it.next();
      n++;
    }
    return n;
  }
  __name(countFilesIn, "countFilesIn");
  function baiMeta(maBai, fresh) {
    var cache = CacheService2.getScriptCache(), key = "baimeta_" + maBai;
    if (!fresh) {
      var c = cache.get(key);
      if (c) return JSON.parse(c);
    }
    var bai = findBai(maBai);
    if (!bai) throw new Error("Kh\xF4ng t\xECm th\u1EA5y b\xE0i");
    var lock = LockService.getScriptLock();
    lock.waitLock(2e4);
    try {
      var anh = getAnhFolder(bai.driveFolderId);
      var meta = { teacher: bai.teacher, driveFolderId: bai.driveFolderId, anhFolderId: anh.getId() };
      cache.put(key, JSON.stringify(meta), 21600);
      return meta;
    } catch (e) {
      throw new Error("Kh\xF4ng m\u1EDF \u0111\u01B0\u1EE3c th\u01B0 m\u1EE5c c\u1EE7a b\xE0i tr\xEAn Drive (c\xF3 th\u1EC3 \u0111\xE3 b\u1ECB x\xF3a ho\u1EB7c \u0111\u1ED5i quy\u1EC1n): " + (e.message || e));
    } finally {
      lock.releaseLock();
    }
  }
  __name(baiMeta, "baiMeta");
  function assertMetaAccess(meta, u) {
    if (u.role !== "admin" && meta.teacher !== u.username) throw new Error("Kh\xF4ng c\xF3 quy\u1EC1n v\u1EDBi b\xE0i n\xE0y");
  }
  __name(assertMetaAccess, "assertMetaAccess");
  function anhInfo(p, u) {
    var meta = baiMeta(p.maBai);
    assertMetaAccess(meta, u);
    var folder;
    try {
      folder = DriveApp.getFolderById(meta.anhFolderId);
      if (folder.isTrashed()) throw new Error("trashed");
    } catch (e) {
      meta = baiMeta(p.maBai, true);
      folder = DriveApp.getFolderById(meta.anhFolderId);
    }
    return { count: countFilesIn(folder), folderUrl: "https://drive.google.com/drive/folders/" + meta.anhFolderId };
  }
  __name(anhInfo, "anhInfo");
  function uploadAnh(p, u) {
    var meta = baiMeta(p.maBai);
    assertMetaAccess(meta, u);
    var imgs = p.images || [{ name: p.name, data: p.data }];
    if (!imgs.length) throw new Error("Kh\xF4ng c\xF3 \u1EA3nh \u0111\u1EC3 t\u1EA3i");
    var folder;
    try {
      folder = DriveApp.getFolderById(meta.anhFolderId);
    } catch (e) {
      CacheService2.getScriptCache().remove("baimeta_" + p.maBai);
      throw new Error("Th\u01B0 m\u1EE5c \u1EA3nh kh\xF4ng c\xF2n. H\xE3y th\u1EED l\u1EA1i, h\u1EC7 th\u1ED1ng s\u1EBD t\u1EA1o l\u1EA1i.");
    }
    var added = 0;
    imgs.forEach(function(im) {
      var data = String(im.data || "");
      if (!data) throw new Error("\u1EA2nh r\u1ED7ng");
      if (data.length > 12 * 1024 * 1024) throw new Error("\u1EA2nh qu\xE1 l\u1EDBn (tr\xEAn 9 MB)");
      var name = String(im.name || "anh-" + Date.now()).replace(/[^A-Za-z0-9._-]/g, "_");
      if (!/\.jpe?g$/i.test(name)) name += ".jpg";
      folder.createFile(Utilities.newBlob(Utilities.base64Decode(data), "image/jpeg", name));
      added++;
    });
    return { ok: true, added };
  }
  __name(uploadAnh, "uploadAnh");
  function clearAnh(p, u) {
    var meta = baiMeta(p.maBai);
    assertMetaAccess(meta, u);
    var it = DriveApp.getFolderById(meta.anhFolderId).getFiles(), n = 0;
    while (it.hasNext()) {
      it.next().setTrashed(true);
      n++;
    }
    _props().deleteProperty("ANHMAIL_" + p.maBai);
    return { ok: true, removed: n };
  }
  __name(clearAnh, "clearAnh");
  function finishAnh(p, u) {
    var meta = baiMeta(p.maBai);
    assertMetaAccess(meta, u);
    return thongBaoAnh(p.maBai);
  }
  __name(finishAnh, "finishAnh");
  function thongBaoAnh(maBai) {
    var out = { soAnh: 0, hasDapAn: false, emailSent: false };
    try {
      var bai = findBai(maBai);
      if (!bai) return out;
      try {
        var folder = DriveApp.getFolderById(bai.driveFolderId);
        out.hasDapAn = folder.getFilesByName("dap-an - " + safeName(bai.tenBai) + ".xlsx").hasNext();
        out.soAnh = countFilesIn(getAnhFolder(bai.driveFolderId));
      } catch (e) {
        return out;
      }
      var props = _props(), key = "ANHMAIL_" + maBai;
      var last = Number(props.getProperty(key)) || 0;
      if (out.soAnh > last) {
        out.emailSent = notifyAnh(bai, out.soAnh, out.hasDapAn);
        if (out.emailSent) props.setProperty(key, String(out.soAnh));
      }
    } catch (e) {
    }
    return out;
  }
  __name(thongBaoAnh, "thongBaoAnh");
  function notifyAnh(bai, soAnh, hasDapAn) {
    try {
      var to = String(getSetting("emailNhanThongBao", "") || "").trim();
      if (!to) return false;
      var cfg = {};
      try {
        cfg = JSON.parse(bai.cauHinh || "{}");
      } catch (e) {
      }
      var body = "C\xF3 \u1EA3nh phi\u1EBFu b\xE0i l\xE0m m\u1EDBi, c\xF3 th\u1EC3 ch\u1EA5m b\u1EB1ng TNMaker.\n\nB\xE0i: " + bai.tenBai + "\nNg\xE0y: " + (ngayStr(bai.ngay) || "(ch\u01B0a ghi)") + "\nGi\xE1o vi\xEAn: " + bai.teacher + "\nNh\xF3m: " + (parseNhomList(bai.cacNhom).join(", ") || "(ch\u01B0a ch\u1ECDn)") + "\nM\xE3 \u0111\u1EC1: " + (cfg.maDeList || []).join(", ") + "\nS\u1ED1 \u1EA3nh phi\u1EBFu: " + soAnh + "\nFile \u0111\xE1p \xE1n tr\xEAn h\u1EC7 th\u1ED1ng: " + (hasDapAn ? "C\xF3" : "Ch\u01B0a c\xF3") + "\n\nTh\u01B0 m\u1EE5c c\u1EE7a b\xE0i tr\xEAn Drive: https://drive.google.com/drive/folders/" + bai.driveFolderId + "\n";
      MailApp.sendEmail({ to, subject: "[\u0110i\u1EC3m] C\xF3 \u1EA3nh phi\u1EBFu m\u1EDBi: " + bai.tenBai, body, name: "H\u1EC7 th\u1ED1ng \u0111i\u1EC3m" });
      return true;
    } catch (e) {
      return false;
    }
  }
  __name(notifyAnh, "notifyAnh");
  function cleanupOldImages() {
    var days = Number(getSetting("soNgayGiuAnh", 30)) || 30;
    var now = Date.now();
    sysRows("BaiKiemTra").forEach(function(r) {
      if (!r.ngayCongBo) return;
      var t = new Date(r.ngayCongBo).getTime();
      if (isNaN(t) || now - t <= days * 864e5) return;
      try {
        var it = DriveApp.getFolderById(r.driveFolderId).getFoldersByName("anh");
        while (it.hasNext()) it.next().setTrashed(true);
        CacheService2.getScriptCache().remove("baimeta_" + r.maBai);
      } catch (e) {
      }
    });
  }
  __name(cleanupOldImages, "cleanupOldImages");
  function installTriggers() {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      var f = t.getHandlerFunction();
      if (f === "cleanupOldImages" || f === "donDepHangNgay") ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger("donDepHangNgay").timeBased().everyDays(1).atHour(2).create();
    Logger.log("\u0110\xE3 c\xE0i l\u1ECBch d\u1ECDn d\u1EB9p h\u1EB1ng ng\xE0y (2h s\xE1ng).");
  }
  __name(installTriggers, "installTriggers");
  var CAI_DAT_MAC_DINH = {
    tenHienThi: "V\u1EADt L\xFD C\xF4 L\xE2m C\u01B0\u1EDDng",
    loiNhanTraDiem: "Nh\u1EADp s\u1ED1 b\xE1o danh \u0111\u1EC3 xem \u0111i\u1EC3m, t\u1EEBng c\xE2u \u0111\xFAng sai v\xE0 chuy\xEAn c\u1EA7n c\u1EE7a con. C\xF3 th\u1EAFc m\u1EAFc, vui l\xF2ng li\xEAn h\u1EC7 tr\u1EF1c ti\u1EBFp c\xF4.",
    hienThiThuHang: "false",
    hienDiemTB: "false",
    hienDiemCaoNhat: "true",
    hienChuyenCan: "false",
    nguongChuyenCan: "80",
    choPhepGopY: "true",
    nguongBoTuc: "5",
    soNgayGiuAnh: "30",
    emailNhanThongBao: ""
  };
  var CAI_DAT_BOOL = { hienThiThuHang: 1, hienDiemTB: 1, choPhepGopY: 1, hienDiemCaoNhat: 1, hienChuyenCan: 1 };
  var CAI_DAT_SO = { nguongBoTuc: [0, 10], soNgayGiuAnh: [1, 365], nguongChuyenCan: [0, 100] };
  function ensureSchema() {
    var c = CacheService2.getScriptCache();
    if (c.get("schema_" + APP_VERSION)) return;
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(1e4);
    } catch (e) {
      return;
    }
    try {
      if (c.get("schema_" + APP_VERSION)) return;
      var ss = getSystemSheet();
      Object.keys(SYS_TABS2).forEach(function(name) {
        var sh = ss.getSheetByName(name);
        if (!sh) sh = ss.insertSheet(name);
        if (sh.getLastRow() === 0) sh.appendRow(SYS_TABS2[name]);
      });
      ensureSysColumns(ss);
      var have = {};
      sysTouch("CaiDat");
      sysRows("CaiDat").forEach(function(r2) {
        have[String(r2.key)] = true;
      });
      var add = [];
      Object.keys(CAI_DAT_MAC_DINH).forEach(function(k) {
        if (have[k]) return;
        var v = CAI_DAT_MAC_DINH[k];
        if (k === "emailNhanThongBao") {
          try {
            v = Session.getEffectiveUser().getEmail() || "";
          } catch (e) {
          }
        }
        add.push([k, v]);
      });
      var CU = {
        tenHienThi: ["L\u1EDBp V\u1EADt l\xFD m\u1EB9 C\u01B0\u1EDDng", "Tra \u0111i\u1EC3m l\u1EDBp h\u1ECDc th\xEAm", "Tra \u0111i\u1EC3m", "V\u1EADt l\xFD c\xF4 C\u01B0\u1EDDng"],
        loiNhanTraDiem: ["\u0110i\u1EC3m c\xE1c b\xE0i ki\u1EC3m tra tr\xEAn l\u1EDBp h\u1ECDc th\xEAm. Th\u1EAFc m\u1EAFc v\u1EC1 \u0111i\u1EC3m, h\xE3y h\u1ECFi tr\u1EF1c ti\u1EBFp gi\xE1o vi\xEAn."]
      };
      var cdSh = ss.getSheetByName("CaiDat");
      sysRows("CaiDat").forEach(function(r2) {
        var k = String(r2.key);
        if (CU[k] && CU[k].indexOf(String(r2.value)) !== -1) setTextCell(cdSh, r2._row, 2, CAI_DAT_MAC_DINH[k]);
      });
      var meRow = findUserRow("me");
      if (meRow && String(userField(meRow, "fullName")) === "M\u1EB9 (gi\xE1o vi\xEAn V\u1EADt l\xFD)") meRow.sheet.getRange(meRow.row, meRow.H["fullName"] + 1).setValue("C\xF4 L\xE2m C\u01B0\u1EDDng");
      if (add.length) {
        var cd = ss.getSheetByName("CaiDat"), r = cd.getLastRow() + 1;
        cd.getRange(r, 1, add.length, 2).setNumberFormat("@");
        cd.getRange(r, 1, add.length, 2).setValues(add);
      }
      sysTouch();
      c.put("schema_" + APP_VERSION, "1", 21600);
    } finally {
      lock.releaseLock();
    }
  }
  __name(ensureSchema, "ensureSchema");
  function getCaiDatAll() {
    var out = {};
    Object.keys(CAI_DAT_MAC_DINH).forEach(function(k) {
      var v = getSetting(k, CAI_DAT_MAC_DINH[k]);
      out[k] = CAI_DAT_BOOL[k] ? toBool(v) : CAI_DAT_SO[k] ? Number(v) : String(v === null || v === void 0 ? "" : v);
    });
    return out;
  }
  __name(getCaiDatAll, "getCaiDatAll");
  function setCaiDat(p) {
    var vals = p.caiDat || {};
    if (typeof vals === "string") vals = JSON.parse(vals);
    var sh = getSystemSheet().getSheetByName("CaiDat");
    var rows = sysRows("CaiDat"), rowOf = {};
    rows.forEach(function(r) {
      rowOf[String(r.key)] = r._row;
    });
    Object.keys(vals).forEach(function(k) {
      if (!CAI_DAT_MAC_DINH.hasOwnProperty(k)) return;
      var v = vals[k];
      if (CAI_DAT_BOOL[k]) v = toBool(v) ? "true" : "false";
      else if (CAI_DAT_SO[k]) {
        var n = Number(v);
        if (!isFinite(n) || n < CAI_DAT_SO[k][0] || n > CAI_DAT_SO[k][1]) throw new Error('Gi\xE1 tr\u1ECB "' + k + '" ph\u1EA3i t\u1EEB ' + CAI_DAT_SO[k][0] + " \u0111\u1EBFn " + CAI_DAT_SO[k][1]);
        v = String(n);
      } else {
        v = String(v || "").trim().slice(0, 500);
        if (k === "emailNhanThongBao" && v && !/^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(v)) throw new Error("Email nh\u1EADn th\xF4ng b\xE1o kh\xF4ng h\u1EE3p l\u1EC7");
      }
      if (rowOf[k]) setTextCell(sh, rowOf[k], 2, v);
      else {
        var r = sh.getLastRow() + 1;
        sh.getRange(r, 1, 1, 2).setNumberFormat("@");
        sh.getRange(r, 1, 1, 2).setValues([[k, v]]);
        rowOf[k] = r;
      }
    });
    sysTouch();
    pubTouch();
    return getCaiDatAll();
  }
  __name(setCaiDat, "setCaiDat");
  function listGopY() {
    var rows = sysRows("GopY").map(function(r) {
      return {
        id: String(r.id || "row" + r._row),
        _row: r._row,
        thoiGian: tStr(r.thoiGian),
        noiDung: String(r.noiDung || ""),
        lienHe: String(r.lienHe || ""),
        daXuLy: toBool(r.daXuLy),
        nguon: String(r.nguon || ""),
        nguoiGui: String(r.nguoiGui || ""),
        sbd: normSbd(r.sbd) || ""
      };
    });
    rows.reverse();
    return { gopY: rows, moi: rows.filter(function(r) {
      return !r.daXuLy;
    }).length };
  }
  __name(listGopY, "listGopY");
  function xuLyGopY(p) {
    var sh = getSystemSheet().getSheetByName("GopY");
    var H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
    var r = sysRows("GopY").filter(function(x) {
      return String(x.id || "row" + x._row) === String(p.id);
    })[0];
    if (!r) throw new Error("Kh\xF4ng t\xECm th\u1EA5y g\xF3p \xFD (c\xF3 th\u1EC3 \u0111\xE3 b\u1ECB x\xF3a)");
    if (p.xoa) {
      deleteRowsList(sh, [r._row]);
      return { ok: true };
    }
    sh.getRange(r._row, H["daXuLy"] + 1).setValue(p.daXuLy ? "TRUE" : "");
    sysTouch("GopY");
    return { ok: true };
  }
  __name(xuLyGopY, "xuLyGopY");
  function donDepHangNgay() {
    try {
      cleanupOldImages();
    } catch (e) {
    }
    try {
      donPhien();
    } catch (e) {
    }
    try {
      catNhatKyLoi();
    } catch (e) {
    }
  }
  __name(donDepHangNgay, "donDepHangNgay");
  function donPhien() {
    var sh = getSystemSheet().getSheetByName("Phien");
    var rows = sysRows("Phien"), now = Date.now(), del = [];
    rows.forEach(function(r) {
      var t = new Date(r.expiry).getTime();
      if (isNaN(t) || t < now) del.push(r._row);
    });
    if (del.length) deleteRowsList(sh, del);
  }
  __name(donPhien, "donPhien");
  function trangThaiTuDong() {
    var on = ScriptApp.getProjectTriggers().some(function(t) {
      var f = t.getHandlerFunction();
      return f === "donDepHangNgay" || f === "cleanupOldImages";
    });
    return { daBat: on };
  }
  __name(trangThaiTuDong, "trangThaiTuDong");
  function batTuDong() {
    installTriggers();
    return trangThaiTuDong();
  }
  __name(batTuDong, "batTuDong");
  function ghiLoi(action, p, err) {
    try {
      var sh = getSystemSheet().getSheetByName("NhatKyLoi");
      if (!sh) return;
      var r = sh.getLastRow() + 1;
      sh.getRange(r, 1, 1, 4).setNumberFormat("@");
      sh.getRange(r, 1, 1, 4).setValues([[
        nowIso(),
        String(action),
        String(p && p.username || ""),
        String(err && err.stack || err).slice(0, 1500)
      ]]);
    } catch (e) {
    }
  }
  __name(ghiLoi, "ghiLoi");
  function catNhatKyLoi() {
    var sh = getSystemSheet().getSheetByName("NhatKyLoi");
    if (!sh) return;
    var n = sh.getLastRow() - 1;
    if (n > 500) sh.deleteRows(2, n - 300);
  }
  __name(catNhatKyLoi, "catNhatKyLoi");
  function nhatKyLoi() {
    var rows = sysRows("NhatKyLoi").slice(-100).reverse();
    return { loi: rows.map(function(r) {
      return { thoiGian: tStr(r.thoiGian), action: String(r.action), user: String(r.user), message: String(r.message) };
    }) };
  }
  __name(nhatKyLoi, "nhatKyLoi");
  function bootstrap(p, u) {
    if (u.role === "loptruong") {
      var lt = ltInfo(u), reg = getNhomMap()[lt.maNhom] || {};
      return { user: u, version: APP_VERSION, laLopTruong: true, maNhom: lt.maNhom, tenNhom: String(reg.tenNhom || lt.maNhom), caiDat: publicConfig(), groups: [], bai: [], gopYMoi: 0, homNay: homNay() };
    }
    var out = { user: u, version: APP_VERSION, groups: [], bai: [], caiDat: publicConfig(), gopYMoi: 0 };
    try {
      out.groups = listGroups(u).groups;
    } catch (e) {
      out.loiNhom = String(e.message || e);
    }
    out.bai = listBai(u, {}).bai;
    if (u.role === "admin") {
      out.caiDat = getCaiDatAll();
      try {
        out.gopYMoi = listGopY().moi;
      } catch (e) {
      }
    }
    try {
      out.ddChoDuyet = ddYeuCauList(u).filter(function(y) {
        return y.trangThai === "Ch\u1EDD duy\u1EC7t";
      }).length;
    } catch (e) {
      out.ddChoDuyet = 0;
    }
    out.homNay = homNay();
    return out;
  }
  __name(bootstrap, "bootstrap");
  function bridgeHandle(p) {
    if (!bridgeAuthOk(p)) return { ok: false, error: "Kh\xF3a c\u1EA7u n\u1ED1i kh\xF4ng \u0111\xFAng" };
    try {
      var fn = BRIDGE_OPS[String(p.op || "")];
      if (!fn) throw new Error("Kh\xF4ng c\xF3 thao t\xE1c c\u1EA7u n\u1ED1i: " + p.op);
      return { ok: true, data: fn(p) };
    } catch (e) {
      return { ok: false, error: String(e && e.message || e) };
    }
  }
  __name(bridgeHandle, "bridgeHandle");
  function bridgeAuthOk(p) {
    var tok = String(p.bridge || "");
    if (!tok) return false;
    var secret = String(_props().getProperty("BRIDGE_SECRET") || "");
    if (secret && tok === secret) return true;
    if (tok.indexOf("nonce:") !== 0) return false;
    var base = String(_props().getProperty("WORKER_URL") || (typeof WORKER_URL !== "undefined" ? WORKER_URL : "") || "").replace(/\/+$/, "");
    if (!/^https:\/\//.test(base)) return false;
    try {
      var res = UrlFetchApp.fetch(base + "/?bridgeVerify=" + encodeURIComponent(tok.slice(6)), { muteHttpExceptions: true, followRedirects: true });
      var j = JSON.parse(res.getContentText());
      return !!(j && j.ok && j.valid);
    } catch (e) {
      return false;
    }
  }
  __name(bridgeAuthOk, "bridgeAuthOk");
  var BRIDGE_OPS = {
    ping: /* @__PURE__ */ __name(function() {
      return { ok: true, luc: nowIso() };
    }, "ping"),
    /* ---- đọc toàn bộ file Sheet hệ thống (để chuyển dữ liệu sang máy chủ mới) ---- */
    sysRead: /* @__PURE__ */ __name(function() {
      var ss = getSystemSheet();
      return { sheetId: ss.getId(), tabs: ss.getSheets().map(function(sh) {
        var lr = sh.getLastRow(), lc = sh.getLastColumn();
        return { name: sh.getName(), values: lr && lc ? sh.getRange(1, 1, lr, lc).getDisplayValues() : [] };
      }) };
    }, "sysRead"),
    /* ---- danh sách lớp: đọc cả file, ghi đè từng tab ---- */
    rosterRead: /* @__PURE__ */ __name(function(p) {
      var ss = SpreadsheetApp.openById(String(p.fileId));
      return { tabs: ss.getSheets().map(function(sh) {
        var lr = sh.getLastRow(), lc = sh.getLastColumn();
        return { name: sh.getName(), values: lr && lc ? sh.getRange(1, 1, lr, lc).getDisplayValues() : [] };
      }) };
    }, "rosterRead"),
    rosterWrite: /* @__PURE__ */ __name(function(p) {
      var ss = SpreadsheetApp.openById(String(p.fileId)), tabs = p.tabs || [];
      if (typeof tabs === "string") tabs = JSON.parse(tabs);
      tabs.forEach(function(t) {
        var sh = ss.getSheetByName(t.name) || ss.insertSheet(t.name);
        var vals = (t.values || []).map(function(r) {
          return r.map(function(v) {
            return v === null || v === void 0 ? "" : v;
          });
        });
        var nc = vals.reduce(function(m, r) {
          return Math.max(m, r.length);
        }, 1);
        vals = vals.map(function(r) {
          while (r.length < nc) r.push("");
          return r;
        });
        var lr = sh.getLastRow(), lc = sh.getLastColumn();
        if (lr && lc) sh.getRange(1, 1, lr, lc).clearContent();
        if (vals.length) {
          sh.getRange(1, 1, vals.length, nc).setNumberFormat("@");
          sh.getRange(1, 1, vals.length, nc).setValues(vals);
        }
      });
      return { ok: true, tabs: tabs.length };
    }, "rosterWrite"),
    /* ---- Drive ---- */
    ensureBaiFolder: /* @__PURE__ */ __name(function(p) {
      if (p.folderId) {
        try {
          var f = DriveApp.getFolderById(String(p.folderId));
          if (!f.isTrashed()) return { folderId: f.getId() };
        } catch (e) {
        }
      }
      var folder = createBaiFolder(String(p.teacher || "me"), String(p.tenBai || "B\xE0i"), String(p.ngay || ""));
      return { folderId: folder.getId() };
    }, "ensureBaiFolder"),
    anhInfo: /* @__PURE__ */ __name(function(p) {
      var anh = getAnhFolder(String(p.folderId));
      return { count: countFilesIn(anh), folderUrl: "https://drive.google.com/drive/folders/" + anh.getId() };
    }, "anhInfo"),
    uploadAnh: /* @__PURE__ */ __name(function(p) {
      var anh = getAnhFolder(String(p.folderId)), imgs = p.images || [];
      if (typeof imgs === "string") imgs = JSON.parse(imgs);
      var added = 0;
      imgs.forEach(function(im) {
        var data = String(im.data || "");
        if (!data) throw new Error("\u1EA2nh r\u1ED7ng");
        if (data.length > 12 * 1024 * 1024) throw new Error("\u1EA2nh qu\xE1 l\u1EDBn (tr\xEAn 9 MB)");
        var name = String(im.name || "anh-" + Date.now()).replace(/[^A-Za-z0-9._-]/g, "_");
        if (!/\.jpe?g$/i.test(name)) name += ".jpg";
        anh.createFile(Utilities.newBlob(Utilities.base64Decode(data), "image/jpeg", name));
        added++;
      });
      return { added, count: countFilesIn(anh) };
    }, "uploadAnh"),
    clearAnh: /* @__PURE__ */ __name(function(p) {
      var it = getAnhFolder(String(p.folderId)).getFiles(), n = 0;
      while (it.hasNext()) {
        it.next().setTrashed(true);
        n++;
      }
      return { removed: n };
    }, "clearAnh"),
    saveFile: /* @__PURE__ */ __name(function(p) {
      var folder = DriveApp.getFolderById(String(p.folderId));
      var file = saveBlobToFolder(folder, Utilities.newBlob(Utilities.base64Decode(String(p.data)), String(p.mime || "application/octet-stream"), String(p.name)), String(p.name));
      return { fileId: file.getId() };
    }, "saveFile"),
    genDapAnFile: /* @__PURE__ */ __name(function(p) {
      var header = p.header, rows = p.rows;
      if (typeof header === "string") header = JSON.parse(header);
      if (typeof rows === "string") rows = JSON.parse(rows);
      var tmp = SpreadsheetApp.create("tmp-dapan-" + Date.now());
      try {
        var s2 = tmp.insertSheet("DapAn"), nCol = header.length, nRow = rows.length + 1;
        s2.getRange(1, 1, 1, nCol).setNumberFormat("@");
        if (rows.length) s2.getRange(2, 2, rows.length, nCol - 1).setNumberFormat("@");
        s2.getRange(1, 1, nRow, nCol).setValues([header].concat(rows));
        SpreadsheetApp.flush();
        var blob = exportSpreadsheetAsXlsxBlob(tmp.getId(), "dap-an.xlsx");
        var folder = DriveApp.getFolderById(String(p.folderId));
        var file = saveBlobToFolder(folder, blob, String(p.name));
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return { fileId: file.getId() };
      } finally {
        try {
          DriveApp.getFileById(tmp.getId()).setTrashed(true);
        } catch (e) {
        }
      }
    }, "genDapAnFile"),
    trash: /* @__PURE__ */ __name(function(p) {
      try {
        DriveApp.getFolderById(String(p.id)).setTrashed(true);
      } catch (e) {
        try {
          DriveApp.getFileById(String(p.id)).setTrashed(true);
        } catch (e2) {
        }
      }
      return { ok: true };
    }, "trash"),
    cleanupOldImages: /* @__PURE__ */ __name(function(p) {
      var folders = p.folders || [];
      if (typeof folders === "string") folders = JSON.parse(folders);
      var n = 0;
      folders.forEach(function(id) {
        try {
          var it = DriveApp.getFolderById(String(id)).getFoldersByName("anh");
          while (it.hasNext()) {
            it.next().setTrashed(true);
            n++;
          }
        } catch (e) {
        }
      });
      return { removed: n };
    }, "cleanupOldImages"),
    /* ---- email ---- */
    mail: /* @__PURE__ */ __name(function(p) {
      MailApp.sendEmail({ to: String(p.to), subject: String(p.subject || ""), body: String(p.body || ""), name: String(p.name || "H\u1EC7 th\u1ED1ng \u0111i\u1EC3m") });
      return { ok: true };
    }, "mail"),
    /* ---- sao lưu từ D1 về file Sheet hệ thống cũ (ghi đè từng tab) ---- */
    backup: /* @__PURE__ */ __name(function(p) {
      var ss = SpreadsheetApp.openById(String(p.sheetId)), tabs = p.tabs || [];
      if (typeof tabs === "string") tabs = JSON.parse(tabs);
      tabs.forEach(function(t) {
        var sh = ss.getSheetByName(t.name) || ss.insertSheet(t.name);
        var vals = (t.values || []).map(function(r) {
          return r.map(function(v) {
            return v === null || v === void 0 ? "" : v;
          });
        });
        var nc = vals.reduce(function(m, r) {
          return Math.max(m, r.length);
        }, 1);
        vals = vals.map(function(r) {
          while (r.length < nc) r.push("");
          return r;
        });
        var lr = sh.getLastRow(), lc = sh.getLastColumn();
        if (lr && lc) sh.getRange(1, 1, lr, lc).clearContent();
        if (vals.length) {
          sh.getRange(1, 1, vals.length, nc).setNumberFormat("@");
          sh.getRange(1, 1, vals.length, nc).setValues(vals);
        }
      });
      var meta = ss.getSheetByName("_SaoLuu") || ss.insertSheet("_SaoLuu");
      meta.getRange(1, 1, 1, 2).setValues([["Sao l\u01B0u t\u1EF1 \u0111\u1ED9ng t\u1EEB m\xE1y ch\u1EE7 m\u1EDBi l\xFAc", nowIso()]]);
      return { ok: true, tabs: tabs.length };
    }, "backup")
  };
  function bridgeCell(v) {
    if (v instanceof Date) return v.toISOString();
    return v === null || v === void 0 ? "" : v;
  }
  __name(bridgeCell, "bridgeCell");
  function taoKhoaCauNoi() {
    var p = _props(), k = p.getProperty("BRIDGE_SECRET");
    if (!k) {
      k = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
      p.setProperty("BRIDGE_SECRET", k);
    }
    Logger.log("BRIDGE_SECRET = " + k);
    return k;
  }
  __name(taoKhoaCauNoi, "taoKhoaCauNoi");
  function doGet(e) {
    return handle(e, "GET");
  }
  __name(doGet, "doGet");
  function doPost(e) {
    return handle(e, "POST");
  }
  __name(doPost, "doPost");
  function handle(e, method) {
    var t0 = Date.now(), p = {}, action = "";
    try {
      _ROWS_MEMO = {};
      _SS_MEMO = {};
      _DVER = null;
      _KHONG_DOI = false;
      if (method === "POST" && e && e.postData && e.postData.contents) p = JSON.parse(e.postData.contents);
      else if (e) p = e.parameter || {};
      if (p.bridge !== void 0) return jsonOut(bridgeHandle(p));
      action = String(p.action || "");
      if (action !== "ping") ensureSchema();
      var data = route(action, p);
      if (!SERVER_READ[action] && !_KHONG_DOI) {
        try {
          dataTouch();
        } catch (e2) {
        }
      }
      return jsonOut({ ok: true, data, ms: Date.now() - t0, v: APP_VERSION });
    } catch (err) {
      var msg = String(err && err.message || err);
      var code = err && err.code || (/Phiên đã hết hạn/.test(msg) ? "AUTH" : /Không đủ quyền|Không có quyền/.test(msg) ? "FORBIDDEN" : "ERR");
      if (err && err.name && err.name !== "Error") {
        ghiLoi(action, p, err);
        msg = "L\u1ED7i h\u1EC7 th\u1ED1ng khi x\u1EED l\xFD (" + action + "). \u0110\xE3 ghi l\u1EA1i \u0111\u1EC3 s\u1EEDa. Chi ti\u1EBFt: " + msg;
      }
      return jsonOut({ ok: false, error: msg, code, ms: Date.now() - t0, v: APP_VERSION });
    }
  }
  __name(handle, "handle");
  var SERVER_READ = {
    ping: 1,
    me: 1,
    bootstrap: 1,
    listTeachers: 1,
    trangThaiTuDong: 1,
    nhatKyLoi: 1,
    validateRoster: 1,
    getBai: 1,
    publicConfig: 1,
    traDiem: 1,
    chiTietBai: 1,
    listBai: 1,
    getBaiFull: 1,
    getKetQua: 1,
    thongKeBai: 1,
    listGroups: 1,
    listStudents: 1,
    listUsers: 1,
    getCaiDat: 1,
    listGopY: 1,
    getNhatKy: 1,
    previewDiem: 1,
    ddTongQuan: 1,
    ddPhanTich: 1,
    ddXemBuoi: 1,
    ngoaiDoiChieu: 1
  };
  function route(action, p) {
    switch (action) {
      /* ---- công khai, không cần đăng nhập ---- */
      case "ping":
        return { time: nowIso(), v: APP_VERSION };
      case "login":
        return apiLogin(p);
      case "quenMatKhau":
        return quenMatKhau(p);
      case "datLaiMatKhau":
        return datLaiMatKhau(p);
      case "publicConfig":
        return publicConfig();
      case "traDiem":
        return traDiem(p);
      case "chiTietBai":
        return chiTietBai(p);
      case "guiGopY":
        return guiGopY(p, optionalUser(p));
      /* ---- tài khoản ---- */
      case "logout":
        return apiLogout(p);
      case "me":
        return { user: currentUser(p) };
      case "changePassword":
        return apiChangePassword(p);
      case "bootstrap":
        return requireAuth(p, function(u) {
          return bootstrap(p, u);
        });
      /* ---- quản trị ---- */
      case "listUsers":
        return requireRole(p, "admin", function() {
          return listUsers();
        });
      case "listTeachers":
        return requireRole(p, "admin", function() {
          return listTeachers();
        });
      case "createUser":
        return requireRole(p, "admin", function() {
          return createUser(p);
        });
      case "setTeacherRoster":
        return requireRole(p, "admin", function() {
          return setTeacherRoster(p);
        });
      case "setUserInfo":
        return requireRole(p, "admin", function() {
          return setUserInfo(p);
        });
      case "setUserActive":
        return requireRole(p, "admin", function(u) {
          return setUserActive(p, u);
        });
      case "resetPassword":
        return requireRole(p, "admin", function() {
          return adminResetPassword(p);
        });
      case "getCaiDat":
        return requireRole(p, "admin", function() {
          return getCaiDatAll();
        });
      case "setCaiDat":
        return requireRole(p, "admin", function() {
          return setCaiDat(p);
        });
      case "listGopY":
        return requireRole(p, "admin", function() {
          return listGopY();
        });
      case "xuLyGopY":
        return requireRole(p, "admin", function() {
          return xuLyGopY(p);
        });
      case "trangThaiTuDong":
        return requireRole(p, "admin", function() {
          return trangThaiTuDong();
        });
      case "batTuDong":
        return requireRole(p, "admin", function() {
          return batTuDong();
        });
      case "nhatKyLoi":
        return requireRole(p, "admin", function() {
          return nhatKyLoi();
        });
      /* ---- nhóm, học sinh ---- */
      case "syncGroups":
        return requireStaff(p, function(u) {
          return apiSyncGroups(p, u);
        });
      case "listGroups":
        return requireStaff(p, function(u) {
          return listGroups(u, !!p.fresh);
        });
      case "listStudents":
        return requireStaff(p, function(u) {
          return listStudents(p, u);
        });
      case "upsertStudent":
        return requireStaff(p, function(u) {
          return upsertStudent(p, u);
        });
      case "setStudentStatus":
        return requireStaff(p, function(u) {
          return setStudentStatus(p, u);
        });
      case "transferStudent":
        return requireStaff(p, function(u) {
          return transferStudent(p, u);
        });
      case "validateRoster":
        return requireStaff(p, function(u) {
          return validateRoster(u);
        });
      /* ---- bài kiểm tra, đáp án ---- */
      case "createBai":
        return requireStaff(p, function(u) {
          return createBai(p, u);
        });
      case "updateBai":
        return requireStaff(p, function(u) {
          return updateBai(p, u);
        });
      case "listBai":
        return requireStaff(p, function(u) {
          return listBai(u, p);
        });
      case "getBai":
        return requireStaff(p, function(u) {
          return getBai(p, u);
        });
      case "getBaiFull":
        return requireStaff(p, function(u) {
          return getBaiFull(p, u);
        });
      case "saveDapAn":
        return requireStaff(p, function(u) {
          return saveDapAn(p, u);
        });
      case "genDapAnFile":
        return requireStaff(p, function(u) {
          return genDapAnFile(p, u);
        });
      case "luuTruBai":
        return requireStaff(p, function(u) {
          return luuTruBai(p, u);
        });
      case "xoaBai":
        return requireStaff(p, function(u) {
          return xoaBai(p, u);
        });
      case "luuFileDrive":
        return requireStaff(p, function(u) {
          return luuFileDrive(p, u);
        });
      /* ---- ảnh phiếu ---- */
      case "anhInfo":
        return requireStaff(p, function(u) {
          return anhInfo(p, u);
        });
      case "uploadAnh":
        return requireStaff(p, function(u) {
          return uploadAnh(p, u);
        });
      case "clearAnh":
        return requireStaff(p, function(u) {
          return clearAnh(p, u);
        });
      case "finishAnh":
        return requireStaff(p, function(u) {
          return finishAnh(p, u);
        });
      /* ---- điểm ---- */
      case "previewDiem":
        return requireStaff(p, function(u) {
          return previewDiem(p, u);
        });
      case "commitDiem":
        return requireStaff(p, function(u) {
          return commitDiem(p, u);
        });
      case "getKetQua":
        return requireStaff(p, function(u) {
          return getKetQua(p, u);
        });
      case "editKetQua":
        return requireStaff(p, function(u) {
          return editKetQua(p, u);
        });
      case "xoaKetQua":
        return requireStaff(p, function(u) {
          return xoaKetQua(p, u);
        });
      case "setCongBo":
        return requireStaff(p, function(u) {
          return setCongBo(p, u);
        });
      case "getNhatKy":
        return requireStaff(p, function(u) {
          return getNhatKy(p, u);
        });
      case "giaiQuyetChoXuLy":
        return requireStaff(p, function(u) {
          return giaiQuyetChoXuLy(p, u);
        });
      case "thongKeBai":
        return requireStaff(p, function(u) {
          return thongKeBai(p, u);
        });
      /* ---- điểm danh (lớp trưởng dùng được: quyền kiểm tra bên trong) ---- */
      case "ddTongQuan":
        return requireAuth(p, function(u) {
          return ddTongQuan(p, u);
        });
      case "ddMoBuoi":
        return requireAuth(p, function(u) {
          return ddMoBuoi(p, u);
        });
      case "ddXemBuoi":
        return requireAuth(p, function(u) {
          return ddXemBuoi(p, u);
        });
      case "ddLuu":
        return requireAuth(p, function(u) {
          return ddLuu(p, u);
        });
      case "ddNop":
        return requireAuth(p, function(u) {
          return ddNop(p, u);
        });
      case "ddYeuCau":
        return requireAuth(p, function(u) {
          return ddYeuCau(p, u);
        });
      case "ddChot":
        return requireStaff(p, function(u) {
          return ddChot(p, u);
        });
      case "ddXoaBuoi":
        return requireStaff(p, function(u) {
          return ddXoaBuoi(p, u);
        });
      case "ddDuyet":
        return requireStaff(p, function(u) {
          return ddDuyet(p, u);
        });
      case "ddPhanTich":
        return requireStaff(p, function(u) {
          return ddPhanTich(p, u);
        });
      case "ltTao":
        return requireStaff(p, function(u) {
          return ltTao(p, u);
        });
      case "ltSua":
        return requireStaff(p, function(u) {
          return ltSua(p, u);
        });
      case "ddBaoSai":
        return ddYeuCau(p, null);
      case "ngoaiDoiChieu":
        return requireStaff(p, function(u) {
          return ngoaiDoiChieu(p, u);
        });
      case "ngoaiLuu":
        return requireStaff(p, function(u) {
          return ngoaiLuu(p, u);
        });
      case "ngoaiSuaDiem":
        return requireStaff(p, function(u) {
          return ngoaiSuaDiem(p, u);
        });
      default:
        throw new Error("H\xE0nh \u0111\u1ED9ng kh\xF4ng h\u1EE3p l\u1EC7: " + action);
    }
  }
  __name(route, "route");
  function requireAuth(p, fn) {
    return fn(currentUser(p));
  }
  __name(requireAuth, "requireAuth");
  function requireStaff(p, fn) {
    var u = currentUser(p);
    if (u.role !== "admin" && u.role !== "teacher") {
      var e = new Error("Kh\xF4ng \u0111\u1EE7 quy\u1EC1n");
      e.code = "FORBIDDEN";
      throw e;
    }
    return fn(u);
  }
  __name(requireStaff, "requireStaff");
  function requireRole(p, role, fn) {
    var u = currentUser(p);
    if (u.role !== role) {
      var e = new Error("Kh\xF4ng \u0111\u1EE7 quy\u1EC1n");
      e.code = "FORBIDDEN";
      throw e;
    }
    return fn(u);
  }
  __name(requireRole, "requireRole");
  const _fn = {};
  try {
    _fn.handle = handle;
  } catch (e) {
  }
  try {
    _fn.route = route;
  } catch (e) {
  }
  try {
    _fn.currentUser = currentUser;
  } catch (e) {
  }
  try {
    _fn.requireAuth = requireAuth;
  } catch (e) {
  }
  try {
    _fn.requireStaff = requireStaff;
  } catch (e) {
  }
  try {
    _fn.findBai = findBai;
  } catch (e) {
  }
  try {
    _fn.requireBaiAccess = requireBaiAccess;
  } catch (e) {
  }
  try {
    _fn.sysRows = sysRows;
  } catch (e) {
  }
  try {
    _fn.sysTouch = sysTouch;
  } catch (e) {
  }
  try {
    _fn.getSetting = getSetting;
  } catch (e) {
  }
  try {
    _fn.getBai = getBai;
  } catch (e) {
  }
  try {
    _fn.parseJsonSafe = parseJsonSafe;
  } catch (e) {
  }
  try {
    _fn.safeName = safeName;
  } catch (e) {
  }
  try {
    _fn.ngayStr = ngayStr;
  } catch (e) {
  }
  try {
    _fn.parseNhomList = parseNhomList;
  } catch (e) {
  }
  try {
    _fn.donDepHangNgay = donDepHangNgay;
  } catch (e) {
  }
  try {
    _fn.cleanupOldImages = cleanupOldImages;
  } catch (e) {
  }
  try {
    _fn.pubTouch = pubTouch;
  } catch (e) {
  }
  try {
    _fn.dataTouch = dataTouch;
  } catch (e) {
  }
  try {
    _fn.ensureSchema = ensureSchema;
  } catch (e) {
  }
  try {
    _fn.setTextCell = setTextCell;
  } catch (e) {
  }
  try {
    _fn.idx = idx;
  } catch (e) {
  }
  try {
    _fn.nowIso = nowIso;
  } catch (e) {
  }
  try {
    _fn.getSystemSheet = getSystemSheet;
  } catch (e) {
  }
  try {
    _fn.getNhomMap = getNhomMap;
  } catch (e) {
  }
  try {
    _fn.rosterTouch = rosterTouch;
  } catch (e) {
  }
  try {
    _fn.listGroups = listGroups;
  } catch (e) {
  }
  try {
    _fn.syncTeacher = syncTeacher;
  } catch (e) {
  }
  try {
    _fn.logSua = logSua;
  } catch (e) {
  }
  return _fn;
}
__name(createApp, "createApp");

// src/sha256.js
var K = [1116352408, 1899447441, 3049323471, 3921009573, 961987163, 1508970993, 2453635748, 2870763221, 3624381080, 310598401, 607225278, 1426881987, 1925078388, 2162078206, 2614888103, 3248222580, 3835390401, 4022224774, 264347078, 604807628, 770255983, 1249150122, 1555081692, 1996064986, 2554220882, 2821834349, 2952996808, 3210313671, 3336571891, 3584528711, 113926993, 338241895, 666307205, 773529912, 1294757372, 1396182291, 1695183700, 1986661051, 2177026350, 2456956037, 2730485921, 2820302411, 3259730800, 3345764771, 3516065817, 3600352804, 4094571909, 275423344, 430227734, 506948616, 659060556, 883997877, 958139571, 1322822218, 1537002063, 1747873779, 1955562222, 2024104815, 2227730452, 2361852424, 2428436474, 2756734187, 3204031479, 3329325298];
var enc = new TextEncoder();
function sha256Hex(str) {
  const bytes = enc.encode(str), l = bytes.length;
  const n = l + 9 + 63 >> 6 << 6, buf = new Uint8Array(n);
  buf.set(bytes);
  buf[l] = 128;
  const dv = new DataView(buf.buffer);
  dv.setUint32(n - 4, l * 8 >>> 0);
  dv.setUint32(n - 8, Math.floor(l * 8 / 4294967296));
  let h0 = 1779033703, h1 = 3144134277, h2 = 1013904242, h3 = 2773480762, h4 = 1359893119, h5 = 2600822924, h6 = 528734635, h7 = 1541459225;
  const w = new Int32Array(64);
  for (let off = 0; off < n; off += 64) {
    for (let i = 0; i < 16; i++) w[i] = dv.getInt32(off + i * 4);
    for (let i = 16; i < 64; i++) {
      const x = w[i - 15], y = w[i - 2];
      const s0 = (x >>> 7 | x << 25) ^ (x >>> 18 | x << 14) ^ x >>> 3;
      const s1 = (y >>> 17 | y << 15) ^ (y >>> 19 | y << 13) ^ y >>> 10;
      w[i] = w[i - 16] + s0 + w[i - 7] + s1 | 0;
    }
    let a = h0, b = h1, c = h2, d = h3, e = h4, f = h5, g = h6, h = h7;
    for (let i = 0; i < 64; i++) {
      const S1 = (e >>> 6 | e << 26) ^ (e >>> 11 | e << 21) ^ (e >>> 25 | e << 7);
      const t1 = h + S1 + (e & f ^ ~e & g) + K[i] + w[i] | 0;
      const S0 = (a >>> 2 | a << 30) ^ (a >>> 13 | a << 19) ^ (a >>> 22 | a << 10);
      const t2 = S0 + (a & b ^ a & c ^ b & c) | 0;
      h = g;
      g = f;
      f = e;
      e = d + t1 | 0;
      d = c;
      c = b;
      b = a;
      a = t1 + t2 | 0;
    }
    h0 = h0 + a | 0;
    h1 = h1 + b | 0;
    h2 = h2 + c | 0;
    h3 = h3 + d | 0;
    h4 = h4 + e | 0;
    h5 = h5 + f | 0;
    h6 = h6 + g | 0;
    h7 = h7 + h | 0;
  }
  return [h0, h1, h2, h3, h4, h5, h6, h7].map((x) => (x >>> 0).toString(16).padStart(8, "0")).join("");
}
__name(sha256Hex, "sha256Hex");

// src/shim.js
var Store = class _Store {
  static {
    __name(this, "Store");
  }
  constructor(db) {
    this.db = db;
    this.ss = /* @__PURE__ */ new Map();
    this.kv = /* @__PURE__ */ new Map();
    this.kvDirty = /* @__PURE__ */ new Map();
    this.sheetsDirty = /* @__PURE__ */ new Set();
  }
  static async load(db) {
    const st = new _Store(db);
    const [rows, sheets, kv] = await Promise.all([
      db.prepare("SELECT ss, tab, rn, data FROM rows ORDER BY ss, tab, rn").all(),
      db.prepare("SELECT ss, tab, pos FROM sheets ORDER BY ss, pos").all(),
      db.prepare("SELECT k, v FROM kv").all()
    ]);
    for (const s of sheets.results) st.spreadsheet(s.ss).addSheet(s.tab, s.pos);
    for (const r of rows.results) {
      const sh = st.spreadsheet(r.ss).getSheetByName(r.tab) || st.spreadsheet(r.ss).addSheet(r.tab);
      sh.rows[r.rn - 1] = JSON.parse(r.data);
    }
    for (const s of st.ss.values()) for (const sh of s.sheets.values()) {
      for (let i = 0; i < sh.rows.length; i++) if (!sh.rows[i]) sh.rows[i] = [];
      sh.dirty = false;
      sh.structure = false;
    }
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
        if (sh.removed) {
          S.push(this.db.prepare("DELETE FROM rows WHERE ss=? AND tab=?").bind(s.id, sh.name));
          S.push(this.db.prepare("DELETE FROM sheets WHERE ss=? AND tab=?").bind(s.id, sh.name));
          continue;
        }
        if (sh.isNew) S.push(this.db.prepare("INSERT OR REPLACE INTO sheets(ss,tab,pos) VALUES(?,?,?)").bind(s.id, sh.name, sh.pos));
        if (!sh.dirty && !sh.isNew) continue;
        if (sh.structure || sh.isNew) {
          S.push(this.db.prepare("DELETE FROM rows WHERE ss=? AND tab=?").bind(s.id, sh.name));
          sh.rows.forEach((r, i) => S.push(this.db.prepare("INSERT INTO rows(ss,tab,rn,data) VALUES(?,?,?,?)").bind(s.id, sh.name, i + 1, JSON.stringify(r))));
        } else {
          for (const rn of sh.dirtyRows) {
            if (rn > sh.rows.length) S.push(this.db.prepare("DELETE FROM rows WHERE ss=? AND tab=? AND rn=?").bind(s.id, sh.name, rn));
            else S.push(this.db.prepare("INSERT OR REPLACE INTO rows(ss,tab,rn,data) VALUES(?,?,?,?)").bind(s.id, sh.name, rn, JSON.stringify(sh.rows[rn - 1])));
          }
        }
      }
    }
    for (const [k, v] of this.kvDirty) S.push(v === null ? this.db.prepare("DELETE FROM kv WHERE k=?").bind(k) : this.db.prepare("INSERT OR REPLACE INTO kv(k,v) VALUES(?,?)").bind(k, v));
    return S;
  }
  async flush() {
    const S = this.flushStatements();
    if (S.length) {
      for (let i = 0; i < S.length; i += 90) await this.db.batch(S.slice(i, i + 90));
    }
    return S.length;
  }
  hasChanges() {
    return this.flushStatements().length > 0;
  }
};
var Spreadsheet = class {
  static {
    __name(this, "Spreadsheet");
  }
  constructor(store, id) {
    this.store = store;
    this.id = id;
    this.sheets = /* @__PURE__ */ new Map();
  }
  addSheet(name, pos) {
    const sh = new Sheet(this, name, pos === void 0 ? this.sheets.size : pos);
    this.sheets.set(name, sh);
    return sh;
  }
  getId() {
    return this.id;
  }
  getUrl() {
    return "https://docs.google.com/spreadsheets/d/" + this.id;
  }
  getName() {
    return this.id;
  }
  getSheetByName(n) {
    const s = this.sheets.get(n);
    return s && !s.removed ? s : null;
  }
  getSheets() {
    return [...this.sheets.values()].filter((s) => !s.removed).sort((a, b) => a.pos - b.pos);
  }
  insertSheet(name) {
    if (this.getSheetByName(name)) throw new Error("\u0110\xE3 c\xF3 t\u1EDD " + name);
    const sh = this.addSheet(name);
    sh.isNew = true;
    sh.dirty = true;
    return sh;
  }
  deleteSheet(sh) {
    sh.removed = true;
  }
};
var Sheet = class {
  static {
    __name(this, "Sheet");
  }
  constructor(ss, name, pos) {
    this.ss = ss;
    this.name = name;
    this.pos = pos;
    this.rows = [];
    this.dirty = false;
    this.structure = false;
    this.dirtyRows = /* @__PURE__ */ new Set();
    this.isNew = false;
    this.removed = false;
  }
  getName() {
    return this.name;
  }
  setName(n) {
    this.ss.sheets.delete(this.name);
    this.name = n;
    this.ss.sheets.set(n, this);
    this.isNew = true;
    this.dirty = true;
    return this;
  }
  getParent() {
    return this.ss;
  }
  getSheetId() {
    return this.pos;
  }
  getLastRow() {
    let n = this.rows.length;
    while (n > 0 && this.rows[n - 1].every((v) => v === "" || v === null || v === void 0)) n--;
    return n;
  }
  getLastColumn() {
    let m = 0;
    for (const r of this.rows) {
      let k = r.length;
      while (k > 0 && (r[k - 1] === "" || r[k - 1] === null || r[k - 1] === void 0)) k--;
      if (k > m) m = k;
    }
    return m;
  }
  getMaxRows() {
    return Math.max(1e3, this.rows.length);
  }
  getMaxColumns() {
    return Math.max(26, this.getLastColumn());
  }
  _touch(rn) {
    this.dirty = true;
    this.dirtyRows.add(rn);
  }
  _get(r, c) {
    const row = this.rows[r - 1];
    const v = row ? row[c - 1] : void 0;
    return v === void 0 || v === null ? "" : v;
  }
  _set(r, c, v) {
    while (this.rows.length < r) this.rows.push([]);
    const row = this.rows[r - 1];
    while (row.length < c) row.push("");
    row[c - 1] = v === void 0 || v === null ? "" : v;
    this._touch(r);
  }
  getRange(a, b, c, d) {
    if (typeof a === "string") throw new Error("getRange(A1) kh\xF4ng h\u1ED7 tr\u1EE3: " + a);
    return new Range(this, a, b, c === void 0 ? 1 : c, d === void 0 ? 1 : d);
  }
  getDataRange() {
    const lr = Math.max(1, this.getLastRow()), lc = Math.max(1, this.getLastColumn());
    return new Range(this, 1, 1, lr, lc);
  }
  appendRow(vals) {
    const r = this.getLastRow() + 1;
    this.rows.length = Math.max(this.rows.length, r - 1);
    this.rows[r - 1] = vals.map((v) => v === void 0 || v === null ? "" : v);
    this._touch(r);
    return this;
  }
  deleteRow(r) {
    return this.deleteRows(r, 1);
  }
  deleteRows(r, n) {
    this.rows.splice(r - 1, n || 1);
    this.dirty = true;
    this.structure = true;
    return this;
  }
  insertRowsAfter(r, n) {
    const a = [];
    for (let i = 0; i < (n || 1); i++) a.push([]);
    this.rows.splice(r, 0, ...a);
    this.dirty = true;
    this.structure = true;
    return this;
  }
  clear() {
    this.rows = [];
    this.dirty = true;
    this.structure = true;
    return this;
  }
  clearContents() {
    return this.clear();
  }
  setFrozenRows() {
    return this;
  }
  autoResizeColumns() {
    return this;
  }
  setColumnWidth() {
    return this;
  }
  hideColumns() {
    return this;
  }
};
var Range = class {
  static {
    __name(this, "Range");
  }
  constructor(sh, r, c, nr, nc) {
    this.sh = sh;
    this.r = r;
    this.c = c;
    this.nr = nr;
    this.nc = nc;
  }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(this.sh._get(this.r + i, this.c + j));
      out.push(row);
    }
    return out;
  }
  getDisplayValues() {
    return this.getValues().map((r) => r.map((v) => String(v)));
  }
  getValue() {
    return this.sh._get(this.r, this.c);
  }
  setValues(vals) {
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh._set(this.r + i, this.c + j, vals[i] ? vals[i][j] : "");
    return this;
  }
  setValue(v) {
    this.sh._set(this.r, this.c, v);
    return this;
  }
  setNumberFormat() {
    return this;
  }
  setNumberFormats() {
    return this;
  }
  setFontWeight() {
    return this;
  }
  setBackground() {
    return this;
  }
  setWrap() {
    return this;
  }
  clearContent() {
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh._set(this.r + i, this.c + j, "");
    return this;
  }
  getRow() {
    return this.r;
  }
  getColumn() {
    return this.c;
  }
  getNumRows() {
    return this.nr;
  }
  getNumColumns() {
    return this.nc;
  }
};
var MEM = /* @__PURE__ */ new Map();
var CacheService = {
  getScriptCache() {
    return {
      get(k) {
        const e = MEM.get(k);
        if (!e) return null;
        if (e.exp && e.exp < Date.now()) {
          MEM.delete(k);
          return null;
        }
        return e.v;
      },
      put(k, v, ttl) {
        MEM.set(k, { v: String(v), exp: Date.now() + (ttl || 600) * 1e3 });
      },
      remove(k) {
        MEM.delete(k);
      },
      putAll(m, ttl) {
        for (const k in m) this.put(k, m[k], ttl);
      },
      getAll(ks) {
        const o = {};
        for (const k of ks) {
          const v = this.get(k);
          if (v !== null) o[k] = v;
        }
        return o;
      },
      removeAll(ks) {
        for (const k of ks) MEM.delete(k);
      }
    };
  },
  getUserCache() {
    return this.getScriptCache();
  },
  getDocumentCache() {
    return this.getScriptCache();
  }
};
function makeServices(store, ctx) {
  const side = { mails: [], drive: [] };
  const props = {
    getProperty: /* @__PURE__ */ __name((k) => store.kv.has(k) ? store.kv.get(k) : null, "getProperty"),
    setProperty(k, v) {
      store.kv.set(k, String(v));
      store.kvDirty.set(k, String(v));
      return this;
    },
    deleteProperty(k) {
      store.kv.delete(k);
      store.kvDirty.set(k, null);
      return this;
    },
    getProperties() {
      const o = {};
      for (const [k, v] of store.kv) o[k] = v;
      return o;
    },
    getKeys() {
      return [...store.kv.keys()];
    }
  };
  const dummyFolder = /* @__PURE__ */ __name((id) => ({
    getId: /* @__PURE__ */ __name(() => id || "", "getId"),
    getUrl: /* @__PURE__ */ __name(() => "https://drive.google.com/drive/folders/" + (id || ""), "getUrl"),
    getName: /* @__PURE__ */ __name(() => "", "getName"),
    getFoldersByName: /* @__PURE__ */ __name(() => iter([]), "getFoldersByName"),
    getFilesByName: /* @__PURE__ */ __name(() => iter([]), "getFilesByName"),
    getFiles: /* @__PURE__ */ __name(() => iter([]), "getFiles"),
    getFolders: /* @__PURE__ */ __name(() => iter([]), "getFolders"),
    createFolder: /* @__PURE__ */ __name(() => dummyFolder(""), "createFolder"),
    createFile: /* @__PURE__ */ __name(() => ({ getId: /* @__PURE__ */ __name(() => "", "getId"), setName() {
      return this;
    }, setSharing() {
      return this;
    } }), "createFile"),
    setTrashed() {
      return this;
    },
    isTrashed: /* @__PURE__ */ __name(() => false, "isTrashed")
  }), "dummyFolder");
  const iter = /* @__PURE__ */ __name((arr) => {
    let i = 0;
    return { hasNext: /* @__PURE__ */ __name(() => i < arr.length, "hasNext"), next: /* @__PURE__ */ __name(() => arr[i++], "next") };
  }, "iter");
  const G = {
    SpreadsheetApp: {
      openById: /* @__PURE__ */ __name((id) => store.spreadsheet(id), "openById"),
      create: /* @__PURE__ */ __name(() => {
        throw new Error("T\u1EA1o file Google Sheet m\u1EDBi ch\u1EC9 l\xE0m \u0111\u01B0\u1EE3c qua c\u1EA7u n\u1ED1i Apps Script");
      }, "create"),
      flush() {
      },
      getActiveSpreadsheet: /* @__PURE__ */ __name(() => null, "getActiveSpreadsheet")
    },
    PropertiesService: { getScriptProperties: /* @__PURE__ */ __name(() => props, "getScriptProperties"), getUserProperties: /* @__PURE__ */ __name(() => props, "getUserProperties"), getDocumentProperties: /* @__PURE__ */ __name(() => props, "getDocumentProperties") },
    CacheService,
    LockService: { getScriptLock: /* @__PURE__ */ __name(() => ({ waitLock() {
    }, tryLock: /* @__PURE__ */ __name(() => true, "tryLock"), releaseLock() {
    }, hasLock: /* @__PURE__ */ __name(() => true, "hasLock") }), "getScriptLock"), getDocumentLock() {
      return this.getScriptLock();
    }, getUserLock() {
      return this.getScriptLock();
    } },
    Utilities: {
      formatDate(d, tz, fmt) {
        const t = new Date(d.getTime() + 7 * 36e5);
        const p = /* @__PURE__ */ __name((n) => String(n).padStart(2, "0"), "p");
        return fmt.replace("yyyy", t.getUTCFullYear()).replace("MM", p(t.getUTCMonth() + 1)).replace("dd", p(t.getUTCDate())).replace("HH", p(t.getUTCHours())).replace("mm", p(t.getUTCMinutes())).replace("ss", p(t.getUTCSeconds()));
      },
      getUuid: /* @__PURE__ */ __name(() => crypto.randomUUID(), "getUuid"),
      sleep() {
      },
      computeDigest: /* @__PURE__ */ __name((alg, s) => sha256Hex(s).match(/../g).map((h) => {
        const b = parseInt(h, 16);
        return b > 127 ? b - 256 : b;
      }), "computeDigest"),
      DigestAlgorithm: { SHA_256: "SHA_256" },
      Charset: { UTF_8: "UTF_8" },
      base64Encode: /* @__PURE__ */ __name((s) => btoa(unescape(encodeURIComponent(s))), "base64Encode"),
      base64Decode: /* @__PURE__ */ __name((s) => s, "base64Decode"),
      newBlob: /* @__PURE__ */ __name((data, mime, name) => ({ data, mime, name, setName(n) {
        this.name = n;
        return this;
      }, getBytes: /* @__PURE__ */ __name(() => data, "getBytes") }), "newBlob")
    },
    ContentService: { createTextOutput: /* @__PURE__ */ __name((s) => ({ s, setMimeType() {
      return this;
    } }), "createTextOutput"), MimeType: { JSON: "json", TEXT: "text" } },
    DriveApp: {
      getFolderById: /* @__PURE__ */ __name((id) => dummyFolder(id), "getFolderById"),
      getFileById: /* @__PURE__ */ __name((id) => dummyFolder(id), "getFileById"),
      getRootFolder: /* @__PURE__ */ __name(() => dummyFolder("root"), "getRootFolder"),
      createFolder: /* @__PURE__ */ __name(() => dummyFolder(""), "createFolder"),
      Access: { ANYONE_WITH_LINK: 1 },
      Permission: { VIEW: 1 }
    },
    MailApp: { sendEmail: /* @__PURE__ */ __name((m) => {
      side.mails.push(m);
    }, "sendEmail"), getRemainingDailyQuota: /* @__PURE__ */ __name(() => 100, "getRemainingDailyQuota") },
    GmailApp: { sendEmail: /* @__PURE__ */ __name((to, subject, body, opt) => {
      side.mails.push(Object.assign({ to, subject, body }, opt || {}));
    }, "sendEmail") },
    ScriptApp: { getProjectTriggers: /* @__PURE__ */ __name(() => [{ getHandlerFunction: /* @__PURE__ */ __name(() => "donDepHangNgay", "getHandlerFunction") }], "getProjectTriggers"), newTrigger: /* @__PURE__ */ __name(() => ({ timeBased() {
      return this;
    }, everyDays() {
      return this;
    }, atHour() {
      return this;
    }, create() {
    } }), "newTrigger"), deleteTrigger() {
    }, getOAuthToken: /* @__PURE__ */ __name(() => "", "getOAuthToken") },
    Session: { getEffectiveUser: /* @__PURE__ */ __name(() => ({ getEmail: /* @__PURE__ */ __name(() => props.getProperty("OWNER_EMAIL") || "", "getEmail") }), "getEffectiveUser"), getActiveUser: /* @__PURE__ */ __name(() => ({ getEmail: /* @__PURE__ */ __name(() => props.getProperty("OWNER_EMAIL") || "", "getEmail") }), "getActiveUser"), getScriptTimeZone: /* @__PURE__ */ __name(() => "Asia/Ho_Chi_Minh", "getScriptTimeZone") },
    UrlFetchApp: { fetch: /* @__PURE__ */ __name(() => {
      throw new Error("UrlFetchApp kh\xF4ng d\xF9ng trong Worker");
    }, "fetch") },
    HtmlService: {},
    Logger: { log: /* @__PURE__ */ __name((...a) => console.log(...a), "log") },
    console
  };
  return { G, side, props };
}
__name(makeServices, "makeServices");

// src/bridge.js
var Bridge = class {
  static {
    __name(this, "Bridge");
  }
  constructor(env, db) {
    this.url = env.APPS_SCRIPT_URL || "";
    this.secret = env.BRIDGE_SECRET || "";
    this.db = db || env.DB;
  }
  get ok() {
    return (/^https:\/\//.test(this.url) || /^http:\/\/(localhost|127\.0\.0\.1)[:/]/.test(this.url)) && (!!this.secret || !!this.db);
  }
  async token() {
    if (this.secret) return this.secret;
    const n = crypto.randomUUID().replace(/-/g, "") + crypto.randomUUID().replace(/-/g, ""), now = Date.now();
    await this.db.batch([
      this.db.prepare("DELETE FROM nonce WHERE exp < ?").bind(now),
      this.db.prepare("INSERT INTO nonce (n, exp) VALUES (?, ?)").bind(n, now + 10 * 60 * 1e3)
    ]);
    return "nonce:" + n;
  }
  /** Apps Script gọi về: mã có thật và chưa hết hạn thì trả valid:true và xóa mã (dùng một lần) */
  static async verify(db, n) {
    try {
      const r = await db.prepare("DELETE FROM nonce WHERE n = ? AND exp > ? RETURNING n").bind(String(n || ""), Date.now()).all();
      return !!(r.results && r.results.length);
    } catch (e) {
      return false;
    }
  }
  async call(op, params, opts) {
    if (!this.ok) throw new Error("Ch\u01B0a c\u1EA5u h\xECnh c\u1EA7u n\u1ED1i Apps Script (APPS_SCRIPT_URL)");
    const body = JSON.stringify(Object.assign({ bridge: await this.token(), op }, params || {}));
    const ctrl = new AbortController();
    const to = setTimeout(() => ctrl.abort(), opts && opts.timeout || 6e4);
    try {
      const r = await fetch(this.url, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body, redirect: "follow", signal: ctrl.signal });
      const txt = await r.text();
      let j;
      try {
        j = JSON.parse(txt);
      } catch (e) {
        throw new Error("C\u1EA7u n\u1ED1i tr\u1EA3 v\u1EC1 d\u1EEF li\u1EC7u l\u1EA1: " + txt.slice(0, 120));
      }
      if (!j.ok) throw new Error(j.error || "C\u1EA7u n\u1ED1i b\xE1o l\u1ED7i");
      return j.data;
    } finally {
      clearTimeout(to);
    }
  }
  rosterRead(fileId) {
    return this.call("rosterRead", { fileId });
  }
  rosterWrite(fileId, tabs) {
    return this.call("rosterWrite", { fileId, tabs });
  }
  ensureBaiFolder(bai) {
    return this.call("ensureBaiFolder", { teacher: bai.teacher, tenBai: bai.tenBai, ngay: bai.ngay, folderId: bai.driveFolderId || "" });
  }
  anhInfo(folderId) {
    return this.call("anhInfo", { folderId });
  }
  uploadAnh(folderId, images) {
    return this.call("uploadAnh", { folderId, images }, { timeout: 12e4 });
  }
  clearAnh(folderId) {
    return this.call("clearAnh", { folderId });
  }
  saveFile(folderId, name, data, mime) {
    return this.call("saveFile", { folderId, name, data, mime }, { timeout: 12e4 });
  }
  genDapAnFile(folderId, name, header, rows) {
    return this.call("genDapAnFile", { folderId, name, header, rows });
  }
  trash(id) {
    return this.call("trash", { id });
  }
  mail(m) {
    return this.call("mail", { to: m.to, subject: m.subject, body: m.body, name: m.name || "H\u1EC7 th\u1ED1ng \u0111i\u1EC3m" });
  }
  backup(sheetId, tabs) {
    return this.call("backup", { sheetId, tabs }, { timeout: 12e4 });
  }
  cleanupOldImages(folders) {
    return this.call("cleanupOldImages", { folders });
  }
  ping() {
    return this.call("ping", {}, { timeout: 3e4 });
  }
};

// src/index.js
var CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Methods": "POST, GET, OPTIONS", "Access-Control-Allow-Headers": "Content-Type", "Access-Control-Max-Age": "86400" };
var json = /* @__PURE__ */ __name((obj, status) => new Response(JSON.stringify(obj), { status: status || 200, headers: Object.assign({ "Content-Type": "application/json;charset=utf-8", "Cache-Control": "no-store" }, CORS) }), "json");
var READ = /* @__PURE__ */ new Set(["ping", "me", "bootstrap", "publicConfig", "traDiem", "chiTietBai", "listBai", "getBai", "getBaiFull", "getKetQua", "thongKeBai", "listGroups", "listStudents", "listUsers", "listTeachers", "getCaiDat", "listGopY", "getNhatKy", "previewDiem", "ddTongQuan", "ddPhanTich", "ddXemBuoi", "ngoaiDoiChieu", "trangThaiTuDong", "nhatKyLoi", "validateRoster"]);
var DRIVE = /* @__PURE__ */ new Set(["anhInfo", "uploadAnh", "clearAnh", "finishAnh", "genDapAnFile", "luuFileDrive"]);
var ROSTER_PULL = /* @__PURE__ */ new Set(["syncGroups", "validateRoster"]);
var ROSTER_WRITE = /* @__PURE__ */ new Set(["upsertStudent", "setStudentStatus", "transferStudent", "syncGroups"]);
var LOCK_MS = 25e3;
async function acquireLock(db) {
  const deadline = Date.now() + 15e3, me = String(Date.now() + LOCK_MS) + "-" + Math.random().toString(36).slice(2);
  while (Date.now() < deadline) {
    const r = await db.prepare("UPDATE kv SET v=? WHERE k='__lock' AND (v='' OR CAST(substr(v,1,13) AS INTEGER) < ?)").bind(me, Date.now()).run();
    if (r.meta.changes > 0) return me;
    await new Promise((res) => setTimeout(res, 60 + Math.random() * 80));
  }
  throw new Error("H\u1EC7 th\u1ED1ng \u0111ang b\u1EADn ghi d\u1EEF li\u1EC7u, th\u1EED l\u1EA1i sau v\xE0i gi\xE2y");
}
__name(acquireLock, "acquireLock");
async function releaseLock(db, me) {
  try {
    await db.prepare("UPDATE kv SET v='' WHERE k='__lock' AND v=?").bind(me).run();
  } catch (e) {
  }
}
__name(releaseLock, "releaseLock");
var index_default = {
  async fetch(request, env, ctx) {
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS });
    const t0 = Date.now();
    let p = {};
    try {
      if (request.method === "POST") {
        const raw = await request.text();
        p = raw ? JSON.parse(raw) : {};
      } else {
        const u = new URL(request.url);
        u.searchParams.forEach((v, k) => {
          p[k] = v;
        });
      }
    } catch (e) {
      return json({ ok: false, error: "D\u1EEF li\u1EC7u g\u1EEDi l\xEAn kh\xF4ng \u0111\u1ECDc \u0111\u01B0\u1EE3c", code: "ERR" }, 400);
    }
    const action = String(p.action || "");
    if (action === "ping") return json({ ok: true, data: { ok: true, worker: true }, ms: 0, v: env.APP_VERSION || "7.0.0" });
    if (p.bridgeVerify !== void 0) return json({ ok: true, valid: await Bridge.verify(env.DB, p.bridgeVerify) });
    if (action === "kiemTraCauNoi") {
      const bridge0 = new Bridge(env);
      if (!bridge0.ok) return json({ ok: false, error: "Ch\u01B0a c\u1EA5u h\xECnh APPS_SCRIPT_URL", ms: Date.now() - t0 });
      try {
        const r = await bridge0.ping();
        return json({ ok: true, data: r, ms: Date.now() - t0 });
      } catch (e) {
        return json({ ok: false, error: String(e && e.message || e), ms: Date.now() - t0 });
      }
    }
    if (!action) return json({ ok: false, error: "Thi\u1EBFu action", code: "ERR" }, 400);
    const bridge = new Bridge(env), isRead = READ.has(action);
    let lock = null, store, result, status = 200;
    try {
      if (!isRead) lock = await acquireLock(env.DB);
      store = await Store.load(env.DB);
      if (ROSTER_PULL.has(action) || p.fresh && ["listStudents", "listGroups"].includes(action)) await pullRoster(store, bridge, p);
      const { G, side } = makeServices(store, ctx);
      G.WORKER_MODE = true;
      const app = createApp(G);
      if (action === "nhapTuSheet") result = await importFromSheet(p, app, store, bridge);
      else if (DRIVE.has(action)) result = await driveAction(action, p, app, store, bridge);
      else {
        let folderToTrash = "";
        if (action === "xoaBai") {
          try {
            const b = app.findBai(p.maBai);
            folderToTrash = b ? String(b.driveFolderId || "") : "";
          } catch (e) {
          }
        }
        const out = app.handle({ postData: { contents: JSON.stringify(p) } }, "POST");
        result = JSON.parse(out.s);
        if (result.ok && folderToTrash && bridge.ok) ctx.waitUntil(bridge.trash(folderToTrash).catch(() => {
        }));
      }
      const nWrites = await store.flush();
      if (side.mails.length && bridge.ok) for (const m of side.mails) ctx.waitUntil(bridge.mail(m).catch((e) => console.log("mail l\u1ED7i", e.message)));
      if (ROSTER_WRITE.has(action) && result.ok && bridge.ok) ctx.waitUntil(pushRoster(store, bridge).catch((e) => console.log("\u0111\u1ED3ng b\u1ED9 danh s\xE1ch l\u1ED7i", e.message)));
      result.ms = Date.now() - t0;
      result.w = nWrites;
      result.v = env.APP_VERSION || result.v;
    } catch (e) {
      console.log("L\u1ED6I", action, e && e.stack);
      result = { ok: false, error: String(e && e.message || e), code: "ERR", ms: Date.now() - t0 };
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
      const { G } = makeServices(store, ctx);
      G.WORKER_MODE = true;
      const app = createApp(G);
      try {
        app.donDepHangNgay();
      } catch (e) {
        console.log("d\u1ECDn d\u1EB9p l\u1ED7i", e.message);
      }
      try {
        const days = Number(app.getSetting("soNgayGiuAnh", 30)) || 30, now = Date.now(), folders = [];
        app.sysRows("BaiKiemTra").forEach((r) => {
          if (!r.ngayCongBo || !r.driveFolderId) return;
          const t = new Date(r.ngayCongBo).getTime();
          if (!isNaN(t) && now - t > days * 864e5) folders.push(String(r.driveFolderId));
        });
        if (folders.length && bridge.ok) await bridge.cleanupOldImages(folders);
      } catch (e) {
        console.log("d\u1ECDn \u1EA3nh l\u1ED7i", e.message);
      }
      await store.flush();
      if (bridge.ok && env.BACKUP_SHEET_ID) {
        const sys = store.spreadsheet("SYS"), tabs = sys.getSheets().filter((s) => s.getName() !== "Phien").map((s) => ({ name: s.getName(), values: s.getDataRange().getValues() }));
        await bridge.backup(env.BACKUP_SHEET_ID, tabs);
      }
    } finally {
      await releaseLock(env.DB, lock);
    }
  }
};
function rosterIds(store) {
  const ids = /* @__PURE__ */ new Set();
  const sys = store.spreadsheet("SYS"), nd = sys.getSheetByName("NguoiDung");
  if (nd) {
    const v = nd.getDataRange().getValues(), H = {};
    v[0].forEach((h, i) => {
      H[String(h).trim()] = i;
    });
    for (let i = 1; i < v.length; i++) {
      const id = String(v[i][H.rosterSheetId] || "").trim();
      if (id) ids.add(id);
    }
  }
  return [...ids];
}
__name(rosterIds, "rosterIds");
async function pullRoster(store, bridge, p) {
  if (!bridge.ok) return;
  for (const id of rosterIds(store)) {
    const data = await bridge.rosterRead(id);
    const ss = store.spreadsheet(id);
    for (const sh of ss.getSheets()) if (!data.tabs.some((t) => t.name === sh.getName())) ss.deleteSheet(sh);
    data.tabs.forEach((t, i) => {
      let sh = ss.getSheetByName(t.name);
      if (!sh) {
        sh = ss.insertSheet(t.name);
      }
      sh.pos = i;
      sh.isNew = true;
      sh.rows = (t.values || []).map((r) => r.map((v) => v === null || v === void 0 ? "" : v));
      sh.dirty = true;
      sh.structure = true;
    });
  }
}
__name(pullRoster, "pullRoster");
async function pushRoster(store, bridge) {
  for (const id of rosterIds(store)) {
    const ss = store.ss.get(id);
    if (!ss) continue;
    const tabs = ss.getSheets().filter((s) => s.dirty || s.isNew).map((s) => ({ name: s.getName(), values: s.rows }));
    if (tabs.length) await bridge.rosterWrite(id, tabs);
  }
}
__name(pushRoster, "pushRoster");
var SYS_TABS = ["NguoiDung", "Nhom", "Phien", "BaiKiemTra", "DapAn", "KetQua", "NhatKySua", "GopY", "NhatKyLoi", "BuoiHoc", "DiemDanhLog", "YeuCauSua", "ChoXuLy", "CaiDat"];
function trimRows(values) {
  const rows = (values || []).map((r) => (r || []).map((v) => v === null || v === void 0 ? "" : v));
  while (rows.length && rows[rows.length - 1].every((v) => String(v).trim() === "")) rows.pop();
  return rows;
}
__name(trimRows, "trimRows");
function overwriteTab(ss, name, pos, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.pos = pos;
  sh.isNew = true;
  sh.dirty = true;
  sh.structure = true;
  sh.rows = rows;
  sh.dirtyRows = /* @__PURE__ */ new Set();
  return sh;
}
__name(overwriteTab, "overwriteTab");
async function importFromSheet(p, app, store, bridge) {
  const sys = store.spreadsheet("SYS"), nd = sys.getSheetByName("NguoiDung");
  const empty = !nd || nd.rows.length <= 1;
  if (!empty) {
    const u = app.currentUser(p);
    if (u.role !== "admin") return { ok: false, error: "Ch\u1EC9 qu\u1EA3n tr\u1ECB m\u1EDBi \u0111\u01B0\u1EE3c nh\u1EADp l\u1EA1i d\u1EEF li\u1EC7u t\u1EEB Sheet", code: "FORBIDDEN" };
  }
  if (!bridge.ok) return { ok: false, error: "Ch\u01B0a c\u1EA5u h\xECnh c\u1EA7u n\u1ED1i Apps Script (APPS_SCRIPT_URL)", code: "ERR" };
  const data = await bridge.call("sysRead", {}, { timeout: 12e4 });
  const report = [];
  const byName = {};
  (data.tabs || []).forEach((t) => {
    byName[t.name] = t;
  });
  SYS_TABS.forEach((name, i) => {
    const t = byName[name];
    if (!t) {
      report.push([name, 0, "kh\xF4ng c\xF3 trong Sheet"]);
      return;
    }
    let rows = trimRows(t.values);
    if (name === "Phien") rows = rows.slice(0, 1);
    overwriteTab(sys, name, i, rows);
    report.push([name, Math.max(0, rows.length - 1), name === "Phien" ? "(phi\xEAn c\u0169 kh\xF4ng chuy\u1EC3n)" : ""]);
  });
  for (const sh of sys.getSheets()) if (!SYS_TABS.includes(sh.getName())) sys.deleteSheet(sh);
  const ids = rosterIds(store), roster = [];
  for (const id of ids) {
    const d = await bridge.rosterRead(id);
    const ss = store.spreadsheet(id);
    for (const sh of ss.getSheets()) if (!d.tabs.some((t) => t.name === sh.getName())) ss.deleteSheet(sh);
    d.tabs.forEach((t, i) => {
      const rows = trimRows(t.values);
      overwriteTab(ss, t.name, i, rows);
      roster.push([t.name, Math.max(0, rows.length - 2)]);
    });
  }
  const setKv = /* @__PURE__ */ __name((k, v) => {
    store.kv.set(k, v);
    store.kvDirty.set(k, v);
  }, "setKv");
  setKv("SYSTEM_SHEET_ID", "SYS");
  setKv("DA_CHUYEN_LUC", (/* @__PURE__ */ new Date()).toISOString());
  setKv("NGUON_SHEET_ID", String(data.sheetId || ""));
  if (p.ownerEmail) setKv("OWNER_EMAIL", String(p.ownerEmail));
  return { ok: true, data: { heThong: report, danhSach: roster, luc: store.kv.get("DA_CHUYEN_LUC") } };
}
__name(importFromSheet, "importFromSheet");
async function driveAction(action, p, app, store, bridge) {
  const u = app.currentUser(p);
  if (u.role !== "admin" && u.role !== "teacher") return { ok: false, error: "Kh\xF4ng \u0111\u1EE7 quy\u1EC1n", code: "FORBIDDEN" };
  const bai = app.findBai(p.maBai);
  app.requireBaiAccess(u, bai);
  const sys = store.spreadsheet("SYS"), sh = sys.getSheetByName("BaiKiemTra");
  const H = app.idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  const setCol = /* @__PURE__ */ __name((name, val) => {
    if (H[name] !== void 0) sh.getRange(bai._row, H[name] + 1).setValue(val);
  }, "setCol");
  const ensureFolder = /* @__PURE__ */ __name(async () => {
    if (bai.driveFolderId) return String(bai.driveFolderId);
    const r = await bridge.ensureBaiFolder({ teacher: bai.teacher, tenBai: bai.tenBai, ngay: app.ngayStr(bai.ngay), driveFolderId: "" });
    bai.driveFolderId = r.folderId;
    setCol("driveFolderId", r.folderId);
    return r.folderId;
  }, "ensureFolder");
  let data;
  if (action === "anhInfo") {
    const n = Number(bai.soAnh) || 0;
    data = { count: n, folderUrl: bai.driveFolderId ? "https://drive.google.com/drive/folders/" + bai.driveFolderId : "" };
  } else if (action === "uploadAnh") {
    const imgs = p.images || [{ name: p.name, data: p.data }];
    if (!imgs.length) throw new Error("Kh\xF4ng c\xF3 \u1EA3nh \u0111\u1EC3 t\u1EA3i");
    const fid = await ensureFolder();
    const r = await bridge.uploadAnh(fid, imgs);
    setCol("soAnh", (Number(bai.soAnh) || 0) + (r.added || 0));
    data = { ok: true, added: r.added };
  } else if (action === "clearAnh") {
    const fid = await ensureFolder();
    const r = await bridge.clearAnh(fid);
    setCol("soAnh", 0);
    store.kv.delete("ANHMAIL_" + bai.maBai);
    store.kvDirty.set("ANHMAIL_" + bai.maBai, null);
    data = { ok: true, removed: r.removed };
  } else if (action === "finishAnh") {
    const soAnh = Number(bai.soAnh) || 0, last = Number(store.kv.get("ANHMAIL_" + bai.maBai)) || 0;
    data = { soAnh, hasDapAn: !!bai.dapAnFileId, emailSent: false };
    const to = String(app.getSetting("emailNhanThongBao", "") || "").trim();
    if (soAnh > last && to && bridge.ok) {
      const cfg = app.parseJsonSafe(bai.cauHinh, {});
      const body = "C\xF3 \u1EA3nh phi\u1EBFu b\xE0i l\xE0m m\u1EDBi, c\xF3 th\u1EC3 ch\u1EA5m b\u1EB1ng TNMaker.\n\nB\xE0i: " + bai.tenBai + "\nNg\xE0y: " + (app.ngayStr(bai.ngay) || "(ch\u01B0a ghi)") + "\nGi\xE1o vi\xEAn: " + bai.teacher + "\nNh\xF3m: " + (app.parseNhomList(bai.cacNhom).join(", ") || "(ch\u01B0a ch\u1ECDn)") + "\nM\xE3 \u0111\u1EC1: " + (cfg.maDeList || []).join(", ") + "\nS\u1ED1 \u1EA3nh phi\u1EBFu: " + soAnh + "\nFile \u0111\xE1p \xE1n tr\xEAn h\u1EC7 th\u1ED1ng: " + (bai.dapAnFileId ? "C\xF3" : "Ch\u01B0a c\xF3") + "\n\nTh\u01B0 m\u1EE5c c\u1EE7a b\xE0i tr\xEAn Drive: https://drive.google.com/drive/folders/" + bai.driveFolderId + "\n";
      try {
        await bridge.mail({ to, subject: "[\u0110i\u1EC3m] C\xF3 \u1EA3nh phi\u1EBFu m\u1EDBi: " + bai.tenBai, body });
        data.emailSent = true;
        store.kv.set("ANHMAIL_" + bai.maBai, String(soAnh));
        store.kvDirty.set("ANHMAIL_" + bai.maBai, String(soAnh));
      } catch (e) {
      }
    }
  } else if (action === "luuFileDrive") {
    const d = String(p.data || "");
    if (!d) throw new Error("File r\u1ED7ng");
    if (d.length > 14 * 1024 * 1024) throw new Error("File qu\xE1 l\u1EDBn");
    const name = String(p.tenFile || "file.xlsx").replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 120) || "file.xlsx";
    const mime = /\.xlsx$/i.test(name) ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/octet-stream";
    const fid = await ensureFolder();
    const r = await bridge.saveFile(fid, name, d, mime);
    data = { ok: true, url: "https://drive.google.com/file/d/" + r.fileId + "/view" };
  } else if (action === "genDapAnFile") {
    const b = app.getBai(p, u), cfg = b.cauHinh, maDeList = b.maDeList;
    if (!maDeList.length) throw new Error("Ch\u01B0a c\xF3 m\xE3 \u0111\u1EC1");
    const parts = [];
    if (cfg.p1 && cfg.p1.on) parts.push(["p1", cfg.p1.soCau, "Ph\u1EA7n I"]);
    if (cfg.p2 && cfg.p2.on) parts.push(["p2", cfg.p2.soCau, "Ph\u1EA7n II"]);
    if (cfg.p3 && cfg.p3.on) parts.push(["p3", cfg.p3.soCau, "Ph\u1EA7n III"]);
    const thieu = [];
    maDeList.forEach((md) => {
      const da = b.dapAn[md] || {};
      parts.forEach((pt) => {
        const arr = da[pt[0]] || [];
        for (let i = 0; i < pt[1]; i++) {
          const v = arr[i];
          if (v === void 0 || v === null || String(v).trim() === "") thieu.push("M\xE3 " + md + " \xB7 " + pt[2] + " c\xE2u " + (i + 1));
        }
      });
    });
    if (thieu.length) throw new Error("C\xF2n thi\u1EBFu \u0111\xE1p \xE1n: " + thieu.slice(0, 8).join("; ") + (thieu.length > 8 ? " \u2026" : ""));
    const header = ["CAU\\MA DE"].concat(maDeList), rows = [];
    let cau = 1;
    parts.forEach((pt) => {
      for (let i = 0; i < pt[1]; i++) {
        const row = [cau];
        maDeList.forEach((md) => {
          const arr = (b.dapAn[md] || {})[pt[0]] || [];
          let val = arr[i] !== void 0 ? arr[i] : "";
          val = pt[0] === "p3" ? String(val).replace(".", ",") : String(val);
          row.push(val);
        });
        rows.push(row);
        cau++;
      }
    });
    const fid = await ensureFolder();
    const name = "dap-an - " + app.safeName(b.tenBai) + ".xlsx";
    const r = await bridge.genDapAnFile(fid, name, header, rows);
    setCol("dapAnFileId", r.fileId);
    data = { ok: true, fileId: r.fileId, viewUrl: "https://drive.google.com/file/d/" + r.fileId + "/view", downloadUrl: "https://drive.google.com/uc?export=download&id=" + r.fileId, soCau: rows.length, soMaDe: maDeList.length };
  }
  return { ok: true, data };
}
__name(driveAction, "driveAction");
export {
  index_default as default
};
//# sourceMappingURL=index.js.map
