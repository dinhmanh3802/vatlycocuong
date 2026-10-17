/** ĐĂNG NHẬP, PHIÊN, TÀI KHOẢN, GIÁO VIÊN */

var LOGIN_MAX_FAIL = 5;         // sai 5 lần liên tiếp thì khóa tạm
var LOGIN_LOCK_MIN = 15;        // trong 15 phút

function _users() { return getSystemSheet().getSheetByName('NguoiDung'); }
function _phien() { return getSystemSheet().getSheetByName('Phien'); }

function findUserRow(username) {
  var sh = _users();
  var rows = sh.getDataRange().getValues();
  var H = idx(rows[0]);
  var want = String(username || '').trim().toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][H['username']]).trim().toLowerCase() === want) return { row: i + 1, data: rows[i], H: H, sheet: sh };
  }
  return null;
}
function userField(u, k) { return u.H[k] === undefined ? '' : u.data[u.H[k]]; }

function authError(msg) { var e = new Error(msg); e.code = 'AUTH'; return e; }

/* ---- Thu hồi phiên: đổi/đặt lại mật khẩu hoặc khóa tài khoản thì mọi phiên cũ hết hiệu lực ngay ---- */
/* Thứ tự "cấp phiên → thu hồi → cấp phiên" luôn tăng dần, kể cả khi rơi cùng một mili-giây:
   thu hồi không sớm hơn phiên gần nhất đã cấp; phiên mới luôn sau lần thu hồi gần nhất. */
function issuedAt(username) {
  var k = 'ISS_' + String(username).toLowerCase(), c = CacheService.getScriptCache(), v = c.get(k);
  if (v === null || v === undefined) { v = _props().getProperty(k) || '0'; c.put(k, v, 21600); }
  return Number(v) || 0;
}
function setIssued(username, t) {
  var k = 'ISS_' + String(username).toLowerCase();
  _props().setProperty(k, String(t)); CacheService.getScriptCache().put(k, String(t), 21600);
}
function revokeSessions(username) {
  var k = 'REV_' + String(username).toLowerCase(), v = String(Math.max(Date.now(), issuedAt(username)));
  _props().setProperty(k, v);
  CacheService.getScriptCache().put(k, v, 21600);
}
function revokedAt(username) {
  var k = 'REV_' + String(username).toLowerCase(), c = CacheService.getScriptCache(), v = c.get(k);
  if (v === null || v === undefined) { v = _props().getProperty(k) || '0'; c.put(k, v, 21600); }
  return Number(v) || 0;
}

function apiLogin(p) {
  var username = String(p.username || '').trim();
  var pw = String(p.password || '');
  if (!username || !pw) throw new Error('Nhập tài khoản và mật khẩu');
  var cache = CacheService.getScriptCache(), fk = 'lf_' + username.toLowerCase();
  var fails = Number(cache.get(fk)) || 0;
  if (fails >= LOGIN_MAX_FAIL) throw new Error('Đăng nhập sai quá ' + LOGIN_MAX_FAIL + ' lần. Thử lại sau ' + LOGIN_LOCK_MIN + ' phút, hoặc dùng "Quên mật khẩu".');
  var u = findUserRow(username);
  var dungMk = u && hashPassword(pw, userField(u, 'salt')) === userField(u, 'passHash');
  if (!dungMk) {
    cache.put(fk, String(fails + 1), LOGIN_LOCK_MIN * 60);
    var con = LOGIN_MAX_FAIL - fails - 1;
    throw new Error('Sai tài khoản hoặc mật khẩu' + (con > 0 && con <= 2 ? ' (còn ' + con + ' lần thử)' : ''));
  }
  if (!toBool(userField(u, 'active'))) throw new Error('Tài khoản đã bị khóa. Liên hệ quản trị.');
  cache.remove(fk);
  var uname = String(userField(u, 'username')), role = userField(u, 'role'), fullName = userField(u, 'fullName');
  var token = createSession(uname, role, fullName);
  return { token: token, user: { username: uname, role: role, fullName: fullName, mustChange: toBool(userField(u, 'mustChange')) } };
}

function createSession(username, role, fullName) {
  var token = uuid() + uuid().replace(/-/g, '');
  var now = Math.max(Date.now(), revokedAt(username) + 1);     // phiên mới luôn sau lần thu hồi gần nhất, kể cả cùng mili-giây
  setIssued(username, now);
  var exp = new Date(now + SESSION_TTL_HOURS * 3600 * 1000).toISOString();
  var obj = { token: token, username: username, role: role, fullName: fullName, expiry: exp, iat: now };
  _phien().appendRow([token, username, role, fullName, exp]);
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(obj), 6 * 3600);
  return token;
}

function getSession(token) {
  if (!token) return null;
  var cached = CacheService.getScriptCache().get('sess_' + token);
  var s = null;
  if (cached) { s = JSON.parse(cached); }
  else {
    var rows = _phien().getDataRange().getValues();
    var H = idx(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][H['token']]) === token) {
        s = { token: token, username: rows[i][H['username']], role: rows[i][H['role']], fullName: rows[i][H['fullName']], expiry: rows[i][H['expiry']], iat: new Date(rows[i][H['expiry']]).getTime() - SESSION_TTL_HOURS * 3600 * 1000 };
        break;
      }
    }
  }
  if (!s) return null;
  if (new Date(s.expiry).getTime() < Date.now()) return null;
  var rev = revokedAt(s.username);
  if (rev && (s.iat || 0) <= rev) return null;               // tạo trước hoặc đúng lúc thu hồi: hết hiệu lực
  return s;
}

function currentUser(p) {
  var s = getSession(p.token);
  if (!s) throw authError('Phiên đã hết hạn, hãy đăng nhập lại');
  return { username: s.username, role: s.role, fullName: s.fullName };
}
/** Dùng cho chức năng vừa công khai vừa có thể đăng nhập (góp ý): không ném lỗi */
function optionalUser(p) { try { return p.token ? currentUser(p) : null; } catch (e) { return null; } }

function apiLogout(p) {
  if (p.token) {
    CacheService.getScriptCache().remove('sess_' + p.token);
    var sh = _phien();
    var rows = sh.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) if (String(rows[i][0]) === p.token) sh.deleteRow(i + 1);
  }
  return { ok: true };
}

function setPassword(u, newPw, mustChange) {
  var salt = randomSalt();
  u.sheet.getRange(u.row, u.H['salt'] + 1).setValue(salt);
  u.sheet.getRange(u.row, u.H['passHash'] + 1).setValue(hashPassword(newPw, salt));
  u.sheet.getRange(u.row, u.H['mustChange'] + 1).setValue(!!mustChange);
}
function checkNewPassword(pw) {
  pw = String(pw || '');
  if (pw.length < 6) throw new Error('Mật khẩu mới cần ít nhất 6 ký tự');
  if (pw.length > 100) throw new Error('Mật khẩu quá dài');
  return pw;
}

/** Đổi mật khẩu: các máy khác bị đăng xuất, máy đang dùng nhận phiên mới */
function apiChangePassword(p) {
  var u0 = currentUser(p);
  var u = findUserRow(u0.username);
  if (!u) throw new Error('Không tìm thấy tài khoản');
  if (hashPassword(String(p.oldPassword || ''), userField(u, 'salt')) !== userField(u, 'passHash')) throw new Error('Mật khẩu hiện tại không đúng');
  var pw = checkNewPassword(p.newPassword);
  if (pw === String(p.oldPassword || '')) throw new Error('Mật khẩu mới phải khác mật khẩu cũ');
  setPassword(u, pw, false);
  revokeSessions(u0.username);
  Utilities.sleep(5);
  var token = createSession(u0.username, u0.role, u0.fullName);
  return { ok: true, token: token };
}

/* ---- Quên mật khẩu: gửi mã 6 số tới email của tài khoản (quản trị nhập email cho từng tài khoản) ---- */
function quenMatKhau(p) {
  var username = String(p.username || '').trim();
  if (!username) throw new Error('Nhập tên tài khoản');
  var res = { ok: true, msg: 'Nếu tài khoản có email, mã đặt lại đã được gửi tới email đó. Mã có hiệu lực 15 phút.' };
  var cache = CacheService.getScriptCache(), rk = 'rs_' + username.toLowerCase();
  if (cache.get(rk + '_wait')) return res;               // mỗi phút gửi tối đa một lần
  cache.put(rk + '_wait', '1', 60);
  var u = findUserRow(username);
  if (!u || !toBool(userField(u, 'active'))) return res;  // không tiết lộ tài khoản có tồn tại hay không
  var email = String(userField(u, 'email') || '').trim();
  if (!email) return res;
  var code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put(rk, JSON.stringify({ h: sha256Hex(code), n: 0 }), 900);
  try {
    MailApp.sendEmail({ to: email, name: 'Hệ thống điểm', subject: '[Điểm] Mã đặt lại mật khẩu: ' + code,
      body: 'Mã đặt lại mật khẩu cho tài khoản "' + username + '": ' + code + '\n\nMã có hiệu lực 15 phút. Nếu không phải bạn yêu cầu, hãy bỏ qua email này.' });
  } catch (e) {}
  return res;
}

function datLaiMatKhau(p) {
  var username = String(p.username || '').trim();
  var cache = CacheService.getScriptCache(), rk = 'rs_' + username.toLowerCase();
  var st = cache.get(rk);
  if (!st) throw new Error('Mã đã hết hạn hoặc chưa được gửi. Bấm "Gửi mã" lại.');
  st = JSON.parse(st);
  if (st.n >= 5) { cache.remove(rk); throw new Error('Nhập sai mã quá nhiều lần. Hãy gửi mã mới.'); }
  if (sha256Hex(String(p.code || '').trim()) !== st.h) {
    st.n++; cache.put(rk, JSON.stringify(st), 900);
    throw new Error('Mã không đúng');
  }
  var pw = checkNewPassword(p.newPassword);
  var u = findUserRow(username);
  if (!u) throw new Error('Không tìm thấy tài khoản');
  setPassword(u, pw, false);
  cache.remove(rk);
  cache.remove('lf_' + username.toLowerCase());
  revokeSessions(username);
  return { ok: true };
}

/* ---- Dải mã ---- */

function rangesOverlap(aS, aE, bS, bE) { return aS <= bE && bS <= aE; }

function checkRangeFree(rangeStart, rangeEnd, exceptUsername) {
  if (rangeStart === '' || rangeStart === undefined || rangeStart === null || rangeEnd === '' || rangeEnd === undefined || rangeEnd === null) return;
  var s = parseInt(rangeStart, 10), e = parseInt(rangeEnd, 10);
  if (isNaN(s) || isNaN(e) || s < 0 || e > 999 || s > e) throw new Error('Dải mã không hợp lệ (0–999, đầu ≤ cuối)');
  sysRows('NguoiDung').forEach(function (r) {
    if (r.role !== 'teacher') return;
    if (exceptUsername && String(r.username).toLowerCase() === String(exceptUsername).toLowerCase()) return;
    if (r.rangeStart === '' || r.rangeStart === undefined) return;
    if (rangesOverlap(s, e, parseInt(r.rangeStart, 10), parseInt(r.rangeEnd, 10)))
      throw new Error('Dải mã trùng với giáo viên ' + r.username + ' (' + pad3(r.rangeStart) + '–' + pad3(r.rangeEnd) + ')');
  });
}

/* ---- Quản trị tài khoản (admin) ---- */

function listUsers() {
  sysTouch('NguoiDung');
  return {
    users: sysRows('NguoiDung').map(function (r) {
      return {
        username: r.username, role: r.role, fullName: r.fullName, email: String(r.email || ''),
        active: toBool(r.active), mustChange: toBool(r.mustChange),
        rosterSheetId: r.rosterSheetId, rangeStart: r.rangeStart, rangeEnd: r.rangeEnd
      };
    })
  };
}

function listTeachers() {
  var users = sysRows('NguoiDung').filter(function (r) { return r.role === 'teacher'; });
  var nhom = sysRows('Nhom');
  return {
    teachers: users.map(function (u) {
      return {
        username: u.username, fullName: u.fullName, active: toBool(u.active),
        rosterSheetId: u.rosterSheetId, rangeStart: u.rangeStart, rangeEnd: u.rangeEnd,
        groupCount: nhom.filter(function (n) { return n.teacher === u.username; }).length
      };
    })
  };
}

function checkEmail(email) {
  email = String(email || '').trim();
  if (email && !/^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(email)) throw new Error('Email không hợp lệ');
  return email;
}

function createUser(p) {
  var username = String(p.username || '').trim();
  if (!/^[A-Za-z0-9._-]{2,40}$/.test(username)) throw new Error('Tên tài khoản 2–40 ký tự, chỉ gồm chữ không dấu, số, dấu chấm, gạch');
  if (findUserRow(username)) throw new Error('Tài khoản đã tồn tại');
  var role = p.role === 'admin' ? 'admin' : 'teacher';
  var rangeStart = '', rangeEnd = '';
  if (role === 'teacher') {
    rangeStart = (p.rangeStart === undefined || p.rangeStart === '') ? '' : parseInt(p.rangeStart, 10);
    rangeEnd = (p.rangeEnd === undefined || p.rangeEnd === '') ? '' : parseInt(p.rangeEnd, 10);
    checkRangeFree(rangeStart, rangeEnd, null);
  }
  var email = checkEmail(p.email);
  var sh = _users(), H = idx(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]), nCol = sh.getLastColumn();
  var row = []; for (var i = 0; i < nCol; i++) row.push('');
  var salt = randomSalt();
  var set = function (k, v) { if (H[k] !== undefined) row[H[k]] = v; };
  set('username', username); set('role', role); set('fullName', String(p.fullName || '').trim());
  set('passHash', hashPassword(DEFAULT_TEMP_PW, salt)); set('salt', salt); set('mustChange', true); set('active', true);
  set('rosterSheetId', String(p.rosterSheetId || '').trim()); set('rangeStart', rangeStart); set('rangeEnd', rangeEnd);
  set('createdAt', nowIso()); set('email', email);
  sh.appendRow(row);
  sysTouch();
  return { ok: true, tempPassword: DEFAULT_TEMP_PW };
}

/** Sửa thông tin tài khoản: họ tên, email, file danh sách, dải mã */
function setUserInfo(p) {
  var u = findUserRow(p.username);
  if (!u) throw new Error('Không tìm thấy tài khoản');
  var isTeacher = userField(u, 'role') === 'teacher';
  if (isTeacher && p.rangeStart !== undefined && p.rangeEnd !== undefined) checkRangeFree(p.rangeStart, p.rangeEnd, p.username);
  var put = function (k, v) { if (u.H[k] !== undefined) u.sheet.getRange(u.row, u.H[k] + 1).setValue(v); };
  if (p.fullName !== undefined) put('fullName', String(p.fullName || '').trim());
  if (p.email !== undefined) { if (u.H['email'] !== undefined) setTextCell(u.sheet, u.row, u.H['email'] + 1, checkEmail(p.email)); }
  if (isTeacher && p.rosterSheetId !== undefined) put('rosterSheetId', String(p.rosterSheetId || '').trim());
  if (isTeacher && p.rangeStart !== undefined) put('rangeStart', p.rangeStart === '' ? '' : parseInt(p.rangeStart, 10));
  if (isTeacher && p.rangeEnd !== undefined) put('rangeEnd', p.rangeEnd === '' ? '' : parseInt(p.rangeEnd, 10));
  sysTouch();
  return { ok: true };
}
/** Giữ cho tương thích bản cũ */
function setTeacherRoster(p) { return setUserInfo(p); }

function setUserActive(p, me) {
  var u = findUserRow(p.username);
  if (!u) throw new Error('Không tìm thấy tài khoản');
  if (me && String(me.username).toLowerCase() === String(p.username).toLowerCase() && !p.active) throw new Error('Không thể tự khóa tài khoản đang dùng');
  u.sheet.getRange(u.row, u.H['active'] + 1).setValue(!!p.active);
  if (!p.active) revokeSessions(userField(u, 'username'));
  sysTouch();
  return { ok: true };
}

function adminResetPassword(p) {
  var u = findUserRow(p.username);
  if (!u) throw new Error('Không tìm thấy tài khoản');
  setPassword(u, DEFAULT_TEMP_PW, true);
  revokeSessions(userField(u, 'username'));
  CacheService.getScriptCache().remove('lf_' + String(p.username).toLowerCase());
  return { ok: true, tempPassword: DEFAULT_TEMP_PW };
}
