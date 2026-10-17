/** ĐIỂM VÀO WEB APP — mọi lượt gọi từ trang web đi qua đây */

function doGet(e) { return handle(e, 'GET'); }
function doPost(e) { return handle(e, 'POST'); }

/*
 * Kết quả trả về: { ok, data, ms, v } hoặc { ok:false, error, code, ms }.
 * code: AUTH (cần đăng nhập lại), FORBIDDEN (không đủ quyền), ERR (lỗi nghiệp vụ, có lời nhắn rõ).
 * Lỗi lập trình (TypeError, …) được ghi vào tab NhatKyLoi để sửa.
 */
function handle(e, method) {
  var t0 = Date.now(), p = {}, action = '';
  try {
    _ROWS_MEMO = {}; _SS_MEMO = {}; _DVER = null; _KHONG_DOI = false;     // mỗi lượt gọi bắt đầu sạch
    if (method === 'POST' && e && e.postData && e.postData.contents) p = JSON.parse(e.postData.contents);
    else if (e) p = e.parameter || {};
    if (p.bridge !== undefined) return jsonOut(bridgeHandle(p));          // cầu nối từ Worker (xác thực bằng khóa bí mật)
    action = String(p.action || '');
    if (action !== 'ping') ensureSchema();
    var data = route(action, p);
    if (!SERVER_READ[action] && !_KHONG_DOI) { try { dataTouch(); } catch (e) {} }        // thao tác ghi: đổi phiên bản dữ liệu
    return jsonOut({ ok: true, data: data, ms: Date.now() - t0, v: APP_VERSION });
  } catch (err) {
    var msg = String((err && err.message) || err);
    var code = (err && err.code) || (/Phiên đã hết hạn/.test(msg) ? 'AUTH' : (/Không đủ quyền|Không có quyền/.test(msg) ? 'FORBIDDEN' : 'ERR'));
    if (err && err.name && err.name !== 'Error') {
      ghiLoi(action, p, err);
      msg = 'Lỗi hệ thống khi xử lý (' + action + '). Đã ghi lại để sửa. Chi tiết: ' + msg;
    }
    return jsonOut({ ok: false, error: msg, code: code, ms: Date.now() - t0, v: APP_VERSION });
  }
}

/** Lệnh CHỈ ĐỌC (không đổi dữ liệu). Lệnh không có trong đây đều được coi là ghi: an toàn khi quên khai báo. */
var SERVER_READ = { ping: 1, me: 1, bootstrap: 1, listTeachers: 1, trangThaiTuDong: 1, nhatKyLoi: 1, validateRoster: 1, getBai: 1, publicConfig: 1, traDiem: 1, chiTietBai: 1, listBai: 1, getBaiFull: 1, getKetQua: 1,
  thongKeBai: 1, listGroups: 1, listStudents: 1, listUsers: 1, getCaiDat: 1, listGopY: 1, getNhatKy: 1, previewDiem: 1,
  ddTongQuan: 1, ddPhanTich: 1, ddXemBuoi: 1, ngoaiDoiChieu: 1 };

function route(action, p) {
  switch (action) {
    /* ---- công khai, không cần đăng nhập ---- */
    case 'ping':             return { time: nowIso(), v: APP_VERSION };
    case 'login':            return apiLogin(p);
    case 'quenMatKhau':      return quenMatKhau(p);
    case 'datLaiMatKhau':    return datLaiMatKhau(p);
    case 'publicConfig':     return publicConfig();
    case 'traDiem':          return traDiem(p);
    case 'chiTietBai':       return chiTietBai(p);
    case 'guiGopY':          return guiGopY(p, optionalUser(p));

    /* ---- tài khoản ---- */
    case 'logout':           return apiLogout(p);
    case 'me':               return { user: currentUser(p) };
    case 'changePassword':   return apiChangePassword(p);
    case 'bootstrap':        return requireAuth(p, function (u) { return bootstrap(p, u); });

    /* ---- quản trị ---- */
    case 'listUsers':        return requireRole(p, 'admin', function () { return listUsers(); });
    case 'listTeachers':     return requireRole(p, 'admin', function () { return listTeachers(); });
    case 'createUser':       return requireRole(p, 'admin', function () { return createUser(p); });
    case 'setTeacherRoster': return requireRole(p, 'admin', function () { return setTeacherRoster(p); });
    case 'setUserInfo':      return requireRole(p, 'admin', function () { return setUserInfo(p); });
    case 'setUserActive':    return requireRole(p, 'admin', function (u) { return setUserActive(p, u); });
    case 'resetPassword':    return requireRole(p, 'admin', function () { return adminResetPassword(p); });
    case 'getCaiDat':        return requireRole(p, 'admin', function () { return getCaiDatAll(); });
    case 'setCaiDat':        return requireRole(p, 'admin', function () { return setCaiDat(p); });
    case 'listGopY':         return requireRole(p, 'admin', function () { return listGopY(); });
    case 'xuLyGopY':         return requireRole(p, 'admin', function () { return xuLyGopY(p); });
    case 'trangThaiTuDong':  return requireRole(p, 'admin', function () { return trangThaiTuDong(); });
    case 'batTuDong':        return requireRole(p, 'admin', function () { return batTuDong(); });
    case 'nhatKyLoi':        return requireRole(p, 'admin', function () { return nhatKyLoi(); });

    /* ---- nhóm, học sinh ---- */
    case 'syncGroups':       return requireStaff(p, function (u) { return apiSyncGroups(p, u); });
    case 'listGroups':       return requireStaff(p, function (u) { return listGroups(u, !!p.fresh); });
    case 'listStudents':     return requireStaff(p, function (u) { return listStudents(p, u); });
    case 'upsertStudent':    return requireStaff(p, function (u) { return upsertStudent(p, u); });
    case 'setStudentStatus': return requireStaff(p, function (u) { return setStudentStatus(p, u); });
    case 'transferStudent':  return requireStaff(p, function (u) { return transferStudent(p, u); });
    case 'validateRoster':   return requireStaff(p, function (u) { return validateRoster(u); });

    /* ---- bài kiểm tra, đáp án ---- */
    case 'createBai':        return requireStaff(p, function (u) { return createBai(p, u); });
    case 'updateBai':        return requireStaff(p, function (u) { return updateBai(p, u); });
    case 'listBai':          return requireStaff(p, function (u) { return listBai(u, p); });
    case 'getBai':           return requireStaff(p, function (u) { return getBai(p, u); });
    case 'getBaiFull':       return requireStaff(p, function (u) { return getBaiFull(p, u); });
    case 'saveDapAn':        return requireStaff(p, function (u) { return saveDapAn(p, u); });
    case 'genDapAnFile':     return requireStaff(p, function (u) { return genDapAnFile(p, u); });
    case 'luuTruBai':        return requireStaff(p, function (u) { return luuTruBai(p, u); });
    case 'xoaBai':           return requireStaff(p, function (u) { return xoaBai(p, u); });
    case 'luuFileDrive':     return requireStaff(p, function (u) { return luuFileDrive(p, u); });

    /* ---- ảnh phiếu ---- */
    case 'anhInfo':          return requireStaff(p, function (u) { return anhInfo(p, u); });
    case 'uploadAnh':        return requireStaff(p, function (u) { return uploadAnh(p, u); });
    case 'clearAnh':         return requireStaff(p, function (u) { return clearAnh(p, u); });
    case 'finishAnh':        return requireStaff(p, function (u) { return finishAnh(p, u); });

    /* ---- điểm ---- */
    case 'previewDiem':      return requireStaff(p, function (u) { return previewDiem(p, u); });
    case 'commitDiem':       return requireStaff(p, function (u) { return commitDiem(p, u); });
    case 'getKetQua':        return requireStaff(p, function (u) { return getKetQua(p, u); });
    case 'editKetQua':       return requireStaff(p, function (u) { return editKetQua(p, u); });
    case 'xoaKetQua':        return requireStaff(p, function (u) { return xoaKetQua(p, u); });
    case 'setCongBo':        return requireStaff(p, function (u) { return setCongBo(p, u); });
    case 'getNhatKy':        return requireStaff(p, function (u) { return getNhatKy(p, u); });
    case 'giaiQuyetChoXuLy': return requireStaff(p, function (u) { return giaiQuyetChoXuLy(p, u); });
    case 'thongKeBai':       return requireStaff(p, function (u) { return thongKeBai(p, u); });

    /* ---- điểm danh (lớp trưởng dùng được: quyền kiểm tra bên trong) ---- */
    case 'ddTongQuan':       return requireAuth(p, function (u) { return ddTongQuan(p, u); });
    case 'ddMoBuoi':         return requireAuth(p, function (u) { return ddMoBuoi(p, u); });
    case 'ddXemBuoi':        return requireAuth(p, function (u) { return ddXemBuoi(p, u); });
    case 'ddLuu':            return requireAuth(p, function (u) { return ddLuu(p, u); });
    case 'ddNop':            return requireAuth(p, function (u) { return ddNop(p, u); });
    case 'ddYeuCau':         return requireAuth(p, function (u) { return ddYeuCau(p, u); });
    case 'ddChot':           return requireStaff(p, function (u) { return ddChot(p, u); });
    case 'ddXoaBuoi':        return requireStaff(p, function (u) { return ddXoaBuoi(p, u); });
    case 'ddDuyet':          return requireStaff(p, function (u) { return ddDuyet(p, u); });
    case 'ddPhanTich':       return requireStaff(p, function (u) { return ddPhanTich(p, u); });
    case 'ltTao':            return requireStaff(p, function (u) { return ltTao(p, u); });
    case 'ltSua':            return requireStaff(p, function (u) { return ltSua(p, u); });
    case 'ddBaoSai':         return ddYeuCau(p, null);
    case 'ngoaiDoiChieu':    return requireStaff(p, function (u) { return ngoaiDoiChieu(p, u); });
    case 'ngoaiLuu':         return requireStaff(p, function (u) { return ngoaiLuu(p, u); });
    case 'ngoaiSuaDiem':     return requireStaff(p, function (u) { return ngoaiSuaDiem(p, u); });

    default: throw new Error('Hành động không hợp lệ: ' + action);
  }
}

function requireAuth(p, fn) { return fn(currentUser(p)); }
/** Chức năng của giáo viên / quản trị (tài khoản lớp trưởng không dùng được) */
function requireStaff(p, fn) {
  var u = currentUser(p);
  if (u.role !== 'admin' && u.role !== 'teacher') { var e = new Error('Không đủ quyền'); e.code = 'FORBIDDEN'; throw e; }
  return fn(u);
}
function requireRole(p, role, fn) {
  var u = currentUser(p);
  if (u.role !== role) { var e = new Error('Không đủ quyền'); e.code = 'FORBIDDEN'; throw e; }
  return fn(u);
}
