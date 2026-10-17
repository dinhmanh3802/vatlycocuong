/**
 * CẤU HÌNH CHUNG
 * ROSTER_SHEET_ID chỉ là file danh sách của giáo viên ĐẦU TIÊN (mẹ/cô Cường).
 * Mỗi giáo viên về sau có file danh sách riêng, lưu trong bảng NguoiDung.
 */

const ROSTER_SHEET_ID = '11S9VfB_83K74Bv9m36QCF8FHTvqXg0s8A_tZWXRG48c';

const DEFAULT_TEMP_PW = 'Matkhau@123';
const SESSION_TTL_HOURS = 12;
const APP_VERSION = '7.0.0';
/** Địa chỉ máy chủ mới (Cloudflare Worker) — dùng để xác nhận lời gọi cầu nối. Để trống nếu chưa chuyển. */
var WORKER_URL = '';
/* Đang chạy trên Cloudflare Worker (lớp tương thích đặt G.WORKER_MODE = true) hay trên Apps Script */
function isWorker() { try { return typeof WORKER_MODE !== 'undefined' && !!WORKER_MODE; } catch (e) { return false; } }
const TEN_HE_THONG = 'Vật Lý Cô Lâm Cường';

function _props() { return PropertiesService.getScriptProperties(); }

/*
 * Mở một file Google Sheet tốn khoảng 0,2–0,5 giây. Trong MỘT lượt gọi, mỗi file chỉ mở MỘT lần
 * (trước đây một thao tác có thể mở lại cùng file 5–10 lần).
 */
var _SS_MEMO = {};
var _SYS_ID = null;

function getSystemSheetId() {
  if (_SYS_ID) return _SYS_ID;
  var id = _props().getProperty('SYSTEM_SHEET_ID');
  if (!id) throw new Error('Chưa chạy setup(). Hãy chạy hàm setup() một lần trong trình soạn Apps Script.');
  _SYS_ID = id;
  return id;
}
function openSpreadsheetMemo(id) {
  if (!_SS_MEMO[id]) _SS_MEMO[id] = SpreadsheetApp.openById(id);
  return _SS_MEMO[id];
}
function getSystemSheet() { return openSpreadsheetMemo(getSystemSheetId()); }

/** Mở file danh sách của một giáo viên theo ID */
function openRoster(rosterSheetId) {
  if (!rosterSheetId) throw new Error('Giáo viên chưa được gán file danh sách');
  return openSpreadsheetMemo(rosterSheetId);
}
