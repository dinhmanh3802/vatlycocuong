/** TỔ CHỨC THƯ MỤC DRIVE */

function getRootFolder() {
  var p = _props();
  var id = p.getProperty('DRIVE_ROOT_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('HỆ THỐNG ĐIỂM - Lớp học thêm');
  p.setProperty('DRIVE_ROOT_ID', f.getId());
  return f;
}

function getTeacherFolder(username) {
  var p = _props();
  var key = 'DRIVEGV_' + username;
  var id = p.getProperty(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var root = getRootFolder();
  var name = 'GV - ' + username;
  var it = root.getFoldersByName(name);
  var f = it.hasNext() ? it.next() : root.createFolder(name);
  p.setProperty(key, f.getId());
  return f;
}

function safeName(s) {
  return String(s || '').replace(/[\\/:*?"<>|]/g, '-').trim().substring(0, 80) || 'Bài';
}

function createBaiFolder(username, tenBai, ngay) {
  var tf = getTeacherFolder(username);
  var name = (ngay ? ngay + ' - ' : '') + safeName(tenBai);
  return tf.createFolder(name);
}

/** Thư mục con "anh" của một bài (Giai đoạn 3) */
function getAnhFolder(baiFolderId) {
  var folder = DriveApp.getFolderById(baiFolderId);
  var it = folder.getFoldersByName('anh');
  return it.hasNext() ? it.next() : folder.createFolder('anh');
}

/** Xuất một Google Sheet thành blob .xlsx */
function exportSpreadsheetAsXlsxBlob(ssId, filename) {
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx';
  var res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Không xuất được file xlsx (' + res.getResponseCode() + ')');
  return res.getBlob().setName(filename);
}

/** Lưu blob vào thư mục, thay file cùng tên nếu có */
function saveBlobToFolder(folder, blob, name) {
  var it = folder.getFilesByName(name);
  while (it.hasNext()) it.next().setTrashed(true);
  return folder.createFile(blob.setName(name));
}
