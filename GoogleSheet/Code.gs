/**
 * DANG KY DAO TAO KY NANG LAI XE AN TOAN - XA VINH HAU (Google Sheets + Apps Script)
 * Trang tinh "DangKy": moi nguoi dang ky la 1 dong.
 * Trang web (GitHub Pages) gui POST JSON {"fn": "...", "args": [...]}.
 */

// Ma de XEM / TAI danh sach (trang quan ly). Nguoi dang ky KHONG can ma.
// Doi thanh ma rieng cua thay truoc khi trien khai.
var MA_QUAN_LY = 'vinhhau2026';

var SHEET_DK = 'DangKy';
var HEAD = ['Mã', 'Họ và tên', 'Ngày sinh', 'Địa chỉ', 'Số GPLX', 'Số điện thoại', 'Thời gian đăng ký', 'Cập nhật lúc'];

function doGet() {
  return ContentService.createTextOutput('API dang ky lai xe an toan dang chay. Mo trang dang ky de dien.');
}

function doPost(e) {
  var api = { dem: dem, dangKy: dangKy, danhSach: danhSach, xoa: xoa };
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!api.hasOwnProperty(req.fn)) throw new Error('Yêu cầu không hợp lệ.');
    var args = Array.isArray(req.args) ? req.args : [];
    out = { ok: true, result: api[req.fn].apply(null, args) };
  } catch (err) {
    out = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Đăng ký lái xe')
    .addItem('Sắp xếp theo thời gian đăng ký', 'sapXep')
    .addToUi();
}

// ---------- Tien ich ----------
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_DK);
  if (!sh) {
    sh = ss.insertSheet(SHEET_DK);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#dbeafe');
    sh.setFrozenRows(1);
  }
  return sh;
}

function safe_(v, max) {
  var s = String(v === null || v === undefined ? '' : v).replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (s.length > max) s = s.slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function plain_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/\s+/g, ' ').trim();
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
}

function checkMa_(ma) {
  if (String(ma || '') !== MA_QUAN_LY) {
    Utilities.sleep(800);
    throw new Error('Sai mã quản lý.');
  }
}

function rows_() {
  var sh = sheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEAD.length).getDisplayValues();
}

// ---------- API ----------
function dem() {
  return rows_().filter(function (r) { return r[1]; }).length;
}

/** d = {ten, ngaySinh, diaChi, gplx, sdt} */
function dangKy(d) {
  d = d || {};
  var ten = safe_(d.ten, 80);
  var ns = safe_(d.ngaySinh, 20);
  var dc = safe_(d.diaChi, 200);
  var gp = safe_(d.gplx, 40);
  var sdt = String(d.sdt || '').replace(/[^\d+]/g, '');
  if (ten.length < 4) throw new Error('Vui lòng nhập đầy đủ họ và tên.');
  if (!/^\d{2}\/\d{2}\/\d{4}$/.test(ns)) throw new Error('Ngày sinh chưa đúng (dd/mm/yyyy).');
  if (dc.length < 3) throw new Error('Vui lòng nhập địa chỉ.');
  if (!/^(0|\+84)\d{9,10}$/.test(sdt)) throw new Error('Số điện thoại chưa đúng.');
  if (!gp) gp = 'Chưa có';

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_();
    var data = rows_();
    var key = plain_(ten) + '|' + sdt;
    for (var i = 0; i < data.length; i++) {
      if (plain_(data[i][1]) + '|' + String(data[i][5]).replace(/[^\d+]/g, '') === key) {
        sh.getRange(i + 2, 2, 1, 5).setValues([[ten, "'" + ns, dc, gp, "'" + sdt]]);
        sh.getRange(i + 2, 8).setValue("'" + now_());
        return { capNhat: true, stt: i + 1 };
      }
    }
    var ma = 'DK' + Utilities.getUuid().slice(0, 8);
    // dau ' giu nguyen so 0 dau SDT va ngay sinh dang chu (khong bi doi thanh ngay kieu My)
    sh.appendRow([ma, ten, "'" + ns, dc, gp, "'" + sdt, "'" + now_(), '']);
    return { capNhat: false, stt: data.length + 1 };
  } finally {
    lock.releaseLock();
  }
}

function danhSach(ma) {
  checkMa_(ma);
  return rows_().filter(function (r) { return r[1]; }).map(function (r) {
    return { ma: r[0], ten: r[1], ngaySinh: r[2], diaChi: r[3], gplx: r[4], sdt: r[5], luc: r[6] };
  });
}

function xoa(ma, id) {
  checkMa_(ma);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data = rows_();
    for (var i = data.length - 1; i >= 0; i--) {
      if (data[i][0] === String(id)) { sheet_().deleteRow(i + 2); return true; }
    }
    return false;
  } finally {
    lock.releaseLock();
  }
}

function sapXep() {
  var sh = sheet_();
  var n = sh.getLastRow() - 1;
  if (n < 2) return;
  var r = sh.getRange(2, 1, n, HEAD.length);
  var v = r.getDisplayValues();
  function t(s) { var m = String(s).match(/(\d+)\/(\d+)\/(\d+) (\d+):(\d+)/); return m ? +(m[3] + m[2] + m[1] + m[4] + m[5]) : 0; }
  v.sort(function (a, b) { return t(a[6]) - t(b[6]); });
  r.setValues(v.map(function (x) {
    [2, 5, 6, 7].forEach(function (k) { if (x[k]) x[k] = "'" + x[k]; });
    return x;
  }));
}
