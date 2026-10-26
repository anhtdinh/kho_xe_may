/**
 * KHO XE MÁY - Backend (Google Apps Script) - bản có đăng nhập Google
 * Gắn script này vào Google Sheet (Extensions > Apps Script).
 * Sheet tự tạo 3 tab: "TonKho" (trạng thái từng xe), "NhatKy" (lịch sử) và "NguoiDung" (danh sách email được phép).
 */

// ====== CẦN SỬA 2 DÒNG NÀY ======
const CLIENT_ID = 'DAN_CLIENT_ID_VAO_DAY.apps.googleusercontent.com'; // OAuth Client ID tạo trong Google Cloud Console
const FOLDER_ID = 'DAN_ID_THU_MUC_DRIVE_VAO_DAY';                      // ID thư mục Drive chứa ảnh
// =================================

const TZ = 'Asia/Ho_Chi_Minh';
const LOG_HEADERS = ['Thời gian', 'Loại', 'Mã QR', 'Số khung', 'Ảnh số khung', 'Ghi chú', 'Người thao tác'];
const STOCK_HEADERS = ['Mã QR', 'Số khung', 'Trạng thái', 'Ngày nhập', 'Ngày xuất', 'Ảnh số khung', 'Ghi chú xuất'];
const USER_HEADERS = ['Email', 'Vai trò', 'Ghi chú'];
const IN_STOCK = 'Trong kho';
const OUT_STOCK = 'Đã xuất';

/** Chạy hàm này MỘT LẦN trong trình soạn thảo để cấp quyền (Sheets, Drive, gọi Google xác thực). */
function authorize_() {
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo', { muteHttpExceptions: true });
  SpreadsheetApp.getActive();
  DriveApp.getFolderById(FOLDER_ID);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return out_({ ok: true, msg: 'Kho xe máy API đang chạy' });
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    // định dạng văn bản để không mất số 0 đầu của mã QR / số khung
    sh.getRange(2, 1, sh.getMaxRows() - 1, headers.length).setNumberFormat('@');
  }
  return sh;
}

function users_() {
  const existed = SpreadsheetApp.getActive().getSheetByName('NguoiDung');
  const sh = sheet_('NguoiDung', USER_HEADERS);
  if (!existed) {
    // người triển khai script được thêm sẵn với quyền đầy đủ
    const owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
    if (owner) sh.appendRow([owner, 'Cả hai', 'Chủ sở hữu']);
  }
  return sh;
}

function findRow_(sh, col, val) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const v = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim().toUpperCase() === val.toUpperCase()) return i + 2;
  }
  return 0;
}

function now_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function savephoto_(b64, name) {
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name + '.jpg');
  const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
  return file.getUrl();
}

function clean_(s) {
  return String(s || '').trim();
}

/** Vai trò: trống / "Cả hai" = nhập và xuất; "Nhập" = chỉ nhập; "Xuất" = chỉ xuất. */
function perm_(role) {
  const r = String(role || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').trim();
  if (r === '' || r.indexOf('ca') === 0 || r === 'both' || r === 'all') return { in: true, out: true };
  if (r === 'nhap') return { in: true, out: false };
  if (r === 'xuat') return { in: false, out: true };
  return null;
}

/** Kiểm tra mã đăng nhập Google và email có trong danh sách không. */
function auth_(token) {
  if (!token) return { fail: { ok: false, code: 'auth', error: 'Chưa đăng nhập' } };

  const cache = CacheService.getScriptCache();
  const key = 't_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
  let email = cache.get(key);

  if (!email) {
    const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) {
      return { fail: { ok: false, code: 'auth', error: 'Phiên đăng nhập không hợp lệ hoặc đã hết hạn' } };
    }
    const t = JSON.parse(r.getContentText());
    if (t.aud !== CLIENT_ID) {
      return { fail: { ok: false, code: 'auth', error: 'Client ID không khớp. Kiểm tra CLIENT_ID trong Code.gs và trong Cài đặt của app' } };
    }
    if (String(t.email_verified) !== 'true' || !t.email) {
      return { fail: { ok: false, code: 'auth', error: 'Email Google chưa được xác minh' } };
    }
    email = String(t.email).toLowerCase();
    const ttl = Math.min(600, Math.max(1, Number(t.exp) - Math.floor(Date.now() / 1000)));
    cache.put(key, email, ttl);
  }

  const sh = users_();
  const row = findRow_(sh, 1, email);
  if (!row) {
    return { fail: { ok: false, code: 'auth', error: 'Email ' + email + ' chưa được cấp quyền. Nhờ quản lý thêm email này vào tab NguoiDung.' } };
  }
  const p = perm_(sh.getRange(row, 2).getValue());
  if (!p) {
    return { fail: { ok: false, code: 'auth', error: 'Vai trò của ' + email + ' không hợp lệ. Dùng: Nhập, Xuất hoặc Cả hai.' } };
  }
  return { email: email, canIn: p.in, canOut: p.out };
}

function doPost(e) {
  let lock;
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'ping') return out_({ ok: true });

    const a = auth_(d.token);
    if (a.fail) return out_(a.fail);
    if (d.action === 'whoami') return out_({ ok: true, email: a.email, canIn: a.canIn, canOut: a.canOut });

    lock = LockService.getScriptLock();
    lock.waitLock(20000);

    const stock = sheet_('TonKho', STOCK_HEADERS);
    const log = sheet_('NhatKy', LOG_HEADERS);
    const qr = clean_(d.qr);
    const user = a.email;
    if (!qr) return out_({ ok: false, error: 'Thiếu mã QR' });

    const row = findRow_(stock, 1, qr);

    // ---- Tra cứu một xe (dùng cho xuất kho) ----
    if (d.action === 'lookup') {
      if (!a.canOut) return out_({ ok: false, error: 'Tài khoản của bạn không có quyền xuất kho' });
      if (!row) return out_({ ok: true, found: false });
      const r = stock.getRange(row, 1, 1, STOCK_HEADERS.length).getValues()[0];
      return out_({ ok: true, found: true, vin: r[1], status: r[2], dateIn: r[3], dateOut: r[4] });
    }

    // ---- Nhập kho ----
    if (d.action === 'in') {
      if (!a.canIn) return out_({ ok: false, error: 'Tài khoản của bạn không có quyền nhập kho' });
      const vin = clean_(d.vin).toUpperCase();
      if (!vin) return out_({ ok: false, error: 'Thiếu số khung' });
      if (!d.photo) return out_({ ok: false, error: 'Thiếu ảnh số khung' });
      if (row) {
        const st = stock.getRange(row, 3).getValue();
        return out_({ ok: false, error: 'Tem QR này đã được nhập trước đó (' + st + ')' });
      }
      const vinRow = findRow_(stock, 2, vin);
      if (vinRow) {
        const otherQr = stock.getRange(vinRow, 1).getValue();
        return out_({ ok: false, error: 'Số khung này đã gắn với tem QR ' + otherQr });
      }
      const url = savephoto_(d.photo, vin + '_' + qr + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss'));
      const t = now_();
      stock.appendRow([qr, vin, IN_STOCK, t, '', url, '']);
      log.appendRow([t, 'Nhập', qr, vin, url, '', user]);
      return out_({ ok: true, vin: vin });
    }

    // ---- Xuất kho ----
    if (d.action === 'out') {
      if (!a.canOut) return out_({ ok: false, error: 'Tài khoản của bạn không có quyền xuất kho' });
      if (!row) return out_({ ok: false, error: 'Tem QR này chưa được nhập kho' });
      const r = stock.getRange(row, 1, 1, STOCK_HEADERS.length).getValues()[0];
      if (r[2] !== IN_STOCK) return out_({ ok: false, error: 'Xe này đã xuất kho lúc ' + r[4] });
      const note = clean_(d.note);
      const t = now_();
      let url = '';
      if (d.photo) url = savephoto_(d.photo, r[1] + '_' + qr + '_XUAT_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss'));
      stock.getRange(row, 3).setValue(OUT_STOCK);
      stock.getRange(row, 5).setValue(t);
      stock.getRange(row, 7).setValue(note);
      log.appendRow([t, 'Xuất', qr, r[1], url, note, user]);
      return out_({ ok: true, vin: r[1] });
    }

    return out_({ ok: false, error: 'Hành động không hợp lệ' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    if (lock) lock.releaseLock();
  }
}
