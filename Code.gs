/**
 * KHO XE MÁY - Backend (Google Apps Script)
 * Gắn script này vào Google Sheet (Extensions > Apps Script).
 * Sheet tự tạo 2 tab: "TonKho" (trạng thái từng xe) và "NhatKy" (lịch sử nhập/xuất).
 */

// ====== CẦN SỬA 2 DÒNG NÀY ======
const FOLDER_ID = 'DAN_ID_THU_MUC_DRIVE_VAO_DAY'; // ID thư mục Drive chứa ảnh (đoạn cuối của link thư mục)
const SECRET = 'DOI_MAT_KHAU_NAY';                 // phải trùng với "Mật khẩu app" trong phần Cài đặt của app
// =================================

const TZ = 'Asia/Ho_Chi_Minh';
const LOG_HEADERS = ['Thời gian', 'Loại', 'Mã QR', 'Số khung', 'Ảnh số khung', 'Ghi chú', 'Người thao tác'];
const STOCK_HEADERS = ['Mã QR', 'Số khung', 'Trạng thái', 'Ngày nhập', 'Ngày xuất', 'Ảnh số khung', 'Ghi chú xuất'];
const IN_STOCK = 'Trong kho';
const OUT_STOCK = 'Đã xuất';

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

function doPost(e) {
  let lock;
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.secret !== SECRET) return out_({ ok: false, error: 'Sai mật khẩu app' });
    if (d.action === 'ping') return out_({ ok: true });

    lock = LockService.getScriptLock();
    lock.waitLock(20000);

    const stock = sheet_('TonKho', STOCK_HEADERS);
    const log = sheet_('NhatKy', LOG_HEADERS);
    const qr = clean_(d.qr);
    const user = clean_(d.user);
    if (!qr) return out_({ ok: false, error: 'Thiếu mã QR' });

    const row = findRow_(stock, 1, qr);

    // ---- Tra cứu một xe ----
    if (d.action === 'lookup') {
      if (!row) return out_({ ok: true, found: false });
      const r = stock.getRange(row, 1, 1, STOCK_HEADERS.length).getValues()[0];
      return out_({ ok: true, found: true, vin: r[1], status: r[2], dateIn: r[3], dateOut: r[4] });
    }

    // ---- Nhập kho ----
    if (d.action === 'in') {
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
