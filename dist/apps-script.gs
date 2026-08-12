/**
 * Backend Google Apps Script untuk "Kenalan Yuk!" — CDP 2026 UBL
 *
 * CARA PASANG (±10 menit):
 * 1. Buka https://sheets.google.com, buat Spreadsheet baru (nama bebas, misal "CDP 2026 Kenalan").
 * 2. Di spreadsheet itu, buka menu Extensions > Apps Script.
 * 3. Hapus semua kode default di editor, lalu paste SELURUH isi file ini.
 * 4. Klik ikon Save (💾), lalu klik Deploy > New deployment.
 * 5. Klik ikon gerigi di "Select type" > pilih "Web app".
 * 6. Isi:
 *    - Description: bebas
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7. Klik Deploy. Saat diminta izin, klik "Authorize access" lalu pilih akun Google-mu,
 *    klik "Advanced" > "Go to (nama project) (unsafe)" > Allow. Ini aman, itu script buatanmu sendiri.
 * 8. Salin URL yang muncul (formatnya https://script.google.com/macros/s/XXXXXXX/exec).
 * 9. Tempel URL itu ke bagian SCRIPT_URL di form.html dan index.html (ada komentar "GANTI INI").
 *
 * Kalau nanti mau redeploy setelah edit script: Deploy > Manage deployments > ikon pensil > Version: New version > Deploy.
 * (URL-nya akan tetap sama, tidak perlu ganti ulang di form.html/index.html.)
 *
 * CARA RESET DATA (hapus semua jawaban, misal setelah rehearsal):
 * 1. Buka Apps Script editor (Extensions > Apps Script dari spreadsheet-nya).
 * 2. Di dropdown pemilih fungsi (sebelah tombol Run/▶ di toolbar), pilih "resetData".
 * 3. Klik tombol Run (▶). Semua baris jawaban terhapus, header (baris 1) tetap ada.
 * (Fungsi ini TIDAK bisa dipanggil dari luar/dari web — cuma bisa dijalankan manual dari editor ini, jadi aman.)
 */

var SHEET_NAME = 'Responses';
var HEADERS = ['timestamp', 'nama', 'prodi', 'asal', 'citaCita', 'bidangIT', 'satuKata', 'hobi', 'alasanUBL', 'motto', 'excited'];

function doGet(e) {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse_({ ok: true, data: [] });
  }
  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var data = values.map(function (row) {
    var obj = {};
    HEADERS.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
  return jsonResponse_({ ok: true, data: data });
}

function doPost(e) {
  // LockService memastikan tidak ada dua submission yang menimpa baris yang sama
  // saat banyak peserta (ratusan) mengirim jawaban hampir bersamaan.
  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(15000);
  if (!gotLock) {
    return jsonResponse_({ ok: false, error: 'busy' });
  }
  try {
    var body = JSON.parse(e.postData.contents);
    var sheet = getSheet_();
    var row = HEADERS.map(function (h) {
      if (h === 'timestamp') return new Date().toISOString();
      return (body[h] || '').toString().slice(0, 300);
    });
    sheet.appendRow(row);
    return jsonResponse_({ ok: true });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Jalankan fungsi ini manual dari Apps Script editor (pilih "resetData" di dropdown, klik Run)
 * untuk menghapus semua jawaban. Header (baris 1) tidak ikut terhapus.
 */
function resetData() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}
