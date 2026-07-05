// ============================================================
// BACKEND cho nut "Dong bo" cua app Hoa don (hoadon.html)
// Ghi hoa don vao Sheet DA05 (tab "HoaDon").
// ------------------------------------------------------------
// CACH DUNG:
//  1. Vao script.google.com -> New project (hoac mo Sheet DA05
//     -> Extensions -> Apps Script) -> dan code nay -> Luu.
//  2. Deploy -> New deployment -> Web app ->
//     Execute as: Me,  Who has access: Anyone -> Deploy.
//  3. Copy URL ket thuc /exec.
//  4. Mo app hoa don tren dien thoai -> bam "Dong bo" ->
//     dan URL /exec do (chi can dan 1 lan).
// Luu y: tai khoan chay script phai co quyen sua Sheet DA05.
// ============================================================

var SHEET_ID = '1NFWBTjzkgk-Rj6syw46gm5F5tRIDxnU5qBc1OK35Y6A';
var HEAD = ['ID hoa don', 'Ngay', 'Khach', 'SDT', 'San pham', 'DVT', 'SL',
            'Don gia', 'Thanh tien', 'VAT %', 'Coc binh', 'Tong thanh toan',
            'Trang thai TT', 'Nguoi thu'];

function doPost(e) {
  var out = { ok: true };
  try {
    var p = JSON.parse(e.postData.getDataAsString());
    if (p.rows && p.rows.length) {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var sh = ss.getSheetByName('HoaDon');
      if (!sh) sh = ss.insertSheet('HoaDon');
      if (sh.getLastRow() === 0) { sh.appendRow(HEAD); sh.setFrozenRows(1); }
      p.rows.forEach(function (r) { sh.appendRow(r); });
      out.added = p.rows.length;
    }
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// Chay tay 1 lan de cap quyen truy cap Sheet (neu deploy bao thieu quyen)
function capQuyen() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log('OK: ' + ss.getName());
}
