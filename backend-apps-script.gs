// ============================================================
// BACKEND - App Kho Nuoc Ion Kiem DA05
// Container-bound: mo Google Sheet KHO -> Extensions -> Apps Script
// -> dan code nay -> Luu.
// B1: chay ham setupDataTab 1 lan (tao tab Data + 8 SKU mau).
// B2: Deploy -> New deployment (hoac Manage deployments -> Edit ->
//     New version) -> Web app -> Execute as: Me, Anyone -> Deploy.
// B3: mo app -> "Danh muc" -> dan URL /exec -> Luu ket noi.
// doGet ho tro: action=catalog (danh muc) va action=report (bao cao).
// ============================================================

var HEAD = {
  NhapKho: ['LineID','PhieuID','Ngay','Nguoi nhap','Ma','Ten','DVT','SL','Gia von','Thanh tien','NCC','Ghi chu'],
  XuatKho: ['LineID','PhieuID','Ngay','Loai','Bac','Khach/DL','SDT','Ma','Ten','DVT','SL','Gia ban','Thanh tien','Da thu','Con no','Ghi chu'],
  ThuTien: ['ID','Ngay','Dai ly/Khach','So tien','Hinh thuc','Nguoi thu','Ghi chu'],
  KiemKe: ['LineID','PhieuID','Ngay','Kho','Ma','Ten','SL dem','Nguoi dem','Ghi chu']
};

function doPost(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = { ok: true };
  try {
    var p = JSON.parse(e.postData.getDataAsString());
    if (p.nhap && p.nhap.length) upsert(ss, 'NhapKho', p.nhap);
    if (p.xuat && p.xuat.length) upsert(ss, 'XuatKho', p.xuat);
    if (p.thu && p.thu.length) upsert(ss, 'ThuTien', p.thu);
    if (p.kiemke && p.kiemke.length) upsert(ss, 'KiemKe', p.kiemke);
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function sheetOf(ss, name){
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); }
  if (sh.getLastRow() === 0) { sh.appendRow(HEAD[name]); sh.setFrozenRows(1); }
  return sh;
}

// Ghi de theo LineID (cot A) -> gui lai khong bi trung dong
function upsert(ss, name, rows){
  var sh = sheetOf(ss, name);
  var w = HEAD[name].length;
  var last = sh.getLastRow();
  var idMap = {};
  if (last > 1) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) idMap[String(ids[i][0])] = i + 2;
  }
  rows.forEach(function(r){
    var vals = r.slice(0, w);
    while (vals.length < w) vals.push('');
    var key = String(vals[0]);
    if (idMap[key]) sh.getRange(idMap[key], 1, 1, w).setValues([vals]);
    else { sh.appendRow(vals); idMap[key] = sh.getLastRow(); }
  });
}

// Chay 1 lan: tao tab Data (danh muc) + 8 SKU mau. Cot gia + Ton dau de 0, anh dien sau.
function setupDataTab(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Data');
  if (!sh) sh = ss.insertSheet('Data');
  sh.clear();
  sh.appendRow(['Ma hang','Ten hang','DVT','Gia von','Gia NPP','Gia C1','Gia C2','Gia C3','Gia le','Ton dau']);
  var d = [
    ['BOCOCTHUYTINH','Bo coc thuy tinh ION Fuji','Bo',39793,0,0,0,0,0,0],
    ['CHAI500ML','Nuoc ion kiem chai 500ml','Thung',11028,0,0,0,0,0,0],
    ['CHAILUX350ML','Nuoc ion kiem Luxury 350ml','Thung',8207,0,0,0,0,0,0],
    ['CHAILUX500ML','Nuoc ion kiem Luxury 500ml','Thung',4354,0,0,0,0,0,0],
    ['ION19LVOI','Nuoc ion kiem binh 19l (co voi)','Binh',114,0,0,0,0,0,0],
    ['LOCRO500','Nuoc Suoi Fuji chai 500ml (12 chai)','Loc',7063,0,0,0,0,0,0],
    ['MOCKHOA','Moc khoa Ion Fuji (qua tang)','Chiec',4700,0,0,0,0,0,0],
    ['RO19IVOI','Nuoc Suoi Fuji binh 19L co voi','Binh',140,0,0,0,0,0,0]
  ];
  sh.getRange(2, 1, d.length, 10).setValues(d);
  sh.setFrozenRows(1);
}

// ============================================================
// doGet: action=catalog (tab Data) hoac action=report (bao cao tu Sheet)
// ============================================================
function doGet(e){
  var cb = (e && e.parameter && e.parameter.callback) || '';
  var action = (e && e.parameter && e.parameter.action) || 'catalog';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out;
  try {
    if (action === 'report') {
      out = { ok: true, cat: readCat_(ss), nhap: readNhap_(ss), xuat: readXuat_(ss), thu: readThu_(ss) };
    } else {
      out = { ok: true, cat: readCat_(ss) };
    }
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  var js = JSON.stringify(out);
  if (cb) return ContentService.createTextOutput(cb + '(' + js + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Helpers doc bao cao ----------
function fmtYmd_(v){ if (v instanceof Date) return Utilities.formatDate(v, 'GMT+7', 'yyyy-MM-dd'); return String(v == null ? '' : v).slice(0, 10); }
function n2_(v){ var s = String(v == null ? '' : v).replace(/[.,\s]/g, ''); var x = parseFloat(s); return isNaN(x) ? 0 : x; }

function readCat_(ss){
  var sh = ss.getSheetByName('Data'); var cat = [];
  if (sh && sh.getLastRow() > 1) {
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      var r = v[i]; if (!r[0]) continue;
      cat.push({
        ma: String(r[0]).trim(), ten: String(r[1]).trim(), dvt: String(r[2]).trim(),
        gv: Number(r[3]) || 0, gnpp: Number(r[4]) || 0, gc1: Number(r[5]) || 0,
        gc2: Number(r[6]) || 0, gc3: Number(r[7]) || 0, gle: Number(r[8]) || 0,
        gtondau: Number(r[9]) || 0
      });
    }
  }
  return cat;
}

function readNhap_(ss){
  var sh = ss.getSheetByName('NhapKho'); if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(); var out = []; var seen = {};
  // 0 LineID,1 PhieuID,2 Ngay,3 Nguoi,4 Ma,5 Ten,6 DVT,7 SL,8 Giavon,9 Thanhtien,10 NCC,11 Ghichu
  for (var i = 1; i < v.length; i++) { var r = v[i]; var lid = String(r[0] || ''); if (!lid) continue; if (seen[lid]) continue; seen[lid] = 1;
    out.push({ ngay: fmtYmd_(r[2]), ma: String(r[4]).trim(), sl: n2_(r[7]), tien: n2_(r[9]) }); }
  return out;
}

function readXuat_(ss){
  // Gop XuatKho + KhoHN (ban le HN). KhoHN khong co Ma nen chi vao doanh thu.
  return readXuatTab_(ss, 'XuatKho').concat(readXuatTab_(ss, 'KhoHN'));
}
function readXuatTab_(ss, name){
  var sh = ss.getSheetByName(name); if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(); var out = []; var seen = {};
  // 0 LineID,1 PhieuID,2 Ngay,3 Loai,4 Bac,5 Khach,6 SDT,7 Ma,8 Ten,9 DVT,10 SL,11 Giaban,12 Thanhtien,13 Dathu,14 Conno,15 Ghichu
  for (var i = 1; i < v.length; i++) { var r = v[i];
    var lid = String(r[0] || '');
    if (lid) { if (seen[lid]) continue; seen[lid] = 1; }
    if (!lid && !r[10] && !r[12]) continue;
    var ten = String(r[8] || '').toUpperCase();
    var ghi = String(r[15] || '').toLowerCase();
    var tien = n2_(r[12]);
    var isTang = ten.indexOf('[TANG]') >= 0 || ((ghi.indexOf('tặng') >= 0 || ghi.indexOf('tang') >= 0) && tien === 0);
    var lt = ten.indexOf('[THUHOI]') >= 0 ? 'ThuHoi' : (isTang ? 'Tang' : 'Ban');
    out.push({ ngay: fmtYmd_(r[2]), ma: String(r[7] || '').trim(), sl: n2_(r[10]), tien: tien, lt: lt, khach: String(r[5] || r[8] || ''), no: n2_(r[14]) });
  }
  return out;
}

function readThu_(ss){
  var sh = ss.getSheetByName('ThuTien'); if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(); var out = [];
  // 0 ID,1 Ngay,2 Daily,3 Sotien,4 Hinhthuc,5 Nguoithu,6 Ghichu
  for (var i = 1; i < v.length; i++) { var r = v[i]; if (!r[0]) continue;
    out.push({ ngay: fmtYmd_(r[1]), dl: String(r[2] || ''), tien: n2_(r[3]) }); }
  return out;
}
