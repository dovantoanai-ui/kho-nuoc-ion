// ============================================================
// BACKEND - App Kho Nuoc Ion HA NOI (DA05)
// Container-bound: mo Google Sheet KHO HA NOI (Sheet MOI, rieng
// hoan toan voi Sheet kho Hung Yen) -> Extensions -> Apps Script
// -> xoa code mau -> dan code nay -> Luu.
// B1: chay ham setupDataTab 1 lan (tao tab Data + 8 SKU mau).
// B2: Deploy -> New deployment -> Web app -> Execute as: Me,
//     Who has access: Anyone -> Deploy -> copy URL /exec.
// B3: mo app hn.html -> tab "Danh muc" -> dan URL /exec -> Luu.
// doGet ho tro: action=catalog (danh muc) va action=report (bao cao).
// ============================================================

var HEAD = {
  NhapKho: ['LineID','PhieuID','Ngay','Nguoi nhan','Nguon','Ma','Ten','DVT','SL','Gia von','Thanh tien','NCC','Ghi chu'],
  XuatKho: ['LineID','PhieuID','Ngay','Loai','Bac','Khach/DL','SDT','Ma','Ten','DVT','SL','Gia ban','Thanh tien','Da thu','Con no','Ghi chu'],
  TraVeHY: ['LineID','PhieuID','Ngay','Nguoi tra','Ma','Ten','DVT','SL','Ghi chu'],
  KiemKe: ['LineID','PhieuID','Ngay','Kho','Ma','Ten','SL dem','Nguoi dem','Ghi chu']
};

// ============================================================
// TU DONG TRU TON THEO HOA DON (16/08/2026)
// App Hoa don ghi vao Sheet DA05 tab "HoaDon" (co cot Kho).
// Backend nay DOC them tab do khi tinh bao cao: hoa don cua
// kho Ha Noi tu ngay HOADON_TU_NGAY duoc cong vao phan Ban
// -> ton kho HN tu tru, khong can nhap phieu Xuat lan 2.
// Ten hang tren hoa don sinh tu tab Data kho HUNG YEN, nen map
// Ten -> Ma dung ca Data kho HY (nguon goc) lan Data kho nay.
// LUU Y VAN HANH: tu ngay moc, don da nhap o app Hoa don thi
// KHONG nhap phieu Xuat o app kho nua (se tru 2 lan).
// ============================================================
var HOADON_SHEET_ID = '1NFWBTjzkgk-Rj6syw46gm5F5tRIDxnU5qBc1OK35Y6A'; // Sheet DA05 (tab HoaDon)
var HY_SHEET_ID     = '1tQEhM51DiYHN3tV_zKa6OT1tLXkPBft4TtOgp1_uAck'; // Sheet kho Hung Yen (nguon ten hang)
var HOADON_TU_NGAY  = '2026-08-16'; // yyyy-mm-dd: chi tinh hoa don tu ngay nay tro di


function doPost(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = { ok: true };
  try {
    var p = JSON.parse(e.postData.getDataAsString());
    if (p.nhap && p.nhap.length) upsert(ss, 'NhapKho', p.nhap);
    if (p.xuat && p.xuat.length) upsert(ss, 'XuatKho', p.xuat);
    if (p.tra && p.tra.length) upsert(ss, 'TraVeHY', p.tra);
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

// Chay 1 lan: tao tab Data (danh muc) + 8 SKU mau.
// Cot "Ton dau" = ton kho HA NOI tai ngay bat dau dung app -> anh dien tay.
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
  // Tao san 4 tab du lieu de anh nhin thay cau truc ngay
  sheetOf(ss, 'NhapKho');
  sheetOf(ss, 'XuatKho');
  sheetOf(ss, 'TraVeHY');
  sheetOf(ss, 'KiemKe');
}

// ============================================================
// doGet: action=catalog (tab Data) hoac action=report (bao cao)
// ============================================================
function doGet(e){
  var cb = (e && e.parameter && e.parameter.callback) || '';
  var action = (e && e.parameter && e.parameter.action) || 'catalog';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out;
  try {
    if (action === 'report') {
      out = { ok: true, cat: readCat_(ss), nhap: readNhap_(ss), xuat: readXuat_(ss).concat(readHoaDonBan_(ss)), tra: readTra_(ss) };
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
  // 0 LineID,1 PhieuID,2 Ngay,3 Nguoi,4 Nguon,5 Ma,6 Ten,7 DVT,8 SL,9 Giavon,10 Thanhtien,11 NCC,12 Ghichu
  for (var i = 1; i < v.length; i++) {
    var r = v[i]; var lid = String(r[0] || ''); if (!lid) continue; if (seen[lid]) continue; seen[lid] = 1;
    out.push({ ngay: fmtYmd_(r[2]), ma: String(r[5]).trim(), sl: n2_(r[8]), tien: n2_(r[10]), nguon: String(r[4] || '') });
  }
  return out;
}

function readXuat_(ss){
  var sh = ss.getSheetByName('XuatKho'); if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(); var out = []; var seen = {};
  // 0 LineID,1 PhieuID,2 Ngay,3 Loai,4 Bac,5 Khach,6 SDT,7 Ma,8 Ten,9 DVT,10 SL,11 Giaban,12 Thanhtien,13 Dathu,14 Conno,15 Ghichu
  for (var i = 1; i < v.length; i++) {
    var r = v[i];
    var lid = String(r[0] || '');
    if (lid) { if (seen[lid]) continue; seen[lid] = 1; }
    if (!lid && !r[10]) continue;
    // Phan loai CHI dua vao tien to app ghi o cot Ten, khong doan tu Ghi chu
    // (kho HN khong ghi tien nen khong the dua vao "thanh tien = 0" nhu ban Hung Yen)
    var ten = String(r[8] || '').toUpperCase();
    var lt = ten.indexOf('[THUHOI]') >= 0 ? 'ThuHoi' : (ten.indexOf('[TANG]') >= 0 ? 'Tang' : 'Ban');
    out.push({ ngay: fmtYmd_(r[2]), ma: String(r[7] || '').trim(), sl: n2_(r[10]), tien: n2_(r[12]), lt: lt, khach: String(r[5] || '') });
  }
  return out;
}

function readTra_(ss){
  var sh = ss.getSheetByName('TraVeHY'); if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(); var out = []; var seen = {};
  // 0 LineID,1 PhieuID,2 Ngay,3 Nguoitra,4 Ma,5 Ten,6 DVT,7 SL,8 Ghichu
  for (var i = 1; i < v.length; i++) {
    var r = v[i]; var lid = String(r[0] || ''); if (!lid) continue; if (seen[lid]) continue; seen[lid] = 1;
    out.push({ ngay: fmtYmd_(r[2]), ma: String(r[4]).trim(), sl: n2_(r[7]), nguoi: String(r[3] || '') });
  }
  return out;
}

// ============================================================
// DOC HOA DON -> QUY VE SKU (tru ton kho HA NOI nhu dong Ban)
// ============================================================
// Chuan hoa ten hang de so khop: bo dau, chu thuong, gon khoang trang
function chuanTen_(s){
  s = String(s == null ? '' : s).toLowerCase();
  try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) {}
  return s.replace(/\u0111/g, 'd').replace(/\s+/g, ' ').trim();
}
// Ngay tren hoa don (dd/MM/yyyy hoac Date) -> yyyy-mm-dd
function ymdHoaDon_(v){
  if (v instanceof Date) return Utilities.formatDate(v, 'GMT+7', 'yyyy-MM-dd');
  var s = String(v == null ? '' : v).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return s.slice(0, 10);
}
// Map Ten -> Ma: lay tu Data kho Hung Yen (nguon sinh ten hoa don) truoc,
// roi de Data kho nay ghi de (neu trung ten thi tin kho minh).
function mapTenMa_(ss){
  var tenMa = {};
  try {
    var hy = SpreadsheetApp.openById(HY_SHEET_ID).getSheetByName('Data');
    if (hy && hy.getLastRow() > 1) {
      var v = hy.getDataRange().getValues();
      for (var i = 1; i < v.length; i++) {
        var ma = String(v[i][0] || '').trim(), ten = String(v[i][1] || '').trim();
        if (ma && ten) tenMa[chuanTen_(ten)] = ma;
      }
    }
  } catch (e) {}
  readCat_(ss).forEach(function(c){ if (c.ten) tenMa[chuanTen_(c.ten)] = c.ma; });
  return tenMa;
}
// Doc tab HoaDon cua Sheet DA05: loc Kho = Ha Noi + tu ngay moc,
// map Ten san pham -> Ma hang, tra ve dang dong Ban.
function readHoaDonBan_(ss){
  var out = [];
  try {
    var sh = SpreadsheetApp.openById(HOADON_SHEET_ID).getSheetByName('HoaDon');
    if (!sh || sh.getLastRow() < 2) return out;
    var tenMa = mapTenMa_(ss);
    var v = sh.getDataRange().getValues();
    var seen = {};
    // 0 ID,1 Ngay,2 MaKH,3 Khach,4 SDT,5 Kho,6 San pham,7 DVT,8 SL,9 Don gia,10 Thanh tien
    for (var i = 1; i < v.length; i++) {
      var r = v[i];
      var id = String(r[0] || '').trim(); if (!id) continue;
      if (String(r[5] || '').indexOf('H\u01b0ng') >= 0) continue; // bo hoa don Hung Yen -> con lai la Ha Noi
      var ngay = ymdHoaDon_(r[1]);
      if (!ngay || ngay < HOADON_TU_NGAY) continue;
      var sl = n2_(r[8]); if (!sl) continue;
      var key = id + '|' + String(r[6] || '') + '|' + sl + '|' + n2_(r[9]);
      if (seen[key]) continue; seen[key] = 1; // chong dong trung (bam dong bo nhieu lan)
      out.push({
        ngay: ngay,
        ma: tenMa[chuanTen_(r[6])] || '',
        sl: sl, tien: n2_(r[10]), lt: 'Ban',
        khach: '[HD] ' + String(r[3] || r[2] || '')
      });
    }
  } catch (e) { /* thieu quyen hoac loi mang: bao cao van chay, chi thieu phan hoa don */ }
  return out;
}
// CHAY TAY de kiem tra ket noi hoa don (Run -> xem Logger)
function testDocHoaDon(){
  var rows = readHoaDonBan_(SpreadsheetApp.getActiveSpreadsheet());
  var thieuMa = rows.filter(function(r){ return !r.ma; }).length;
  Logger.log('Doc duoc ' + rows.length + ' dong hoa don kho Ha Noi tu ' + HOADON_TU_NGAY +
             (thieuMa ? (' | CANH BAO: ' + thieuMa + ' dong khong khop duoc Ma hang') : ' | Tat ca khop Ma hang'));
}

// ============================================================
// CHAY TAY 1 LAN (16/08/2026): dong bo tab Data kho HN theo
// tab Data kho HUNG YEN — du 13 SKU, ten co dau y het hoa don.
// GIU NGUYEN cot Ton dau hien co theo Ma hang.
// CACH CHAY: chon ham capNhatDanhMucTuHY tren thanh cong cu -> Run.
// Xong vao app hn.html -> Danh muc -> "Tai lai danh muc tu Google Sheet".
// ============================================================
function capNhatDanhMucTuHY(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cur = ss.getSheetByName('Data');
  var tondau = {};
  if (cur && cur.getLastRow() > 1) {
    var v0 = cur.getDataRange().getValues();
    for (var i = 1; i < v0.length; i++) {
      var m0 = String(v0[i][0] || '').trim();
      if (m0) tondau[m0] = Number(v0[i][9]) || 0;
    }
  }
  var hy = SpreadsheetApp.openById(HY_SHEET_ID).getSheetByName('Data');
  if (!hy || hy.getLastRow() < 2) throw new Error('Khong doc duoc tab Data kho Hung Yen');
  var v = hy.getDataRange().getValues();
  var rows = [];
  for (var j = 1; j < v.length; j++) {
    var r = v[j]; var m = String(r[0] || '').trim(); if (!m) continue;
    rows.push([m, r[1] || '', r[2] || '', r[3] || 0, r[4] || 0, r[5] || 0, r[6] || 0, r[7] || 0, r[8] || 0, tondau[m] || 0]);
  }
  if (!cur) cur = ss.insertSheet('Data');
  cur.clear();
  cur.appendRow(['Ma hang','Ten hang','DVT','Gia von','Gia NPP','Gia C1','Gia C2','Gia C3','Gia le','Ton dau']);
  if (rows.length) cur.getRange(2, 1, rows.length, 10).setValues(rows);
  cur.setFrozenRows(1);
  Logger.log('Da dong bo ' + rows.length + ' SKU tu kho Hung Yen. Ton dau giu nguyen theo Ma hang.');
}
