// ============================================================
// BACKEND - App Kho Nuoc Ion Kiem DA05
// Container-bound: mo Google Sheet KHO (moi, rieng) -> Extensions
// -> Apps Script -> dan code nay -> Luu.
// B1: chay ham setupDataTab 1 lan (tao tab Data + 8 SKU mau).
// B2: Deploy -> New deployment -> Web app -> Execute as: Me,
//     Who has access: Anyone -> Deploy -> copy URL ket thuc /exec.
// B3: mo app -> man "Danh muc" -> dan URL /exec + Sheet ID (doan
//     trong link Sheet giua /d/ va /edit) -> Luu ket noi.
// ============================================================

var HEAD = {
  NhapKho: ['LineID','PhieuID','Ngay','Nguoi nhap','Ma','Ten','DVT','SL','Gia von','Thanh tien','NCC','Ghi chu'],
  XuatKho: ['LineID','PhieuID','Ngay','Loai','Bac','Khach/DL','SDT','Ma','Ten','DVT','SL','Gia ban','Thanh tien','Da thu','Con no','Ghi chu'],
  ThuTien: ['ID','Ngay','Dai ly/Khach','So tien','Hinh thuc','Nguoi thu','Ghi chu']
};

function doPost(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = { ok: true };
  try {
    var p = JSON.parse(e.postData.getDataAsString());
    if (p.nhap && p.nhap.length) upsert(ss, 'NhapKho', p.nhap);
    if (p.xuat && p.xuat.length) upsert(ss, 'XuatKho', p.xuat);
    if (p.thu && p.thu.length) upsert(ss, 'ThuTien', p.thu);
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

// Chay 1 lan: tao tab Data (danh muc) + 8 SKU mau. Cot gia de 0, anh dien sau.
function setupDataTab(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Data');
  if (!sh) sh = ss.insertSheet('Data');
  sh.clear();
  sh.appendRow(['Ma hang','Ten hang','DVT','Gia von','Gia NPP','Gia C1','Gia C2','Gia C3','Gia le']);
  var d = [
    ['BOCOCTHUYTINH','Bo coc thuy tinh ION Fuji','Bo',39793,0,0,0,0,0],
    ['CHAI500ML','Nuoc ion kiem chai 500ml','Thung',11028,0,0,0,0,0],
    ['CHAILUX350ML','Nuoc ion kiem Luxury 350ml','Thung',8207,0,0,0,0,0],
    ['CHAILUX500ML','Nuoc ion kiem Luxury 500ml','Thung',4354,0,0,0,0,0],
    ['ION19LVOI','Nuoc ion kiem binh 19l (co voi)','Binh',114,0,0,0,0,0],
    ['LOCRO500','Nuoc Suoi Fuji chai 500ml (12 chai)','Loc',7063,0,0,0,0,0],
    ['MOCKHOA','Moc khoa Ion Fuji (qua tang)','Chiec',4700,0,0,0,0,0],
    ['RO19IVOI','Nuoc Suoi Fuji binh 19L co voi','Binh',140,0,0,0,0,0]
  ];
  sh.getRange(2, 1, d.length, 9).setValues(d);
  sh.setFrozenRows(1);
}

// ============================================================
// doGet: tra danh muc (tab Data) cho app qua JSONP.
// App goi: {URL}/exec?action=catalog&callback=xxx -> chay duoc
// ca khi Sheet de rieng tu (script chay bang quyen chu Sheet).
// ============================================================
function doGet(e){
  var cb = (e && e.parameter && e.parameter.callback) || '';
  var out;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Data');
    var cat = [];
    if (sh && sh.getLastRow() > 1) {
      var v = sh.getDataRange().getValues();
      for (var i = 1; i < v.length; i++) {
        var r = v[i];
        if (!r[0]) continue;
        cat.push({
          ma: String(r[0]).trim(), ten: String(r[1]).trim(), dvt: String(r[2]).trim(),
          gv: Number(r[3]) || 0, gnpp: Number(r[4]) || 0, gc1: Number(r[5]) || 0,
          gc2: Number(r[6]) || 0, gc3: Number(r[7]) || 0, gle: Number(r[8]) || 0
        });
      }
    }
    out = { ok: true, cat: cat };
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  var js = JSON.stringify(out);
  if (cb) return ContentService.createTextOutput(cb + '(' + js + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JSON);
}
