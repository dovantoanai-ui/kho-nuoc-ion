// ============================================================
// BACKEND cho app Hoa don + CRM (hoadon.html)
// Ghi & doc du lieu tren Sheet DA05: tab "KhachHang" va "HoaDon".
// Co khoa ghi tuan tu (LockService) + backend tu cap ma khach.
// ------------------------------------------------------------
// CACH DUNG:
//  1. script.google.com -> mo project cu -> dan TOAN BO code nay
//     de len ban cu -> Luu.
//  2. Deploy -> Manage deployments -> ban Web app hien tai
//     -> but chi (Edit) -> Version: New version -> Deploy.
//     (Chua co thi: Deploy -> New deployment -> Web app ->
//      Execute as: Me,  Who has access: Anyone -> Deploy.)
//  3. URL /exec giu nguyen. Neu bao thieu quyen: chay tay ham capQuyen 1 lan.
// ============================================================

var SHEET_ID = '1NFWBTjzkgk-Rj6syw46gm5F5tRIDxnU5qBc1OK35Y6A';

var KH_HEAD = ['Ma KH', 'Ten', 'SDT', 'Quan/Huyen', 'Kho',
               'Thanh toan', 'Ky han no', 'Ngay tao'];

var HD_HEAD = ['ID hoa don', 'Ngay', 'Ma KH', 'Khach', 'SDT', 'Kho',
               'San pham', 'DVT', 'SL', 'Don gia', 'Thanh tien',
               'VAT %', 'Coc binh', 'Tong thanh toan', 'Da thu',
               'Trang thai TT', 'Nguoi thu', 'Hinh thuc TT'];

// ---------- GHI (POST tu app, che do no-cors) ----------
function doPost(e) {
  var out = { ok: true };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // xep hang: 1 lenh ghi tai 1 thoi diem -> khong trung/mat dong
    var p = JSON.parse(e.postData.getDataAsString());
    var ss = SpreadsheetApp.openById(SHEET_ID);

    if (p.loai === 'khach' && p.kh) {
      var sh = ensureSheet_(ss, 'KhachHang', KH_HEAD);
      var k = p.kh;
      var ma = nextMa_(sh, k.kho); // backend tu cap ma duoi khoa -> chong trung tuyet doi
      sh.appendRow([ma, k.ten || '', k.sdt || '', k.qh || '',
                    k.kho || '', k.tt || '', k.kyhan || 0,
                    Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm')]);
      out.added = 1; out.ma = ma;
    } else if (p.loai === 'hoadon' && p.rows && p.rows.length) {
      var sh2 = ensureSheet_(ss, 'HoaDon', HD_HEAD);
      p.rows.forEach(function (r) { sh2.appendRow(r); });
      out.added = p.rows.length;
    }
  } catch (err) {
    out = { ok: false, error: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Cap ma khach tiep theo theo kho (HN0001 / HY0001...). Goi trong khoa doPost.
function nextMa_(sh, kho) {
  var pre = (String(kho).indexOf('Hưng') >= 0) ? 'HY' : 'HN';
  var max = 0, last = sh.getLastRow();
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, 1).getValues();
    var re = new RegExp('^' + pre + '(\\d+)');
    vals.forEach(function (r) {
      var m = String(r[0]).match(re);
      if (m) { var num = parseInt(m[1], 10); if (num > max) max = num; }
    });
  }
  return pre + ('0000' + (max + 1)).slice(-4);
}

// ---------- DOC (GET, tra ve JSONP cho app doc duoc Sheet rieng tu) ----------
function doGet(e) {
  var cb = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : 'callback';
  var out = { ok: true, khach: [], hoadon: [] };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    out.khach = readKhach_(ss);
    out.hoadon = readHoaDon_(ss);
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ---------- Helpers ----------
function ensureSheet_(ss, name, head) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); return sh; }
  if (sh.getLastRow() === 0) { sh.appendRow(head); sh.setFrozenRows(1); return sh; }
  var ec = sh.getLastColumn();
  var eh = sh.getRange(1, 1, 1, ec).getValues()[0].map(function (x) { return String(x).trim(); });
  var isPrefix = eh.length <= head.length;
  for (var i = 0; i < eh.length; i++) { if (eh[i] !== head[i]) { isPrefix = false; break; } }
  if (isPrefix) {
    if (eh.length < head.length) sh.getRange(1, 1, 1, head.length).setValues([head]); // them cot, giu data
    return sh;
  }
  // khac schema hoan toan -> giu du lieu cu sang _old
  sh.setName(name + '_old_' + Date.now());
  var ns = ss.insertSheet(name); ns.appendRow(head); ns.setFrozenRows(1);
  return ns;
}

function rowsOf_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) return { head: [], rows: [] };
  var v = sh.getDataRange().getValues();
  if (v.length < 1) return { head: [], rows: [] };
  var head = v[0].map(function (x) { return String(x).trim(); });
  return { head: head, rows: v.slice(1) };
}

function idx_(head, name) { return head.indexOf(name); }

function num_(v) {
  var s = String(v == null ? '' : v).replace(/[.,\s]/g, '');
  var x = parseFloat(s); return isNaN(x) ? 0 : x;
}

// Ep ngay ve chuoi dd/MM/yyyy (Sheet hay tu chuyen text -> Date)
function fmtD_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'GMT+7', 'dd/MM/yyyy');
  return String(v == null ? '' : v).slice(0, 10);
}

function readKhach_(ss) {
  var d = rowsOf_(ss, 'KhachHang'); if (!d.head.length) return [];
  var h = d.head;
  var iMa = idx_(h, 'Ma KH'), iTen = idx_(h, 'Ten'), iSdt = idx_(h, 'SDT'),
      iQh = idx_(h, 'Quan/Huyen'), iKho = idx_(h, 'Kho'),
      iTt = idx_(h, 'Thanh toan'), iKy = idx_(h, 'Ky han no'), iNg = idx_(h, 'Ngay tao');
  var out = [];
  d.rows.forEach(function (r) {
    if (!String(r[iMa] || '').trim()) return;
    out.push({
      ma: String(r[iMa]).trim(), ten: r[iTen] || '', sdt: String(r[iSdt] || ''),
      qh: r[iQh] || '', kho: r[iKho] || '', tt: r[iTt] || '',
      kyhan: Number(r[iKy]) || 0, ngay: fmtD_(r[iNg])
    });
  });
  return out;
}

function readHoaDon_(ss) {
  var d = rowsOf_(ss, 'HoaDon'); if (!d.head.length) return [];
  var h = d.head;
  var iId = idx_(h, 'ID hoa don'), iNg = idx_(h, 'Ngay'), iMa = idx_(h, 'Ma KH'),
      iKh = idx_(h, 'Khach'), iSd = idx_(h, 'SDT'), iKho = idx_(h, 'Kho'),
      iSp = idx_(h, 'San pham'), iDv = idx_(h, 'DVT'), iSl = idx_(h, 'SL'),
      iGia = idx_(h, 'Don gia'), iTien = idx_(h, 'Thanh tien'), iVat = idx_(h, 'VAT %'),
      iCoc = idx_(h, 'Coc binh'), iTong = idx_(h, 'Tong thanh toan'),
      iDathu = idx_(h, 'Da thu'), iTt = idx_(h, 'Trang thai TT'), iThu = idx_(h, 'Nguoi thu'),
      iHt = idx_(h, 'Hinh thuc TT');
  var out = [];
  d.rows.forEach(function (r) {
    if (!String(r[iId] || '').trim()) return;
    out.push({
      id: String(r[iId]).trim(),
      ngay: fmtD_(r[iNg]),
      ma: String(r[iMa] || '').trim(),
      khach: r[iKh] || '', sdt: String(r[iSd] || ''), kho: r[iKho] || '',
      sp: r[iSp] || '', dvt: r[iDv] || '', sl: num_(r[iSl]), gia: num_(r[iGia]),
      tien: num_(r[iTien]), vat: num_(r[iVat]), coc: num_(r[iCoc]),
      tong: num_(r[iTong]), dathu: num_(r[iDathu]),
      tt: r[iTt] || '', thu: r[iThu] || '', ht: iHt >= 0 ? (r[iHt] || '') : ''
    });
  });
  return out;
}

// Chay tay 1 lan de cap quyen truy cap Sheet (neu deploy bao thieu quyen)
function capQuyen() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log('OK: ' + ss.getName());
}
