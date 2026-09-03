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
var KHO_SHEET_ID = '1tQEhM51DiYHN3tV_zKa6OT1tLXkPBft4TtOgp1_uAck'; // Sheet kho, tab Data co gia le

var KH_HEAD = ['Ma KH', 'Ten', 'SDT', 'Quan/Huyen', 'Kho',
               'Thanh toan', 'Ky han no', 'Ngay tao'];

var HD_HEAD = ['ID hoa don', 'Ngay', 'Ma KH', 'Khach', 'SDT', 'Kho',
               'San pham', 'DVT', 'SL', 'Don gia', 'Thanh tien',
               'VAT %', 'Coc binh', 'Tong thanh toan', 'Da thu',
               'Trang thai TT', 'Nguoi thu', 'Hinh thuc TT', 'Ghi chu',
               'Ngày CK', 'Thang']; // 13/08/2026: cot 20 'Ngày CK' | 29/08/2026: cot 21 'Thang' (yyyy-MM, backend tu dien - khong go tay)

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
      // 16/08/2026: chong ghi trung khi bam "dong bo" nhieu lan -
      // ID hoa don da co tren Sheet thi bo qua ca chum dong cua ID do
      var daCoId = {};
      var lastR = sh2.getLastRow();
      if (lastR >= 2) {
        sh2.getRange(2, 1, lastR - 1, 1).getValues().forEach(function (r0) {
          var id0 = String(r0[0] || '').trim(); if (id0) daCoId[id0] = 1;
        });
      }
      var moi = p.rows.filter(function (r) { return !daCoId[String(r[0] || '').trim()]; });
      // 29/08/2026: ghi them cot 21 'Thang' (yyyy-MM) suy ra tu cot Ngay.
      // Ghi cot A:T bang appendRow, rieng cot U phai ep dinh dang text ('@')
      // TRUOC khi ghi - neu khong Sheet tu doi '2026-08' thanh kieu NGAY,
      // pivot se tach lam 2 dong khac nhau (text vs ngay).
      var dongDau = sh2.getLastRow() + 1;
      moi.forEach(function (r) {
        var row = r.slice(0, 20);
        while (row.length < 20) row.push('');
        sh2.appendRow(row);
      });
      if (moi.length) {
        var cotThang = moi.map(function (r) { return [thang_(r[1])]; });
        sh2.getRange(dongDau, 21, moi.length, 1)
           .setNumberFormat('@')
           .setValues(cotThang);
      }
      out.added = moi.length;
      out.skipped = p.rows.length - moi.length;
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
    out.cat = readCatLe_();
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

// 29/08/2026: doi gia tri cot Ngay -> chuoi thang 'yyyy-MM' (vd 07/08/2026 -> '2026-08').
// Tra ve '' neu khong doc duoc ngay -> de trong, khong bia so.
function thang_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'GMT+7', 'yyyy-MM');
  var s = String(v == null ? '' : v).trim();
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);   // dd/MM/yyyy
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2);
  m = s.match(/^(\d{4})[\/\-](\d{1,2})/);                     // yyyy-MM-dd
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  return '';
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
      iHt = idx_(h, 'Hinh thuc TT'), iGc = idx_(h, 'Ghi chu');
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
      tt: r[iTt] || '', thu: r[iThu] || '', ht: iHt >= 0 ? (r[iHt] || '') : '',
      ghichu: iGc >= 0 ? (r[iGc] || '') : ''
    });
  });
  return out;
}

// Doc danh muc gia le tu tab Data cua Sheet kho (cot Ten hang, DVT, Gia le)
function readCatLe_() {
  var out = [];
  try {
    var sh = SpreadsheetApp.openById(KHO_SHEET_ID).getSheetByName('Data');
    if (!sh || sh.getLastRow() < 2) return out;
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      var r = v[i];
      var ten = String(r[1] || '').trim();
      var gle = num_(r[8]); // cot I "Gia le"
      if (!ten || gle <= 0) continue;
      out.push({ t: ten, dv: String(r[2] || '').trim(), g: gle });
    }
  } catch (e) {}
  return out;
}

// Chay tay 1 lan de cap quyen truy cap Sheet (neu deploy bao thieu quyen)
function capQuyen() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log('OK: ' + ss.getName());
}

// ============================================================
// CHAY TAY 1 LAN (13/08/2026): gop sheet hoa don ve 1 sheet HoaDon.
// Lam gi: dien header 'Ngày CK' vao o T1 cua sheet gop, xoa cot thua
// sau cot T, copy hoa don moi tu sheet HoaDon trang sang (khong trung),
// xoa sheet HoaDon trang, doi ten sheet gop thanh HoaDon.
// CACH CHAY: chon ham gopHoaDon_13082026 tren thanh cong cu -> Run.
// LUU Y: deploy New version TRUOC roi moi chay ham nay.
// Chay xong co the xoa ca khoi code nay di (khong bat buoc).
// ============================================================
function gopHoaDon_13082026() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // chan doPost ghi xen vao giua luc gop
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var gop = ss.getSheetByName('HoaDon_old_1786592187847');
    if (!gop) {
      if (ss.getSheetByName('HoaDon') && !ss.getSheetByName('HoaDon_old_1786592187847')) {
        Logger.log('Khong thay sheet gop — co the da chay xong truoc do. Khong lam gi.');
        return;
      }
      throw new Error('Khong tim thay sheet HoaDon_old_1786592187847');
    }

    // 1) Header o T1 phai dung y het HD_HEAD cot 20
    gop.getRange(1, 20).setValue(HD_HEAD[19]); // 'Ngày CK'

    // 2) Xoa han cac cot sau cot T (U tro di) cho sach
    var mc = gop.getMaxColumns();
    if (mc > 20) gop.deleteColumns(21, mc - 20);

    // 3) Copy hoa don tu sheet HoaDon trang sang cuoi sheet gop.
    //    Dedupe theo ID HOA DON (1 hoa don nhieu san pham = nhieu dong cung ID
    //    -> phai copy du ca chum dong, chi bo qua ID da co san trong sheet gop).
    var cur = ss.getSheetByName('HoaDon');
    var daCopy = 0;
    if (cur && cur.getLastRow() >= 2) {
      var daCo = {};
      var last = gop.getLastRow();
      if (last >= 2) {
        gop.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
          var id = String(r[0] || '').trim(); if (id) daCo[id] = 1;
        });
      }
      var rows = cur.getRange(2, 1, cur.getLastRow() - 1, 19).getValues();
      var idMoi = {};
      rows.forEach(function (r) {
        var id = String(r[0] || '').trim();
        if (id && !daCo[id]) idMoi[id] = 1;
      });
      rows.forEach(function (r) {
        var id = String(r[0] || '').trim();
        if (idMoi[id]) { gop.appendRow(r); daCopy++; }
      });
    }

    // 4) Xoa sheet HoaDon trang, doi ten sheet gop thanh HoaDon
    if (cur) ss.deleteSheet(cur);
    gop.setName('HoaDon');

    Logger.log('XONG: copy them ' + daCopy + ' dong. Sheet HoaDon hien co ' +
               (gop.getLastRow() - 1) + ' dong du lieu.');
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ============================================================
// CHAY TAY 1 LAN (29/08/2026): them cot 21 'Thang' vao sheet HoaDon.
// Lam gi: ghi header 'Thang' vao o U1, tinh 'yyyy-MM' tu cot Ngay (cot B)
// va dien cho toan bo dong da co. Ghi GIA TRI TINH SAN dang text (khong
// dung cong thuc) de khong lam lech getLastRow() cua doPost khi ghi hoa don.
// CACH CHAY: deploy New version TRUOC -> chon ham themCotThang_29082026 -> Run.
// Chay lai nhieu lan van an toan (chi ghi de dung gia tri do).
// 29/08/2026 (lan 2): chay lai ham nay de chuan hoa cac dong da bi
// Sheet doi thanh kieu NGAY ve lai text -> pivot gop dung 1 dong/thang.
// ============================================================
function themCotThang_29082026() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // chan doPost ghi xen vao giua
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('HoaDon');
    if (!sh) throw new Error('Khong tim thay sheet HoaDon');

    var mc = sh.getMaxColumns();
    if (mc < 21) sh.insertColumnsAfter(mc, 21 - mc);

    // Ep CA cot U ve dinh dang text truoc, ke ca cac dong con trong ->
    // dong hoa don moi them sau nay cung thua dinh dang text, khong bi
    // Sheet doi '2026-08' thanh kieu ngay.
    sh.getRange(1, 21, sh.getMaxRows(), 1).setNumberFormat('@');
    sh.getRange(1, 21).setValue(HD_HEAD[20]); // 'Thang'

    var n = sh.getLastRow() - 1;
    if (n < 1) { Logger.log('Sheet chua co dong du lieu nao. Chi them header.'); return; }

    var ngay = sh.getRange(2, 2, n, 1).getValues(); // cot B 'Ngay'
    var out = ngay.map(function (r) { return [thang_(r[0])]; });
    sh.getRange(2, 21, n, 1).setNumberFormat('@').setValues(out); // '@' = text, tranh bi ep thanh ngay

    var trong = out.filter(function (r) { return !r[0]; }).length;
    Logger.log('XONG: dien Thang cho ' + n + ' dong. Khong doc duoc ngay: ' + trong + ' dong.');
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
