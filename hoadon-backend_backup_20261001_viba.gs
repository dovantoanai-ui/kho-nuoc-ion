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

// 01/10/2026: cot 'Link bill' (link anh bill chuyen tien tren Drive) KHONG gan cung vi tri:
// backend tim theo TEN cot, chua co thi them vao sau cot cuoi -> khong dung cot anh tu them.
var COT_LINK = 'Link bill';

// 01/10/2026: tab HoaDonHuy = cac dong hoa don bi huy (nhap sai / VIBA hoan hang),
// cung cau truc HoaDon + 5 cot luu vet. HoaDon chi giu hoa don dang hieu luc ->
// kho HY/HN, doanh thu, dashboard tu dung ma khong phai sua cho khac.
var HUY_HEAD = HD_HEAD.concat([COT_LINK, 'Ly do huy', 'Nguoi huy', 'Ngay huy', 'Thay bang']);

// 01/10/2026: so doi soat tien CK. Moi dong = 1 hoa don trong 1 dot nop tien.
var DS_HEAD = ['ID dot', 'Ngay nop', 'Nguoi nop', 'TK nhan', 'Tong nop',
               'ID hoa don', 'Khach', 'So tien', 'Ngay khach TT',
               'Trang thai truoc', 'Trang thai sau', 'Ghi chu', 'Link anh', 'Ghi luc', 'Thang'];
var DS_FOLDER = 'Doi soat CK ION FUJI'; // folder Drive luu anh bill doi soat

// 18/08/2026: so phieu thu cong no. Moi lan thu tien = 1 dong o day, backend
// dong thoi cap nhat lai 'Da thu' + 'Trang thai TT' cua hoa don goc ben tab HoaDon.
// Doanh thu VAN doc o tab HoaDon -> phieu thu khong lam doanh thu bi dem doi.
var TT_HEAD = ['ID phieu', 'Ngay', 'Ma KH', 'Khach', 'Kho', 'ID hoa don',
               'So tien', 'Hinh thuc', 'Nguoi thu', 'Ghi chu', 'Link anh',
               'Ghi luc', 'Thang'];

var ANH_FOLDER = 'Anh thanh toan ION FUJI'; // backend tu tao trong Drive neu chua co

// 23/09/2026: don giao qua doi tac VIBA. Moi hoa don tich "Giao qua VIBA" = 1 dong.
// 23/09/2026 (ban 2): don VIBA CHUA ghi tab HoaDon khi lap. Cac dong hoa don nam tam
// o tab 'HoaDonCho'; bam "Da giao" moi chuyen sang HoaDon (ngay = ngay giao) ->
// doanh thu, tru kho HY/HN, cong no chi phat sinh khi khach da nhan hang.
// Bam "Giao khong thanh cong" -> cac dong chuyen sang 'HoaDonHuy' (luu vet), HoaDon khong bi dong.
var VB_HEAD = ['ID hoa don', 'Ngay', 'Kho', 'Ma KH', 'Khach',
               'Nguoi nhan', 'SDT nhan', 'Dia chi giao', 'Hang hoa',
               'Tong tien', 'Thu ho', 'Ghi chu giao',
               'Trang thai', 'Ngay giao', 'Nguoi xac nhan', 'Cap nhat luc', 'Thang',
               'Ly do hoan', 'Ngay hoan'];
var VB_COT_TEXT = [2, 7, 14, 17, 19]; // Ngay, SDT nhan, Ngay giao, Thang, Ngay hoan: ep text de khong mat so 0 dau / bi doi kieu ngay

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
      var kq = ghiHoaDonRows_(ss, 'HoaDon', p.rows);
      out.added = kq.added;
      out.skipped = kq.skipped;
      // Tuong thich app ban cu (truoc 23/09 ban 2) van gui viba kem loai 'hoadon'
      if (p.viba && p.viba.row) out.viba = ghiViba_(ss, p.viba.row);
    } else if (p.loai === 'vibamoi' && p.rows && p.rows.length && p.viba && p.viba.row) {
      // Don VIBA moi: ghi GiaoVIBA + giu tam hoa don o HoaDonCho, CHUA vao HoaDon
      out.viba = ghiViba_(ss, p.viba.row);
      if (out.viba.added) out.cho = ghiHoaDonRows_(ss, 'HoaDonCho', p.rows);
    } else if (p.loai === 'thanhtoan' && p.rows && p.rows.length) {
      out = ghiThanhToan_(ss, p);
    } else if (p.loai === 'vibagiao' && p.id) {
      out = xacNhanViba_(ss, p);
    } else if (p.loai === 'vibahuy' && p.id) {
      out = huyViba_(ss, p);
    } else if (p.loai === 'huyhd' && p.id) {
      out = huyHoaDon_(ss, p);
    } else if (p.loai === 'doisoatck' && p.dot && p.rows && p.rows.length) {
      out = doiSoatCK_(ss, p);
    }
  } catch (err) {
    out = { ok: false, error: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Ghi cac dong hoa don vao tab name (HoaDon / HoaDonCho / HoaDonHuy) - cung 1 cau truc HD_HEAD.
// 16/08/2026: ID hoa don da co trong tab thi bo qua ca chum dong cua ID do (chong ghi trung).
// 29/08/2026: cot 21 'Thang' (yyyy-MM) ep dinh dang text TRUOC khi ghi, neu khong Sheet
// tu doi '2026-08' thanh kieu NGAY va pivot tach lam 2 dong (text vs ngay).
// 01/10/2026: tab HoaDonHuy dung HUY_HEAD (dai hon). Phan tu tu vi tri 21 tro di cua
// moi dong (Link bill, Ly do huy...) ghi tiep vao cot 22+. Cot 21 'Thang' luon tinh lai.
function ghiHoaDonRows_(ss, name, rows) {
  var head = (name === 'HoaDonHuy') ? HUY_HEAD : HD_HEAD;
  var sh = ensureSheet_(ss, name, head);
  var daCoId = {};
  var lastR = sh.getLastRow();
  if (lastR >= 2) {
    sh.getRange(2, 1, lastR - 1, 1).getValues().forEach(function (r0) {
      var id0 = String(r0[0] || '').trim(); if (id0) daCoId[id0] = 1;
    });
  }
  var moi = rows.filter(function (r) { return !daCoId[String(r[0] || '').trim()]; });
  var dongDau = sh.getLastRow() + 1;
  moi.forEach(function (r) {
    var row = r.slice(0, 20);
    while (row.length < 20) row.push('');
    sh.appendRow(row);
  });
  if (moi.length) {
    var cotThang = moi.map(function (r) { return [thang_(r[1])]; });
    sh.getRange(dongDau, 21, moi.length, 1).setNumberFormat('@').setValues(cotThang);
    var soThem = head.length - 21;
    if (soThem > 0) {
      var them = moi.map(function (r) {
        var x = r.slice(21, 21 + soThem);
        while (x.length < soThem) x.push('');
        return x;
      });
      var coChu = them.some(function (x) { return x.some(function (v) { return v !== '' && v != null; }); });
      if (coChu) sh.getRange(dongDau, 22, moi.length, soThem).setNumberFormat('@').setValues(them);
    }
  }
  return { added: moi.length, skipped: rows.length - moi.length };
}

// Doc (CHUA xoa) moi dong cua 1 ID hoa don trong tab name.
// Tra ve { sh, rows (21 cot dau, kem them gia tri Link bill o vi tri 21), dong }.
// Goi xoaDong_ SAU khi da ghi xong noi moi: buoc ghi loi giua chung thi du lieu cu con nguyen.
function docHoaDonId_(ss, name, id) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return { sh: sh, rows: [], dong: [] };
  var n = sh.getLastRow() - 1;
  var w = Math.max(sh.getLastColumn(), HD_HEAD.length);
  if (sh.getMaxColumns() < w) w = sh.getMaxColumns();
  var head = sh.getRange(1, 1, 1, w).getValues()[0].map(function (x) { return String(x).trim(); });
  var iLink = head.indexOf(COT_LINK);
  var vals = sh.getRange(2, 1, n, w).getValues();
  var rows = [], dong = [];
  for (var i = 0; i < n; i++) {
    if (String(vals[i][0] || '').trim() !== id) continue;
    var r = vals[i].slice(0, HD_HEAD.length);
    while (r.length < HD_HEAD.length) r.push('');
    r.push(iLink >= 0 ? vals[i][iLink] : '');
    rows.push(r); dong.push(i + 2);
  }
  return { sh: sh, rows: rows, dong: dong };
}

// So cot (1-based) cua cot ten name trong sh; chua co thi tao o sau cot cuoi.
function cotTheoTen_(sh, name) {
  var lc = Math.max(sh.getLastColumn(), 1);
  var head = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return String(x).trim(); });
  var i = head.indexOf(name);
  if (i >= 0) return i + 1;
  var c = lc + 1;
  if (sh.getMaxColumns() < c) sh.insertColumnsAfter(sh.getMaxColumns(), c - sh.getMaxColumns());
  sh.getRange(1, c).setValue(name);
  return c;
}
function xoaDong_(sh, dong) {
  for (var j = dong.length - 1; j >= 0; j--) sh.deleteRow(dong[j]); // xoa tu duoi len de khong lech so dong
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
  var out = { ok: true, khach: [], hoadon: [], thanhtoan: [], viba: [], cho: [], huy: [] };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    out.khach = readKhach_(ss);
    out.hoadon = readHoaDon_(ss);
    out.thanhtoan = readThanhToan_(ss);
    out.viba = readViba_(ss);
    out.cho = readHoaDon_(ss, 'HoaDonCho'); // hang dang di giao VIBA (chua vao HoaDon)
    out.huy = readHoaDon_(ss, 'HoaDonHuy'); // 01/10/2026: hoa don da huy -> app gan nhan 'Da huy'
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
  // 01/10/2026: so khop tren phan chung. Tab co THEM cot o cuoi (anh tu them) van hop le.
  var n = Math.min(eh.length, head.length), khop = true;
  for (var i = 0; i < n; i++) { if (eh[i] !== head[i]) { khop = false; break; } }
  if (khop) {
    if (eh.length < head.length) {
      if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
      sh.getRange(1, 1, 1, head.length).setValues([head]); // them cot, giu data
    }
    return sh;
  }
  // 01/10/2026: tab DA CO du lieu ma khac cau truc -> DUNG LAI, khong doi ten/ghi de
  // (truoc day doi ten sang _old roi tao tab trang -> de mat dau du lieu dang dung).
  if (sh.getLastRow() > 1) {
    throw new Error('Tab ' + name + ' co tieu de cot ' + (i + 1) + ' la "' + eh[i] +
                    '", backend can "' + head[i] + '". Khong ghi de - sua lai tieu de cot roi thu lai.');
  }
  sh.getRange(1, 1, 1, head.length).setValues([head]); // tab chi co dong tieu de -> dat lai tieu de
  return sh;
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

// name mac dinh 'HoaDon'; 23/09/2026 dung lai de doc 'HoaDonCho' (cung cau truc)
function readHoaDon_(ss, name) {
  var d = rowsOf_(ss, name || 'HoaDon'); if (!d.head.length) return [];
  var h = d.head;
  var iId = idx_(h, 'ID hoa don'), iNg = idx_(h, 'Ngay'), iMa = idx_(h, 'Ma KH'),
      iKh = idx_(h, 'Khach'), iSd = idx_(h, 'SDT'), iKho = idx_(h, 'Kho'),
      iSp = idx_(h, 'San pham'), iDv = idx_(h, 'DVT'), iSl = idx_(h, 'SL'),
      iGia = idx_(h, 'Don gia'), iTien = idx_(h, 'Thanh tien'), iVat = idx_(h, 'VAT %'),
      iCoc = idx_(h, 'Coc binh'), iTong = idx_(h, 'Tong thanh toan'),
      iDathu = idx_(h, 'Da thu'), iTt = idx_(h, 'Trang thai TT'), iThu = idx_(h, 'Nguoi thu'),
      iHt = idx_(h, 'Hinh thuc TT'), iGc = idx_(h, 'Ghi chu');
  // 01/10/2026: them Ngay CK, Link bill, va 4 cot luu vet khi doc tab HoaDonHuy
  var iCk = idx_(h, 'Ngày CK'), iLink = idx_(h, COT_LINK), iLd = idx_(h, 'Ly do huy'),
      iNh = idx_(h, 'Nguoi huy'), iNgh = idx_(h, 'Ngay huy'), iTb = idx_(h, 'Thay bang');
  function o(i, r) { return i >= 0 ? String(r[i] == null ? '' : r[i]) : ''; }
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
      ghichu: iGc >= 0 ? (r[iGc] || '') : '',
      ngayck: iCk >= 0 ? fmtD_(r[iCk]) : '', link: o(iLink, r),
      lydo: o(iLd, r), nguoihuy: o(iNh, r), ngayhuy: iNgh >= 0 ? fmtD_(r[iNgh]) : '', thaythe: o(iTb, r)
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

// ============================================================
// 18/08/2026 — THU CONG NO
// Ghi phieu thu vao tab 'ThanhToan', up anh chung tu len Drive,
// roi cap nhat lai 'Da thu' + 'Trang thai TT' cua hoa don goc.
// ============================================================
function ghiThanhToan_(ss, p) {
  var out = { ok: true, added: 0, skipped: 0, capnhat: 0, anh: '' };
  var sh = ensureSheet_(ss, 'ThanhToan', TT_HEAD);

  // Chong ghi trung khi bam Luu nhieu lan: ID phieu da co tren Sheet thi bo qua
  var daCo = {};
  var last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      var id = String(r[0] || '').trim(); if (id) daCo[id] = 1;
    });
  }
  var moi = p.rows.filter(function (r) { return !daCo[String(r[0] || '').trim()]; });
  out.skipped = p.rows.length - moi.length;
  if (!moi.length) return out;

  // 1 lan chuyen khoan co the tra nhieu hoa don -> up anh 1 lan, dung chung link
  var link = '';
  if (p.anh && p.anh.data) link = uploadAnh_(p.anh.data, p.anh.mime, p.anh.ten);

  var ghiLuc = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm');
  var dongDau = sh.getLastRow() + 1;
  moi.forEach(function (r) {
    var row = r.slice(0, 10);
    while (row.length < 10) row.push('');
    row.push(link);   // cot 11 Link anh
    row.push(ghiLuc); // cot 12 Ghi luc
    sh.appendRow(row);
  });
  // Cot 13 'Thang' phai ep dinh dang text TRUOC khi ghi, neu khong Sheet
  // tu doi '2026-08' thanh kieu NGAY (giong xu ly ben tab HoaDon).
  var cotThang = moi.map(function (r) { return [thang_(r[1])]; });
  sh.getRange(dongDau, 13, moi.length, 1).setNumberFormat('@').setValues(cotThang);

  // Cong so tien vua thu vao dung hoa don ben tab HoaDon
  var themTheoHD = {};
  moi.forEach(function (r) {
    var idhd = String(r[5] || '').trim(); if (!idhd) return;
    themTheoHD[idhd] = (themTheoHD[idhd] || 0) + num_(r[6]);
  });
  out.capnhat = capNhatDaThu_(ss, themTheoHD);
  out.added = moi.length;
  out.anh = link;
  return out;
}

// Cong don tien thu vao dong DAU TIEN cua moi hoa don (dong mang tong tien
// va trang thai; cac dong sau chi la san pham). Chi cong phan VUA ghi ->
// gui lai cung ID phieu se bi bo qua o tren nen khong cong hai lan.
function capNhatDaThu_(ss, themTheoHD) {
  var sh = ss.getSheetByName('HoaDon');
  if (!sh || sh.getLastRow() < 2) return 0;
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
               .map(function (x) { return String(x).trim(); });
  var cId = idx_(head, 'ID hoa don') + 1;
  var cTong = idx_(head, 'Tong thanh toan') + 1;
  var cThu = idx_(head, 'Da thu') + 1;
  var cTt = idx_(head, 'Trang thai TT') + 1;
  if (cId < 1 || cTong < 1 || cThu < 1 || cTt < 1) return 0;

  var n = sh.getLastRow() - 1;
  var ids = sh.getRange(2, cId, n, 1).getValues();
  var dem = 0;
  for (var id in themTheoHD) {
    var dong = -1;
    for (var i = 0; i < n; i++) {
      if (String(ids[i][0] || '').trim() === id) { dong = i + 2; break; }
    }
    if (dong < 0) continue;
    var tong = num_(sh.getRange(dong, cTong).getValue());
    var thuMoi = num_(sh.getRange(dong, cThu).getValue()) + themTheoHD[id];
    if (tong > 0 && thuMoi > tong) thuMoi = tong; // khong ghi vuot tong hoa don
    sh.getRange(dong, cThu).setValue(thuMoi);
    sh.getRange(dong, cTt).setValue(thuMoi >= tong ? 'Đã thanh toán' : 'Thanh toán 1 phần');
    dem++;
  }
  return dem;
}

// Luu anh chung tu vao Drive cua tai khoan chay script. Loi anh KHONG lam
// hong phieu thu - tra ve chuoi bao loi de con nhin thay tren Sheet.
function uploadAnh_(b64, mime, ten, tenFolder) {
  try {
    var folder = layFolderAnh_(tenFolder);
    var blob = Utilities.newBlob(Utilities.base64Decode(b64),
                                 mime || 'image/jpeg',
                                 ten || ('thu-' + Date.now() + '.jpg'));
    var f = folder.createFile(blob);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) {}
    return 'https://drive.google.com/file/d/' + f.getId() + '/view';
  } catch (err) {
    return 'LOI ANH: ' + err.message;
  }
}

function layFolderAnh_(tenFolder) {
  var ten = tenFolder || ANH_FOLDER;
  var it = DriveApp.getFoldersByName(ten);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ten);
}

function readThanhToan_(ss) {
  var d = rowsOf_(ss, 'ThanhToan'); if (!d.head.length) return [];
  var h = d.head;
  var iId = idx_(h, 'ID phieu'), iNg = idx_(h, 'Ngay'), iMa = idx_(h, 'Ma KH'),
      iKh = idx_(h, 'Khach'), iKho = idx_(h, 'Kho'), iHd = idx_(h, 'ID hoa don'),
      iTien = idx_(h, 'So tien'), iHt = idx_(h, 'Hinh thuc'), iThu = idx_(h, 'Nguoi thu'),
      iGc = idx_(h, 'Ghi chu'), iAnh = idx_(h, 'Link anh');
  var out = [];
  d.rows.forEach(function (r) {
    if (!String(r[iId] || '').trim()) return;
    out.push({
      id: String(r[iId]).trim(), ngay: fmtD_(r[iNg]),
      ma: String(r[iMa] || '').trim(), khach: r[iKh] || '', kho: r[iKho] || '',
      idhd: String(r[iHd] || '').trim(), tien: num_(r[iTien]),
      ht: r[iHt] || '', thu: r[iThu] || '',
      ghichu: iGc >= 0 ? (r[iGc] || '') : '',
      anh: iAnh >= 0 ? String(r[iAnh] || '') : ''
    });
  });
  return out;
}

// ============================================================
// 23/09/2026 — GIAO HANG QUA VIBA
// ============================================================
// Ghi 1 dong don giao VIBA. row tu app: [ID hoa don, Ngay, Kho, Ma KH, Khach,
// Nguoi nhan, SDT nhan, Dia chi giao, Hang hoa, Tong tien, Thu ho, Ghi chu giao]
function ghiViba_(ss, r) {
  var sh = ensureSheet_(ss, 'GiaoVIBA', VB_HEAD);
  var id = String(r[0] || '').trim();
  if (!id) return { added: 0 };
  if (timDongViba_(sh, id) > 0) return { added: 0, skipped: 1 }; // bam dong bo nhieu lan -> khong ghi trung

  var row = r.slice(0, 12);
  while (row.length < 12) row.push('');
  row.push('Chờ giao');                                              // 13 Trang thai
  row.push('');                                                      // 14 Ngay giao
  row.push('');                                                      // 15 Nguoi xac nhan
  row.push(Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm')); // 16 Cap nhat luc
  row.push(thang_(r[1]));                                            // 17 Thang
  row.push('');                                                      // 18 Ly do hoan
  row.push('');                                                      // 19 Ngay hoan

  // Ep text cac cot de mat so 0 dau SDT / bi doi kieu ngay TRUOC khi ghi
  var dong = sh.getLastRow() + 1;
  VB_COT_TEXT.forEach(function (c) { sh.getRange(dong, c).setNumberFormat('@'); });
  sh.getRange(dong, 1, 1, VB_HEAD.length).setValues([row.map(function (v, i) {
    return VB_COT_TEXT.indexOf(i + 1) >= 0 ? String(v == null ? '' : v) : v;
  })]);
  return { added: 1 };
}

// ION bam "Xac nhan da giao" -> chuyen hoa don tu HoaDonCho sang HoaDon (ngay = ngay giao),
// roi doi trang thai GiaoVIBA. Don kieu cu (da nam san trong HoaDon) thi chi doi trang thai.
function xacNhanViba_(ss, p) {
  var id = String(p.id).trim();
  var sh = ss.getSheetByName('GiaoVIBA');
  if (!sh) return { ok: false, error: 'Chua co tab GiaoVIBA' };
  var dong = timDongViba_(sh, id);
  if (dong < 0) return { ok: false, error: 'Khong tim thay don ' + id };
  var ttHien = String(sh.getRange(dong, 13).getValue() || '');
  if (ttHien === 'Đã giao') return { ok: true, skipped: 1 };                 // bam 2 lan -> khong chuyen 2 lan
  if (ttHien === 'Hoàn hàng') return { ok: false, error: 'Don ' + id + ' da hoan hang' };

  var ngayGiao = String(p.ngaygiao || Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy'));
  var cho = docHoaDonId_(ss, 'HoaDonCho', id);
  var chuyen = 0;
  if (cho.rows.length) {
    var rows = cho.rows.map(function (r) { var x = r.slice(); x[1] = ngayGiao; return x; });
    chuyen = ghiHoaDonRows_(ss, 'HoaDon', rows).added;
    xoaDong_(cho.sh, cho.dong); // ghi HoaDon xong moi xoa ben HoaDonCho
  }
  sh.getRange(dong, 14).setNumberFormat('@');
  sh.getRange(dong, 13, 1, 4).setValues([[
    'Đã giao', ngayGiao, String(p.nguoi || ''),
    Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm')
  ]]);
  return { ok: true, capnhat: 1, chuyenHoaDon: chuyen };
}

// ION bam "Giao khong thanh cong" -> dong hoa don tu HoaDonCho sang HoaDonHuy (luu vet),
// GiaoVIBA doi 'Hoàn hàng' + ly do. HoaDon / kho / cong no KHONG bi dong.
function huyViba_(ss, p) {
  var id = String(p.id).trim();
  var sh = ss.getSheetByName('GiaoVIBA');
  if (!sh) return { ok: false, error: 'Chua co tab GiaoVIBA' };
  var dong = timDongViba_(sh, id);
  if (dong < 0) return { ok: false, error: 'Khong tim thay don ' + id };
  var ttHien = String(sh.getRange(dong, 13).getValue() || '');
  if (ttHien === 'Hoàn hàng') return { ok: true, skipped: 1 };
  if (ttHien === 'Đã giao') return { ok: false, error: 'Don ' + id + ' da giao, khong hoan duoc' };

  var cho = docHoaDonId_(ss, 'HoaDonCho', id);
  // Don kieu cu da nam trong HoaDon: KHONG tu xoa o HoaDon, bao de xu ly tay
  if (!cho.rows.length) return { ok: false, error: 'Don ' + id + ' khong co trong HoaDonCho (don kieu cu)' };
  var ngay = String(p.ngay || Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy'));
  var kqHuy = ghiHoaDonRows_(ss, 'HoaDonHuy', cho.rows.map(function (r) {
    return r.concat(['VIBA hoàn hàng: ' + String(p.lydo || ''), String(p.nguoi || ''), ngay, '']);
  }));
  if (!kqHuy.added) return { ok: false, error: 'Ma ' + id + ' da co trong HoaDonHuy - khong xoa, kiem tra tay' };
  xoaDong_(cho.sh, cho.dong);
  var rows = cho.rows;
  sh.getRange(dong, 19).setNumberFormat('@');
  sh.getRange(dong, 13).setValue('Hoàn hàng');
  sh.getRange(dong, 15, 1, 2).setValues([[String(p.nguoi || ''),
    Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm')]]);
  sh.getRange(dong, 18, 1, 2).setValues([[String(p.lydo || ''), ngay]]);
  return { ok: true, capnhat: 1, huy: rows.length };
}

// Tra ve so dong (>=2) cua ID hoa don trong tab GiaoVIBA, -1 neu chua co
function timDongViba_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === id) return i + 2;
  }
  return -1;
}

function readViba_(ss) {
  var d = rowsOf_(ss, 'GiaoVIBA'); if (!d.head.length) return [];
  var h = d.head;
  var c = {};
  VB_HEAD.forEach(function (name) { c[name] = idx_(h, name); });
  function g(r, name) { return c[name] >= 0 ? r[c[name]] : ''; }
  var out = [];
  d.rows.forEach(function (r) {
    var id = String(g(r, 'ID hoa don') || '').trim(); if (!id) return;
    out.push({
      id: id, ngay: fmtD_(g(r, 'Ngay')), kho: g(r, 'Kho') || '',
      ma: String(g(r, 'Ma KH') || '').trim(), khach: g(r, 'Khach') || '',
      nhan: g(r, 'Nguoi nhan') || '', sdt: String(g(r, 'SDT nhan') || ''),
      dc: g(r, 'Dia chi giao') || '', hang: g(r, 'Hang hoa') || '',
      tong: num_(g(r, 'Tong tien')), thuho: num_(g(r, 'Thu ho')),
      gc: g(r, 'Ghi chu giao') || '', tt: g(r, 'Trang thai') || 'Chờ giao',
      ngaygiao: fmtD_(g(r, 'Ngay giao')), nguoi: g(r, 'Nguoi xac nhan') || '',
      lydo: g(r, 'Ly do hoan') || '', ngayhoan: fmtD_(g(r, 'Ngay hoan'))
    });
  });
  return out;
}

// ============================================================
// 01/10/2026 — HUY HOA DON (nhap sai) + LAP PHIEU THAY THE
// p = { id, lydo, nguoi, ngay, thaythe: [dong hoa don moi] | null }
// Thu tu an toan: ghi phieu moi -> chep phieu cu sang HoaDonHuy -> CUOI CUNG moi xoa
// phieu cu o HoaDon. Loi giua chung thi phieu cu van con, bam lai khong bi trung.
// ============================================================
function huyHoaDon_(ss, p) {
  var id = String(p.id).trim();
  var lydo = String(p.lydo || '').trim(), nguoi = String(p.nguoi || '').trim();
  if (!lydo || !nguoi) return { ok: false, error: 'Can ghi ly do va nguoi huy' };
  var cu = docHoaDonId_(ss, 'HoaDon', id);
  if (!cu.rows.length) {
    if (docHoaDonId_(ss, 'HoaDonHuy', id).rows.length) return { ok: true, skipped: 1 }; // bam 2 lan
    return { ok: false, error: 'Khong tim thay hoa don ' + id + ' trong HoaDon' };
  }
  var ngay = String(p.ngay || Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy'));
  var idMoi = '';
  if (p.thaythe && p.thaythe.length) {
    idMoi = String(p.thaythe[0][0] || '').trim();
    if (!idMoi || idMoi === id) return { ok: false, error: 'Ma phieu thay the khong hop le' };
    ghiHoaDonRows_(ss, 'HoaDon', p.thaythe);
    var shHD = ss.getSheetByName('HoaDon');
    var dongMoi = timDongDau_(shHD, idMoi);
    var dau = cu.rows[0];
    // Mang sang phieu moi: Ngay CK + Link bill da doi soat cua phieu cu (neu co)
    if (dongMoi > 0 && dau[19]) shHD.getRange(dongMoi, 20).setNumberFormat('@').setValue(fmtD_(dau[19]));
    if (dongMoi > 0 && dau[21]) shHD.getRange(dongMoi, cotTheoTen_(shHD, COT_LINK)).setValue(dau[21]);
    doiIdHoaDon_(ss, 'ThanhToan', 'ID hoa don', id, idMoi); // phieu thu cu tro sang phieu moi
    doiIdHoaDon_(ss, 'GiaoVIBA', 'ID hoa don', id, idMoi);
  }
  var kqHuy = ghiHoaDonRows_(ss, 'HoaDonHuy', cu.rows.map(function (r) {
    return r.concat([lydo, nguoi, ngay, idMoi]);
  }));
  // Ma nay da co san trong HoaDonHuy (hiem) -> KHONG xoa o HoaDon, tranh mat dong khong luu vet
  if (!kqHuy.added) return { ok: false, error: 'Ma ' + id + ' da co trong HoaDonHuy - khong xoa, kiem tra tay' };
  xoaDong_(cu.sh, cu.dong); // phieu moi chi them o cuoi tab -> so dong phieu cu khong doi
  return { ok: true, huy: cu.rows.length, thaythe: idMoi };
}

// So dong (>=2) dau tien cua ID hoa don trong sh, -1 neu khong co
function timDongDau_(sh, id) {
  if (!sh || sh.getLastRow() < 2) return -1;
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0] || '').trim() === id) return i + 2;
  return -1;
}

// Doi ma hoa don cu -> moi o cot tenCot cua tab name (neu tab/cot ton tai)
function doiIdHoaDon_(ss, name, tenCot, idCu, idMoi) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return 0;
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
  var c = head.indexOf(tenCot) + 1;
  if (c < 1) return 0;
  var vals = sh.getRange(2, c, sh.getLastRow() - 1, 1).getValues();
  var dem = 0;
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === idCu) { sh.getRange(i + 2, c).setValue(idMoi); dem++; }
  }
  return dem;
}

// ============================================================
// 01/10/2026 — DOI SOAT TIEN CK (nhan vien thu ho, nop gop ve TK cong ty)
// p = { dot: {id, ngaynop, nguoinop, tknhan, tongnop, ghichu},
//       rows: [{id: ID hoa don, tien, ngaykh: ngay khach TT, ghichu}],
//       anh: [{data: base64, mime, ten}], ht: 'Tiền mặt', thu: true (chi tinh thu, khong ghi) }
// Sua THANG dong dau cua hoa don o tab HoaDon: Da thu / Trang thai TT (neu chua thanh toan),
// Hinh thuc TT, Ngày CK = ngay nop, Ghi chu (them, khong xoa cu), Link bill.
// Ghi them tab DoiSoatCK de tra cuu theo dot. Thieu 1 hoa don -> dung, KHONG ghi gi.
// ============================================================
function doiSoatCK_(ss, p) {
  var d = p.dot, idDot = String(d.id || '').trim();
  if (!idDot || !d.ngaynop) return { ok: false, error: 'Thieu ma dot / ngay nop' };
  var ghi = !p.thu;
  var shDs = ss.getSheetByName('DoiSoatCK');
  if (shDs && timDongDau_(shDs, idDot) > 0) return { ok: true, skipped: 1, error: 'Dot ' + idDot + ' da ghi truoc do' };

  var shHD = ss.getSheetByName('HoaDon');
  if (!shHD) return { ok: false, error: 'Khong co tab HoaDon' };
  var head = shHD.getRange(1, 1, 1, shHD.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
  function cot(t) { return head.indexOf(t) + 1; }
  var cTong = cot('Tong thanh toan'), cThu = cot('Da thu'), cTt = cot('Trang thai TT'),
      cHt = cot('Hinh thuc TT'), cGc = cot('Ghi chu'), cCk = cot('Ngày CK'), cKh = cot('Khach');
  if (cTong < 1 || cThu < 1 || cTt < 1 || cHt < 1 || cGc < 1 || cCk < 1) return { ok: false, error: 'HoaDon thieu cot can thiet' };

  var ids = shHD.getRange(2, 1, Math.max(shHD.getLastRow() - 1, 1), 1).getValues();
  var dongCua = {};
  for (var i = ids.length - 1; i >= 0; i--) { var k = String(ids[i][0] || '').trim(); if (k) dongCua[k] = i + 2; } // giu dong DAU
  var thieu = p.rows.filter(function (x) { return !dongCua[String(x.id).trim()]; }).map(function (x) { return x.id; });
  if (thieu.length) return { ok: false, error: 'Khong tim thay hoa don: ' + thieu.join(', '), thieu: thieu };

  var links = [];
  if (ghi && p.anh) links = p.anh.map(function (a) { return uploadAnh_(a.data, a.mime, a.ten, DS_FOLDER); });
  var linkStr = links.join(' | ');
  var ht = String(p.ht || 'Tiền mặt');
  var cLink = ghi ? cotTheoTen_(shHD, COT_LINK) : 0;
  var ghiLuc = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm');
  var ketqua = [], dsRows = [];

  p.rows.forEach(function (x) {
    var id = String(x.id).trim(), dong = dongCua[id], tien = num_(x.tien);
    var tong = num_(shHD.getRange(dong, cTong).getValue());
    var thuCu = num_(shHD.getRange(dong, cThu).getValue());
    var ttCu = String(shHD.getRange(dong, cTt).getValue() || '');
    var thuMoi = thuCu, ttMoi = ttCu;
    if (ttCu !== 'Đã thanh toán') {           // da thanh toan tu luc ban -> khong cong tien lan 2
      thuMoi = thuCu + tien; if (tong > 0 && thuMoi > tong) thuMoi = tong;
      ttMoi = thuMoi >= tong ? 'Đã thanh toán' : 'Thanh toán 1 phần';
    }
    var note = 'ĐS CK ' + idDot + ' nộp ' + d.ngaynop + ' (KH TT ' + (x.ngaykh || '?') +
               (x.ghichu ? ', ' + x.ghichu : '') + ')';
    ketqua.push({ id: id, khach: cKh > 0 ? shHD.getRange(dong, cKh).getValue() : '', tong: tong,
                  thuCu: thuCu, thuMoi: thuMoi, ttCu: ttCu, ttMoi: ttMoi, tien: tien,
                  lech: (ttCu !== 'Đã thanh toán' && thuCu + tien !== tong) ? (thuCu + tien - tong) : 0 });
    if (!ghi) return;
    shHD.getRange(dong, cThu).setValue(thuMoi);
    shHD.getRange(dong, cTt).setValue(ttMoi);
    shHD.getRange(dong, cHt).setValue(ht);
    shHD.getRange(dong, cCk).setNumberFormat('@').setValue(String(d.ngaynop));
    var gcCu = String(shHD.getRange(dong, cGc).getValue() || '');
    shHD.getRange(dong, cGc).setValue(gcCu ? gcCu + ' | ' + note : note);
    var lkCu = String(shHD.getRange(dong, cLink).getValue() || '');
    shHD.getRange(dong, cLink).setValue(lkCu ? lkCu + ' | ' + linkStr : linkStr);
    dsRows.push([idDot, String(d.ngaynop), String(d.nguoinop || ''), String(d.tknhan || ''), num_(d.tongnop),
                 id, ketqua[ketqua.length - 1].khach, tien, String(x.ngaykh || ''), ttCu, ttMoi,
                 String(x.ghichu || ''), linkStr, ghiLuc, thang_(d.ngaynop)]);
  });
  if (ghi && dsRows.length) {
    shDs = ensureSheet_(ss, 'DoiSoatCK', DS_HEAD);
    var dongDau = shDs.getLastRow() + 1;
    [2, 9, 15].forEach(function (c) { shDs.getRange(dongDau, c, dsRows.length, 1).setNumberFormat('@'); });
    shDs.getRange(dongDau, 1, dsRows.length, DS_HEAD.length).setValues(dsRows);
  }
  var tongDs = 0; p.rows.forEach(function (x) { tongDs += num_(x.tien); });
  return { ok: true, ghi: ghi, added: dsRows.length, anh: links, tongDoiSoat: tongDs,
           tongNop: num_(d.tongnop), chenhLech: num_(d.tongnop) - tongDs, ketqua: ketqua };
}

// Chay tay 1 lan de cap quyen truy cap Sheet + Drive (neu deploy bao thieu quyen).
// 18/08/2026: co them DriveApp de duoc hoi quyen luu anh chung tu.
function capQuyen() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var f = layFolderAnh_();
  Logger.log('OK: ' + ss.getName() + ' | folder anh: ' + f.getName());
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
