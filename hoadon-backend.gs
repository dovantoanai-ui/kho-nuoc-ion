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

// ============================================================
// 13/10/2026 — KHUON MISA. Tab HoaDon / HoaDonCho / HoaDonHuy:
//  - 67 cot dau = DUNG ten cot file mau "Chung tu ban hang" cua MISA AMIS
//    (template_banhang.xlsx). Tai tab HoaDon ve -> MISA > Ban hang > Nhap tu excel,
//    chon dong tieu de = 1 la len duoc.
//  - Sau do la cot rieng cua app (ID, cong no, coc binh, doi soat...) - MISA bo qua.
//  - App VAN gui dong theo thu tu cu (19-20 phan tu, xem O_ ben duoi); backend tu doi
//    sang khuon MISA khi ghi va doi nguoc khi doc -> app khong phai sua.
//  - Ma hang / ma khach / ma kho / ma NV cua MISA: ke toan dien 1 lan o tab MaMISA,
//    backend tu dien vao cac dong (khong ghi de o ke toan da go tay).
//  - Coc binh KHONG dua vao phan MISA (ke toan lam tay).
// ============================================================
var MISA_HEAD = ['Hình thức bán hàng', 'Phương thức thanh toán', 'Kiêm phiếu xuất kho', 'Lập kèm hóa đơn',
  'Đã lập hóa đơn', 'Ngày hạch toán (*)', 'Ngày chứng từ (*)', 'Số chứng từ (*)', 'Số phiếu xuất', 'Mẫu số HĐ',
  'Ký hiệu HĐ', 'Số hóa đơn', 'Ngày hóa đơn', 'Mã khách hàng', 'Tên khách hàng', 'Địa chỉ', 'Mã số thuế',
  'Đơn vị giao đại lý', 'Người nộp', 'Nộp vào TK', 'Tên ngân hàng', 'Diễn giải/Lý do nộp', 'Lý do xuất',
  'Mã nhân viên bán hàng', 'Số chứng từ kèm theo (Phiếu thu)', 'Số chứng từ kèm theo (Phiếu xuất)',
  'Hạn thanh toán', 'Mã hàng (*)', 'Thuộc combo', 'Tên hàng', 'Là dòng ghi chú', 'Hàng khuyến mại',
  'Chiết khấu thương mại', 'TK Tiền/Chi phí/Nợ (*)', 'TK Doanh thu/Có (*)', 'ĐVT', 'Số lượng', 'Đơn giá',
  'Thành tiền', 'Tỷ lệ CK (%)', 'Tiền chiết khấu', 'TK chiết khấu', 'Giá tính thuế XK', '% thuế xuất khẩu',
  'Tiền thuế xuất khẩu', 'TK thuế xuất khẩu', '% thuế GTGT', '% thuế suất KHAC', 'Tiền thuế GTGT',
  'TK thuế GTGT', 'Biển kiểm soát', 'HH không TH trên tờ khai thuế GTGT', 'Mã khoản mục chi phí', 'Mã đơn vị',
  'Mã đối tượng THCP', 'Mã công trình', 'Số đơn đặt hàng', 'Số hợp đồng bán', 'Mã thống kê',
  'Số khế ước cho vay', 'CP không hợp lý', 'Mã kho', 'TK giá vốn', 'TK Kho', 'Đơn giá vốn', 'Tiền vốn',
  'Hàng hóa giữ hộ/bán hộ'];

// 01/10/2026: 'Link bill' = link anh bill chuyen tien tren Drive (doi soat CK)
var COT_LINK = 'Link bill';

// Cot rieng cua app (sau 67 cot MISA). Ten giu nhu cu de cac noi doc theo ten van dung.
// Cac cot cap HOA DON (Coc binh -> Ngày CK) chi ghi o DONG DAU moi hoa don, nhu truoc.
var APP_HEAD = ['ID hoa don', 'Ma KH app', 'SDT', 'Kho', 'Coc binh', 'Tong thanh toan', 'Da thu',
                'Trang thai TT', 'Nguoi thu', 'Hinh thuc TT', 'Ghi chu', 'Ngày CK', 'Thang', COT_LINK];

// 16/10/2026: anh Toan rut tab HoaDon con 27 cot (bo cac cot MISA khong dung).
// Backend GHI THEO TIEU DE THAT cua tab (xem ghiHoaDonRows_) -> them/bot/doi thu tu cot
// khong lam hong app. HD_HEAD chi dung khi phai TAO tab moi (vd HoaDonCho chua co).
var HD_HEAD = ['Ngày chứng từ (*)', 'Số chứng từ (*)', 'Mã khách hàng', 'Tên khách hàng', 'Tên hàng', 'ĐVT',
  'Số lượng', 'Đơn giá', 'Thành tiền', '% thuế GTGT', '% thuế suất KHAC', 'Tiền thuế GTGT', 'TK thuế GTGT']
  .concat(APP_HEAD);
// Cot BAT BUOC tren tab HoaDon / HoaDonCho / HoaDonHuy (moi cap: ten moi hoac ten cu).
// Thieu 1 cot -> DUNG, bao loi ro (thay vi ghi am tham lam mat san pham / tien).
var HD_COT_BAT_BUOC = [['ID hoa don'], ['Ngày chứng từ (*)', 'Ngay'], ['Tên khách hàng', 'Khach'],
  ['Tên hàng', 'San pham'], ['Số lượng', 'SL'], ['Đơn giá', 'Don gia'], ['Thành tiền', 'Thanh tien'],
  ['Kho'], ['Tong thanh toan'], ['Da thu'], ['Trang thai TT']];

// 01/10/2026: tab HoaDonHuy = cac dong hoa don bi huy (nhap sai / VIBA hoan hang),
// cung cau truc HoaDon + 4 cot luu vet. HoaDon chi giu hoa don dang hieu luc ->
// kho HY/HN, doanh thu, dashboard tu dung ma khong phai sua cho khac.
var HUY_HEAD = HD_HEAD.concat(['Ly do huy', 'Nguoi huy', 'Ngay huy', 'Thay bang']);

// Tai khoan hach toan mac dinh (MISA rieng ION FUJI). Ke toan dung so TK khac -> sua o day.
var MISA_TK = { tienMat: '1111', nganHang: '1121', congNo: '131', doanhThu: '5111',
                thue: '33311', giaVon: '632', kho: '156' };

// DVT app go khong dau (Thung, Binh...) -> ten DVT co dau cho cot 'ĐVT' cua MISA.
// Khoa = ten bo dau, chu thuong. DVT MISA cua ION dat khac thi sua o day. Khong co trong bang -> giu nguyen.
var DVT_MISA = { 'thung': 'Thùng', 'binh': 'Bình', 'loc': 'Lốc', 'can': 'Can', 'chai': 'Chai', 'hop': 'Hộp', 'goi': 'Gói' };

// Cot ep dang CHU khi ghi (giu so 0 dau SDT/ma, khong cho Sheet tu doi ngay/thang)
var HD_COT_CHU = ['Ngày hạch toán (*)', 'Ngày chứng từ (*)', 'Số chứng từ (*)', 'Số phiếu xuất',
  'Mã khách hàng', 'Mã nhân viên bán hàng', 'Mã hàng (*)', 'Mã kho', 'TK Tiền/Chi phí/Nợ (*)',
  'TK Doanh thu/Có (*)', 'TK thuế GTGT', 'TK giá vốn', 'TK Kho', 'ID hoa don', 'Ma KH app', 'SDT',
  'Ngày CK', 'Thang', 'Ngay huy'];

// Thu tu phan tu cua 1 dong hoa don app gui len (giu nguyen giao thuc cu):
// 0 ID,1 Ngay,2 Ma KH,3 Khach,4 SDT,5 Kho,6 San pham,7 DVT,8 SL,9 Don gia,10 Thanh tien,
// 11 VAT %,12 Coc binh,13 Tong TT,14 Da thu,15 Trang thai TT,16 Nguoi thu,17 Hinh thuc TT,
// 18 Ghi chu,19 Ngay CK,20 Thang,21 Link bill,22 Ly do huy,23 Nguoi huy,24 Ngay huy,25 Thay bang
// Truong -> ten cot (khuon MISA truoc, khuon cu sau) -> doc duoc ca tab chua chuyen khuon.
var HD_TRUONG = [
  ['ID hoa don'], ['Ngày chứng từ (*)', 'Ngay'], ['Ma KH app', 'Ma KH'], ['Tên khách hàng', 'Khach'],
  ['SDT'], ['Kho'], ['Tên hàng', 'San pham'], ['ĐVT', 'DVT'], ['Số lượng', 'SL'], ['Đơn giá', 'Don gia'],
  ['Thành tiền', 'Thanh tien'], ['% thuế GTGT', 'VAT %'], ['Coc binh'], ['Tong thanh toan'], ['Da thu'],
  ['Trang thai TT'], ['Nguoi thu'], ['Hinh thuc TT'], ['Ghi chu'], ['Ngày CK'], ['Thang'], [COT_LINK, 'Link'],
  ['Ly do huy'], ['Nguoi huy'], ['Ngay huy'], ['Thay bang']];
var HD_DAU = [3, 4, 11, 12, 13, 14, 15, 16, 17, 18, 19]; // truong chi co o dong dau hoa don

// Tab MaMISA: ke toan dien cot 'Ma MISA'. Backend tu them dong (ma de trong) khi gap ten moi.
var MM_HEAD = ['Loai', 'Ten trong app', 'Ma MISA', 'Ghi chu', 'Them luc'];
// [cot MISA can ma, Loai trong tab MaMISA]
var MM_COT = [['Mã hàng (*)', 'Hàng hóa'], ['Mã khách hàng', 'Khách hàng'], ['Mã kho', 'Kho'],
              ['Mã nhân viên bán hàng', 'Nhân viên']];

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
               'Ly do hoan', 'Ngay hoan',
               // 01/10/2026: thong tin giao thuc te tu phieu VIBA (cot 20-26)
               'So phieu VIBA', 'NV giao', 'So SP', 'Vo thu', 'Len tang', 'Ngoai thanh', 'Cong giao'];
var VB_COT_TEXT = [2, 7, 14, 17, 19, 20]; // + So phieu VIBA (giu so 0 dau: 0007208)

// 01/10/2026: KHO VIBA (kho rieng, DOC LAP voi kho HN/HY). Moi dong = 1 ma hang trong 1 phieu.
// Loai: 'Tồn đầu kỳ' | 'Nhập kho' | 'Trả về ION' | 'Điều chỉnh tồn' (SL co the am).
// Phieu kho VIBA chi cong/tru kho VIBA; hang roi kho HN/HY thi lap phieu xuat o app kho do.
// 01/10/2026: KIEM KHO VIBA = bien ban doi chieu so dem vs so sach (app tinh), KHONG sua ton.
var KKV_HEAD = ['ID dong', 'ID phieu', 'Ngay', 'Kho', 'Ma VIBA', 'Ten hang', 'Ton so sach', 'So dem',
                'Chenh lech', 'Nguoi kiem', 'Nguoi VIBA', 'Ghi chu', 'Ghi luc', 'Thang', 'Link anh'];
var KV_HEAD = ['ID dong', 'ID phieu', 'Ngay', 'Kho', 'Loai', 'Ma VIBA', 'Ten hang', 'DVT',
               'So luong', 'Nguoi', 'Ghi chu', 'Ghi luc', 'Thang', 'Link anh'];
var VBANH_FOLDER = 'Anh kho VIBA ION FUJI'; // 01/10/2026: anh bat buoc cua phieu kho + kiem kho VIBA
// 01/10/2026: cai dat don gia cong giao / chi phi VIBA - 1 dong khoa 'cfg', gia tri JSON.
var CFG_HEAD = ['Khoa', 'Gia tri', 'Cap nhat luc', 'Nguoi']; // Ngay, SDT nhan, Ngay giao, Thang, Ngay hoan: ep text de khong mat so 0 dau / bi doi kieu ngay

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
      try { out.dienMa = dienMaMisa_(ss); } catch (e3) {} // 13/10/2026: bu ma MISA ke toan vua dien
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
    } else if (p.loai === 'vibabosung' && p.id && p.tt) {
      out = boSungViba_(ss, p);
    } else if (p.loai === 'khoviba' && p.rows && p.rows.length) {
      out = ghiKhoViba_(ss, p.rows, p.anh);
    } else if (p.loai === 'kiemkhoviba' && p.rows && p.rows.length) {
      out = ghiKiemKhoViba_(ss, p.rows, p.anh);
    } else if (p.loai === 'vibadagiao' && p.rows && p.rows.length && p.viba && p.viba.row) {
      out = vibaDaGiao_(ss, p);
    } else if (p.loai === 'vibasheet' && p.tabs && p.tabs.length) {
      out = ghiSheetViba_(p);
    } else if (p.loai === 'vibacfg' && p.cfg) {
      out = ghiCfgViba_(ss, p.cfg, p.nguoi);
    } else if (p.loai === 'huyhd' && p.id) {
      out = huyHoaDon_(ss, p);
    } else if (p.loai === 'cskh' && p.cs && p.cs.id) {
      out = ghiCskh_(ss, p.cs); // 17/10/2026: bao cao cham soc khach hang
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

// Ghi cac dong hoa don vao tab name (HoaDon / HoaDonCho / HoaDonHuy).
// rows = dong theo thu tu app gui (xem HD_TRUONG). Backend doi sang khuon MISA roi ghi 1 lan.
// 16/08/2026: ID hoa don da co trong tab thi bo qua ca chum dong cua ID do (chong ghi trung).
// 13/10/2026: tim cot theo TEN (cot 1 khong con la ID); ep dang chu truoc khi ghi.
function ghiHoaDonRows_(ss, name, rows) {
  // 16/10/2026: tab DA CO -> doc tieu de that, ghi theo ten cot (cot da xoa thi bo qua).
  // Tab chua co / trang -> tao theo khuon mac dinh.
  var sh = ss.getSheetByName(name), head;
  if (sh && sh.getLastRow() >= 1 && sh.getLastColumn() > 0) {
    head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
    var thieu = HD_COT_BAT_BUOC.filter(function (ds) {
      return !ds.some(function (t) { return head.indexOf(t) >= 0; });
    }).map(function (ds) { return ds[0]; });
    if (thieu.length) throw new Error('Tab ' + name + ' thieu cot: ' + thieu.join(', ') +
                                     '. Them lai cot (dung ten) roi luu lai hoa don.');
  } else {
    head = (name === 'HoaDonHuy') ? HUY_HEAD : HD_HEAD;
    sh = ensureSheet_(ss, name, head);
  }
  var cId = head.indexOf('ID hoa don') + 1;
  var daCoId = {};
  var lastR = sh.getLastRow();
  if (lastR >= 2) {
    sh.getRange(2, cId, lastR - 1, 1).getValues().forEach(function (r0) {
      var id0 = String(r0[0] || '').trim(); if (id0) daCoId[id0] = 1;
    });
  }
  var moi = rows.filter(function (r) { return !daCoId[String(r[0] || '').trim()]; });
  if (!moi.length) return { added: 0, skipped: rows.length };
  var mm = docMaMisa_(ss);
  var data = hdSangMisa_(moi, head, mm);
  var dongDau = sh.getLastRow() + 1;
  if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
  HD_COT_CHU.forEach(function (n) {
    var c = head.indexOf(n);
    if (c >= 0) sh.getRange(dongDau, c + 1, data.length, 1).setNumberFormat('@');
  });
  sh.getRange(dongDau, 1, data.length, head.length).setValues(data);
  if (name === 'HoaDon') {
    try { themKhoaMaMisa_(ss, mm, moi); } catch (e) {} // loi tab MaMISA khong lam hong hoa don
  }
  return { added: moi.length, skipped: rows.length - moi.length };
}

// Doi cac dong (thu tu app) sang khuon tab (head = HD_HEAD hoac HUY_HEAD).
// Thong tin cap hoa don (khach, VAT, thanh toan...) lay o dong DAU cua moi ID trong rows.
var TEN_CU_ = { 'Ngày chứng từ (*)': 'Ngay', 'Tên khách hàng': 'Khach', 'Tên hàng': 'San pham',
  'ĐVT': 'DVT', 'Số lượng': 'SL', 'Đơn giá': 'Don gia', 'Thành tiền': 'Thanh tien', '% thuế GTGT': 'VAT %',
  'Ma KH app': 'Ma KH' };
function hdSangMisa_(rows, head, mm) {
  var c = {}; head.forEach(function (n, i) { c[n] = i; });
  var chu = {}; HD_COT_CHU.forEach(function (n) { if (c[n] != null) chu[c[n]] = 1; });
  var dau = {};
  rows.forEach(function (r) { var id = String(r[0] || '').trim(); if (!dau[id]) dau[id] = r; });
  return rows.map(function (r) {
    var id = String(r[0] || '').trim(), h = dau[id], laDau = (h === r);
    var o = head.map(function () { return ''; });
    function set(n, v) { // 16/10/2026: tab con ten cot cu thi ghi vao ten cu
      var k = (c[n] != null) ? n : (TEN_CU_[n] && c[TEN_CU_[n]] != null ? TEN_CU_[n] : null);
      if (k != null) o[c[k]] = (v == null ? '' : v);
    }
    var khach = String(h[3] || '').trim(), tt = String(h[15] || '').trim(), ht = String(h[17] || '').trim();
    var vat = num_(h[11]), sl = num_(r[8]), tien = num_(r[10]), ngay = ngayChu_(r[1]);
    var pt = ptThanhToan_(tt, ht);
    var dien = 'Bán hàng ' + khach + ' [' + id + ']';
    if (tt === 'Thanh toán 1 phần') dien += ' - đã thu ' + num_(h[14]);
    // --- phan MISA (lap lai o moi dong cua phieu) ---
    set('Hình thức bán hàng', 'Bán hàng hóa trong nước');
    set('Phương thức thanh toán', pt[0]);
    set('Kiêm phiếu xuất kho', 'Có');
    set('Lập kèm hóa đơn', 'Không');
    set('Ngày hạch toán (*)', ngay);
    set('Ngày chứng từ (*)', ngay);
    set('Số chứng từ (*)', id);
    set('Số phiếu xuất', /^HD/.test(id) ? id.replace(/^HD/, 'XK') : 'XK' + id);
    set('Mã khách hàng', maMisa_(mm, 'Khách hàng', r[2] || h[2]));
    set('Tên khách hàng', khach);
    set('Diễn giải/Lý do nộp', dien);
    set('Lý do xuất', 'Xuất bán ' + khach + ' [' + id + ']');
    set('Mã nhân viên bán hàng', maMisa_(mm, 'Nhân viên', h[16]));
    set('Mã hàng (*)', maMisa_(mm, 'Hàng hóa', r[6]));
    set('Tên hàng', r[6]);
    set('Là dòng ghi chú', 'Không');
    set('Hàng khuyến mại', (tien === 0 && sl > 0) ? 'Có' : 'Không'); // dong 0d = hang tang
    set('TK Tiền/Chi phí/Nợ (*)', pt[1]);
    set('TK Doanh thu/Có (*)', MISA_TK.doanhThu);
    set('ĐVT', DVT_MISA[chuanTen_(r[7])] || r[7]);
    set('Số lượng', sl);
    set('Đơn giá', num_(r[9]));
    set('Thành tiền', tien);
    if (vat > 0) {
      set('% thuế GTGT', vat);
      set('Tiền thuế GTGT', Math.round(tien * vat / 100));
      set('TK thuế GTGT', MISA_TK.thue);
    }
    set('Mã kho', maMisa_(mm, 'Kho', r[5]));
    set('TK giá vốn', MISA_TK.giaVon);
    set('TK Kho', MISA_TK.kho);
    // --- phan rieng app ---
    set('ID hoa don', id);
    set('Ma KH app', String(r[2] || '').trim());
    set('Kho', r[5]);
    set('Thang', thang_(r[1]));
    set(COT_LINK, r[21]);
    if (laDau) {
      set('SDT', sdtChu_(h[4]));
      set('Coc binh', h[12]); set('Tong thanh toan', h[13]); set('Da thu', h[14]);
      set('Trang thai TT', h[15]); set('Nguoi thu', h[16]); set('Hinh thuc TT', h[17]);
      set('Ghi chu', h[18]); set('Ngày CK', h[19] ? ngayChu_(h[19]) : '');
    }
    set('Ly do huy', r[22]); set('Nguoi huy', r[23]);
    set('Ngay huy', r[24] ? ngayChu_(r[24]) : ''); set('Thay bang', r[25]);
    return o.map(function (v, i) { return (chu[i] && v !== '') ? String(v) : v; });
  });
}

// Trang thai + hinh thuc TT luc ban -> [Phuong thuc thanh toan MISA, TK No].
// Chua thu / thu 1 phan -> 'Chua thu tien' (phan da thu ke toan lap phieu thu rieng).
// Sau nay thu no / doi soat chi sua cot app, KHONG doi phan MISA (chung tu ban giu nguyen).
function ptThanhToan_(tt, ht) {
  if (tt === 'Đã thanh toán') {
    if (!ht || ht === 'Tiền mặt') return ['Thu tiền ngay - Tiền mặt', MISA_TK.tienMat];
    return ['Thu tiền ngay - Chuyển khoản', MISA_TK.nganHang];
  }
  return ['Chưa thu tiền', MISA_TK.congNo];
}

// Ngay (Date / dd/MM/yyyy / yyyy-MM-dd) -> chuoi 'dd/MM/yyyy' nhu mau MISA
function ngayChu_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'GMT+7', 'dd/MM/yyyy');
  var s = String(v == null ? '' : v).trim(), m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return ('0' + m[3]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[1];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[3];
  return s;
}

// SDT bi Sheet doi thanh so (mat 0 dau, vd 349650229) -> them lai so 0
function sdtChu_(v) {
  var s = String(v == null ? '' : v).trim();
  if (typeof v === 'number' && /^\d{9}$/.test(s)) s = '0' + s;
  return s;
}

// Ban do ten cot -> vi tri (cot dau tien trung ten)
function cotMap_(head) {
  var m = {};
  head.forEach(function (x, i) { var k = String(x).trim(); if (m[k] == null) m[k] = i; });
  return m;
}

// 1 dong cua tab (khuon MISA hoac khuon cu) -> dong theo thu tu app (26 phan tu).
// laDau = dong dau tien cua ID: chi dong dau mang khach/SDT/VAT/thanh toan (nhu app gui).
function hdCu_(m, r, laDau) {
  var o = HD_TRUONG.map(function (ten) {
    for (var k = 0; k < ten.length; k++) if (m[ten[k]] != null) return r[m[ten[k]]];
    return '';
  });
  o[0] = String(o[0] || '').trim();
  o[1] = ngayChu_(o[1]);
  if (!laDau) HD_DAU.forEach(function (i) { o[i] = ''; });
  return o;
}

// Doc (CHUA xoa) moi dong cua 1 ID hoa don trong tab name.
// Tra ve { sh, rows (thu tu app, 22 phan tu 0..21 - xem HD_TRUONG), dong }.
// Goi xoaDong_ SAU khi da ghi xong noi moi: buoc ghi loi giua chung thi du lieu cu con nguyen.
function docHoaDonId_(ss, name, id) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return { sh: sh, rows: [], dong: [] };
  var v = sh.getDataRange().getValues();
  var m = cotMap_(v[0]);
  var cId = m['ID hoa don'];
  var rows = [], dong = [];
  if (cId == null) return { sh: sh, rows: rows, dong: dong };
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][cId] || '').trim() !== id) continue;
    // chi lay 22 phan tu dau (0..21 = den Link bill) nhu truoc: noi goi se concat 4 cot luu vet huy vao sau
    rows.push(hdCu_(m, v[i], rows.length === 0).slice(0, 22)); dong.push(i + 1);
  }
  return { sh: sh, rows: rows, dong: dong };
}

// ---------- MA MISA (tab MaMISA) ----------
function chuanTen_(s) {
  s = String(s == null ? '' : s).toLowerCase();
  try { s = s.normalize('NFD').replace(/[̀-ͯ]/g, ''); } catch (e) {}
  return s.replace(/đ/g, 'd').replace(/\s+/g, ' ').trim();
}
// { ma: {'Loai|ten chuan': 'Ma MISA'}, co: {'Loai|ten chuan': 1} }
function docMaMisa_(ss) {
  var out = { ma: {}, co: {} };
  var d = rowsOf_(ss, 'MaMISA'); if (!d.head.length) return out;
  var iL = idx_(d.head, 'Loai'), iT = idx_(d.head, 'Ten trong app'), iM = idx_(d.head, 'Ma MISA');
  if (iL < 0 || iT < 0 || iM < 0) return out;
  d.rows.forEach(function (r) {
    var k = String(r[iL] || '').trim() + '|' + chuanTen_(r[iT]);
    out.co[k] = 1;
    var ma = String(r[iM] == null ? '' : r[iM]).trim();
    if (ma) out.ma[k] = ma;
  });
  return out;
}
function maMisa_(mm, loai, ten) {
  if (!mm || ten == null || String(ten).trim() === '') return '';
  return mm.ma[loai + '|' + chuanTen_(ten)] || '';
}
// Ten moi chua co trong MaMISA -> them 1 dong (Ma MISA de trong cho ke toan dien)
function themKhoaMaMisa_(ss, mm, rows) {
  var dau = {}, moi = [], luc = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm');
  function them(loai, ten, gc) {
    ten = String(ten == null ? '' : ten).trim(); if (!ten) return;
    var k = loai + '|' + chuanTen_(ten); if (mm.co[k]) return;
    mm.co[k] = 1; moi.push([loai, ten, '', String(gc || ''), luc]);
  }
  rows.forEach(function (r) {
    var id = String(r[0] || '').trim(); if (!dau[id]) dau[id] = r;
    var h = dau[id];
    them('Hàng hóa', r[6], r[7] ? 'ĐVT: ' + r[7] : '');
    them('Khách hàng', r[2] || h[2], h[3]);
    them('Kho', r[5], '');
    them('Nhân viên', h[16], '');
  });
  if (!moi.length) return 0;
  var sh = ensureSheet_(ss, 'MaMISA', MM_HEAD);
  var dong = sh.getLastRow() + 1;
  sh.getRange(dong, 2, moi.length, 2).setNumberFormat('@');
  sh.getRange(dong, 1, moi.length, MM_HEAD.length).setValues(moi);
  return moi.length;
}
// Dien ma MISA con trong o tab HoaDon theo tab MaMISA. KHONG ghi de o da co ma.
// Chay sau moi lan ghi hoa don, hoac mo link <URL exec>?lenh=dienma sau khi ke toan dien MaMISA.
function dienMaMisa_(ss) {
  var sh = ss.getSheetByName('HoaDon');
  if (!sh || sh.getLastRow() < 2) return 0;
  var mm = docMaMisa_(ss);
  if (!Object.keys(mm.ma).length) return 0;
  var n = sh.getLastRow() - 1;
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
  var m = cotMap_(head);
  function cot(t) { return m[t] == null ? null : sh.getRange(2, m[t] + 1, n, 1).getValues(); }
  var ids = cot('ID hoa don'), khoa = { 'Hàng hóa': cot('Tên hàng'), 'Khách hàng': cot('Ma KH app'),
    'Kho': cot('Kho'), 'Nhân viên': cot('Nguoi thu') };
  if (!ids) return 0;
  var nvTheoId = {};
  for (var i = 0; i < n; i++) {
    var id = String(ids[i][0] || '').trim(), nv = khoa['Nhân viên'] ? String(khoa['Nhân viên'][i][0] || '').trim() : '';
    if (id && nv && !nvTheoId[id]) nvTheoId[id] = nv;
  }
  var dem = 0;
  MM_COT.forEach(function (x) {
    var tenCot = x[0], loai = x[1];
    if (m[tenCot] == null || !khoa[loai]) return;
    var rg = sh.getRange(2, m[tenCot] + 1, n, 1), v = rg.getValues(), doi = false;
    for (var i = 0; i < n; i++) {
      if (String(v[i][0] == null ? '' : v[i][0]).trim() !== '') continue;
      var ten = (loai === 'Nhân viên') ? nvTheoId[String(ids[i][0] || '').trim()] : khoa[loai][i][0];
      var ma = maMisa_(mm, loai, ten);
      if (ma) { v[i][0] = ma; doi = true; dem++; }
    }
    if (doi) rg.setNumberFormat('@').setValues(v);
  });
  return dem;
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
  var out = { ok: true, khach: [], hoadon: [], thanhtoan: [], viba: [], cho: [], huy: [], khoviba: [], kiemkhoviba: [], vbcfg: null };
  // 13/10/2026: mo <URL exec>?lenh=dienma sau khi ke toan dien tab MaMISA -> dien ma cho cac dong cu
  if (e && e.parameter && e.parameter.lenh === 'dienma') {
    var lk = LockService.getScriptLock();
    try { lk.waitLock(30000); out = { ok: true, dienMa: dienMaMisa_(SpreadsheetApp.openById(SHEET_ID)) }; }
    catch (err0) { out = { ok: false, error: err0.message }; }
    finally { try { lk.releaseLock(); } catch (e4) {} }
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    out.khach = readKhach_(ss);
    out.hoadon = readHoaDon_(ss);
    out.thanhtoan = readThanhToan_(ss);
    out.viba = readViba_(ss);
    out.cho = readHoaDon_(ss, 'HoaDonCho'); // hang dang di giao VIBA (chua vao HoaDon)
    out.huy = readHoaDon_(ss, 'HoaDonHuy'); // 01/10/2026: hoa don da huy -> app gan nhan 'Da huy'
    out.cat = readCatLe_();
    out.khoviba = readKhoViba_(ss); // 01/10/2026
    out.kiemkhoviba = readKiemKhoViba_(ss);
    out.vbcfg = readCfgViba_(ss);
    try { out.cskh = readCskh_(ss); } catch (e5) { out.cskh = []; } // 17/10/2026
    var vbid = PropertiesService.getScriptProperties().getProperty(VBSHEET_KEY);
    if (vbid) out.vbSheet = 'https://docs.google.com/spreadsheets/d/' + vbid + '/edit';
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ============================================================
// 17/10/2026 — CHAM SOC KHACH HANG (tab CSKH). Moi lan NV bao cao = 1 dong.
// 'Bam goi luc' / 'Roi app (giay)': app ghi luc NV bam nut Goi/Zalo trong app va so giay
// roi app sau do (UOC LUONG thoi luong cuoc goi - app web khong doc duoc nhat ky cuoc goi).
// ============================================================
var CS_HEAD = ['ID', 'Thoi gian', 'Ngay', 'Ma KH', 'Ten KH', 'SDT', 'Kho', 'Nguoi CS', 'Kenh',
               'Ket qua', 'Ghi chu', 'Hen goi lai', 'Bam goi luc', 'Roi app (giay)', 'So lan bam Goi',
               'So lan bam Zalo', 'Thang'];
function ghiCskh_(ss, c) {
  var sh = ensureSheet_(ss, 'CSKH', CS_HEAD);
  var id = String(c.id).trim();
  if (sh.getLastRow() >= 2) {
    var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === id) return { ok: true, skipped: 1 };
  }
  var luc = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm');
  var ngay = String(c.ngay || luc.slice(0, 10));
  var row = [id, luc, ngay, String(c.ma || ''), String(c.ten || ''), sdtChu_(c.sdt), String(c.kho || ''),
             String(c.nguoi || ''), String(c.kenh || ''), String(c.ketqua || ''), String(c.ghichu || ''),
             String(c.hen || ''), String(c.goiluc || ''), num_(c.roi), num_(c.sogoi), num_(c.sozalo), thang_(ngay)];
  var dong = sh.getLastRow() + 1;
  [1, 2, 3, 4, 6, 12, 13, 17].forEach(function (k) { sh.getRange(dong, k).setNumberFormat('@'); });
  sh.getRange(dong, 1, 1, CS_HEAD.length).setValues([row]);
  return { ok: true, added: 1, id: id };
}
function readCskh_(ss) {
  var d = rowsOf_(ss, 'CSKH'); if (!d.head.length) return [];
  var m = cotMap_(d.head);
  function g(r, n) { return m[n] == null ? '' : r[m[n]]; }
  return d.rows.filter(function (r) { return String(g(r, 'ID')).trim(); }).map(function (r) {
    return { id: String(g(r, 'ID')), luc: String(g(r, 'Thoi gian')), ngay: fmtD_(g(r, 'Ngay')), ma: String(g(r, 'Ma KH')),
             ten: String(g(r, 'Ten KH')), kho: String(g(r, 'Kho')), nguoi: String(g(r, 'Nguoi CS')), kenh: String(g(r, 'Kenh')),
             ketqua: String(g(r, 'Ket qua')), ghichu: String(g(r, 'Ghi chu')), hen: g(r, 'Hen goi lai') ? fmtD_(g(r, 'Hen goi lai')) : '',
             roi: num_(g(r, 'Roi app (giay)')) };
  });
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
                    '", backend can "' + head[i] + '". Khong ghi de - sua lai tieu de cot roi thu lai.' +
                    (eh[0] === 'ID hoa don' ? ' (Tab dang o khuon cu: chay ham chuyenKhuonMISA_20261013 1 lan.)' : ''));
  }
  if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
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
  if (typeof v === 'number') return isNaN(v) ? 0 : v; // 13/10/2026: so that (vd 8.5) giu nguyen, khong bo dau cham
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
// 13/10/2026: doc theo TEN cot qua hdCu_ -> doc duoc ca khuon MISA lan khuon cu.
// Tra ve dung cau truc cu: khach/SDT/VAT/thanh toan chi co o dong dau moi hoa don.
function readHoaDon_(ss, name) {
  var d = rowsOf_(ss, name || 'HoaDon'); if (!d.head.length) return [];
  var m = cotMap_(d.head);
  if (m['ID hoa don'] == null) return [];
  function s(v) { return String(v == null ? '' : v); }
  var daGap = {}, out = [];
  d.rows.forEach(function (r0) {
    var id = String(r0[m['ID hoa don']] || '').trim(); if (!id) return;
    var r = hdCu_(m, r0, !daGap[id]); daGap[id] = 1;
    out.push({
      id: id, ngay: r[1], ma: String(r[2] || '').trim(),
      khach: r[3] || '', sdt: sdtChu_(r[4]), kho: r[5] || '',
      sp: r[6] || '', dvt: r[7] || '', sl: num_(r[8]), gia: num_(r[9]),
      tien: num_(r[10]), vat: num_(r[11]), coc: num_(r[12]),
      tong: num_(r[13]), dathu: num_(r[14]),
      tt: r[15] || '', thu: r[16] || '', ht: r[17] || '', ghichu: r[18] || '',
      ngayck: r[19] ? fmtD_(r[19]) : '', link: s(r[21]),
      lydo: s(r[22]), nguoihuy: s(r[23]), ngayhuy: r[24] ? fmtD_(r[24]) : '', thaythe: s(r[25])
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
  while (row.length < VB_HEAD.length) row.push('');                  // 01/10/2026: cot 20-26 (thong tin phieu VIBA) de trong -> du so cot

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
  if (!ss.getSheetByName('GiaoVIBA')) return { ok: false, error: 'Chua co tab GiaoVIBA' };
  var sh = ensureSheet_(ss, 'GiaoVIBA', VB_HEAD); // 01/10/2026: tu them 7 cot thong tin giao
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
  if (p.tt) ghiTtViba_(sh, dong, p.tt);
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

// ============================================================
// 01/10/2026 — THONG TIN GIAO VIBA + KHO VIBA + CAI DAT
// ============================================================
// Ghi 7 cot thong tin giao thuc te (so phieu, NV, so SP, vo thu, len tang, ngoai thanh, cong)
function ghiTtViba_(sh, dong, t) {
  var c0 = VB_HEAD.indexOf('So phieu VIBA') + 1;
  sh.getRange(dong, c0).setNumberFormat('@');
  sh.getRange(dong, c0, 1, 7).setValues([[
    String(t.sophieu || ''), String(t.nv || ''), num_(t.sosp), num_(t.vothu),
    t.tang ? 'x' : '', t.ngoai ? 'x' : '', num_(t.cong)
  ]]);
}

// Bo sung / sua thong tin giao cho don DA GIAO (don giao truoc khi co tinh nang nay)
function boSungViba_(ss, p) {
  var id = String(p.id).trim();
  if (!ss.getSheetByName('GiaoVIBA')) return { ok: false, error: 'Chua co tab GiaoVIBA' };
  var sh = ensureSheet_(ss, 'GiaoVIBA', VB_HEAD);
  var dong = timDongViba_(sh, id);
  if (dong < 0) return { ok: false, error: 'Khong tim thay don ' + id };
  var tt = String(sh.getRange(dong, 13).getValue() || '');
  if (tt !== 'Đã giao') return { ok: false, error: 'Don ' + id + ' chua giao - dung nut Xac nhan da giao' };
  ghiTtViba_(sh, dong, p.tt);
  sh.getRange(dong, 16).setValue(Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm'));
  return { ok: true, capnhat: 1 };
}

// Ghi phieu kho VIBA. rows: [ID dong, ID phieu, Ngay, Kho, Loai, Ma VIBA, Ten hang, DVT, So luong, Nguoi, Ghi chu]
// Chong ghi trung theo 'ID dong' (gui lai cung phieu khong nhan doi).
function ghiKhoViba_(ss, rows, anh) {
  var sh = ensureSheet_(ss, 'KhoVIBA', KV_HEAD);
  var daCo = {}, last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { daCo[String(r[0] || '').trim()] = 1; });
  var luc = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm');
  var moi = [];
  rows.forEach(function (r) {
    var key = String(r[0] || '').trim();
    if (!key || daCo[key]) return; daCo[key] = 1;
    var x = r.slice(0, 11); while (x.length < 11) x.push('');
    x[8] = num_(x[8]); // num_ giu dau am (dieu chinh kiem ke)
    x.push(luc); x.push(thang_(x[2]));
    moi.push(x.map(function (v, i) { return (i === 2 || i === 12 || i === 0 || i === 1) ? String(v == null ? '' : v) : v; }));
  });
  if (!moi.length) return { ok: true, added: 0, skipped: rows.length };
  // Upload anh 1 lan cho ca phieu (chi khi co dong moi -> gui lai khong tao anh trung)
  var link = (anh && anh.data) ? uploadAnh_(anh.data, anh.mime, anh.ten, VBANH_FOLDER) : '';
  moi.forEach(function (x) { x.push(link); });
  var dau = sh.getLastRow() + 1;
  [1, 2, 3, 13].forEach(function (c) { sh.getRange(dau, c, moi.length, 1).setNumberFormat('@'); });
  sh.getRange(dau, 1, moi.length, KV_HEAD.length).setValues(moi);
  return { ok: true, added: moi.length, skipped: rows.length - moi.length };
}

function readKhoViba_(ss) {
  var d = rowsOf_(ss, 'KhoVIBA'); if (!d.head.length) return [];
  var h = d.head, c = {};
  KV_HEAD.forEach(function (n) { c[n] = idx_(h, n); });
  function g(r, n) { return c[n] >= 0 ? r[c[n]] : ''; }
  var out = [];
  d.rows.forEach(function (r) {
    var idd = String(g(r, 'ID dong') || '').trim(); if (!idd) return;
    var slRaw = g(r, 'So luong'); var sl = (typeof slRaw === 'number') ? slRaw : num_(slRaw);
    out.push({
      idd: idd, id: String(g(r, 'ID phieu') || '').trim(), ngay: fmtD_(g(r, 'Ngay')),
      kho: g(r, 'Kho') || '', loai: g(r, 'Loai') || '', ma: String(g(r, 'Ma VIBA') || '').trim(),
      ten: g(r, 'Ten hang') || '', dvt: g(r, 'DVT') || '', sl: sl,
      nguoi: g(r, 'Nguoi') || '', gc: g(r, 'Ghi chu') || '', link: String(g(r, 'Link anh') || '')
    });
  });
  return out;
}

// Ghi bien ban kiem kho VIBA. rows: [ID dong, ID phieu, Ngay, Kho, Ma VIBA, Ten hang,
// Ton so sach, So dem, Chenh lech, Nguoi kiem, Nguoi VIBA, Ghi chu]. Chong ghi trung theo ID dong.
function ghiKiemKhoViba_(ss, rows, anh) {
  var sh = ensureSheet_(ss, 'KiemKhoVIBA', KKV_HEAD);
  var daCo = {}, last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { daCo[String(r[0] || '').trim()] = 1; });
  var luc = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm');
  var moi = [];
  rows.forEach(function (r) {
    var key = String(r[0] || '').trim();
    if (!key || daCo[key]) return; daCo[key] = 1;
    var x = r.slice(0, 12); while (x.length < 12) x.push('');
    x[6] = num_(x[6]); x[7] = num_(x[7]); x[8] = num_(x[8]);
    x.push(luc); x.push(thang_(x[2]));
    moi.push(x.map(function (v, i) { return (i <= 2 || i === 13) ? String(v == null ? '' : v) : v; }));
  });
  if (!moi.length) return { ok: true, added: 0, skipped: rows.length };
  var link = (anh && anh.data) ? uploadAnh_(anh.data, anh.mime, anh.ten, VBANH_FOLDER) : '';
  moi.forEach(function (x) { x.push(link); });
  var dau = sh.getLastRow() + 1;
  [1, 2, 3, 14].forEach(function (c) { sh.getRange(dau, c, moi.length, 1).setNumberFormat('@'); });
  sh.getRange(dau, 1, moi.length, KKV_HEAD.length).setValues(moi);
  return { ok: true, added: moi.length, skipped: rows.length - moi.length };
}

function readKiemKhoViba_(ss) {
  var d = rowsOf_(ss, 'KiemKhoVIBA'); if (!d.head.length) return [];
  var h = d.head, c = {};
  KKV_HEAD.forEach(function (n) { c[n] = idx_(h, n); });
  function g(r, n) { return c[n] >= 0 ? r[c[n]] : ''; }
  function so(v) { return (typeof v === 'number') ? v : num_(v); }
  var out = [];
  d.rows.forEach(function (r) {
    var idd = String(g(r, 'ID dong') || '').trim(); if (!idd) return;
    out.push({
      idd: idd, id: String(g(r, 'ID phieu') || '').trim(), ngay: fmtD_(g(r, 'Ngay')), kho: g(r, 'Kho') || '',
      ma: String(g(r, 'Ma VIBA') || '').trim(), ten: g(r, 'Ten hang') || '',
      so: so(g(r, 'Ton so sach')), dem: so(g(r, 'So dem')), lech: so(g(r, 'Chenh lech')),
      nguoi: g(r, 'Nguoi kiem') || '', vb: g(r, 'Nguoi VIBA') || '', gc: g(r, 'Ghi chu') || '', link: String(g(r, 'Link anh') || '')
    });
  });
  return out;
}

// 01/10/2026 ban 3: don VIBA nhap SAU KHI DA GIAO (tab Giao VIBA > Nhap don).
// Ghi 1 lan: GiaoVIBA (trang thai 'Đã giao' + thong tin phieu) + HoaDon (ngay = ngay giao).
// Gui lai cung ID: GiaoVIBA va HoaDon deu chong trung theo ID -> an toan.
function vibaDaGiao_(ss, p) {
  var id = String(p.viba.row[0] || '').trim();
  if (!id) return { ok: false, error: 'Thieu ID hoa don' };
  var kq = ghiViba_(ss, p.viba.row);
  var sh = ensureSheet_(ss, 'GiaoVIBA', VB_HEAD);
  var dong = timDongViba_(sh, id);
  if (dong < 0) return { ok: false, error: 'Khong ghi duoc GiaoVIBA ' + id };
  if (String(sh.getRange(dong, 13).getValue() || '') === 'Hoàn hàng') return { ok: false, error: 'Don ' + id + ' da hoan hang' };
  var hd = ghiHoaDonRows_(ss, 'HoaDon', p.rows);
  sh.getRange(dong, 14).setNumberFormat('@');
  sh.getRange(dong, 13, 1, 4).setValues([[
    'Đã giao', String(p.ngaygiao || p.viba.row[1] || ''), String(p.nguoi || ''),
    Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm')
  ]]);
  if (p.tt) ghiTtViba_(sh, dong, p.tt);
  return { ok: true, added: hd.added, skipped: hd.skipped, viba: kq };
}

// 01/10/2026 ban 7: SHEET RIENG CHO NHAN VIEN VIBA (chi xem).
// App tinh so (cung cong thuc bao cao) roi gui {loai:'vibasheet', tabs:[{ten, rows, dam, rong}]}.
// Lan dau tu tao file (dat canh sheet goc), luu ID vao Script Properties. Moi lan ghi de toan bo tab.
var VBSHEET_KEY = 'VIBA_SHEET_ID';
var VBSHEET_TEN = 'VIBA - Kho ION FUJI Ha Noi';
function sheetViba_() {
  var pr = PropertiesService.getScriptProperties(), id = pr.getProperty(VBSHEET_KEY);
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  var s = SpreadsheetApp.create(VBSHEET_TEN);
  try {
    var cha = DriveApp.getFileById(SHEET_ID).getParents();
    if (cha.hasNext()) DriveApp.getFileById(s.getId()).moveTo(cha.next());
  } catch (e) {}
  pr.setProperty(VBSHEET_KEY, s.getId());
  return s;
}
function ghiSheetViba_(p) {
  var s = sheetViba_(), ten = [];
  p.tabs.forEach(function (t, i) {
    var name = String(t.ten || ('Tab ' + (i + 1))).slice(0, 90); ten.push(name);
    var sh = s.getSheetByName(name) || s.insertSheet(name, i);
    sh.clear();
    var rows = t.rows || []; if (!rows.length) return;
    var w = 1; rows.forEach(function (r) { if (r.length > w) w = r.length; });
    var data = rows.map(function (r) {
      var o = [];
      for (var j = 0; j < w; j++) {
        var v = r[j]; if (v === null || v === undefined) v = '';
        // chuoi bat dau bang so (so phieu 0007208, ngay 11/09/2026) -> giu nguyen dang chu
        if (typeof v === 'string' && /^[0-9+\-]/.test(v)) v = "'" + v;
        o.push(v);
      }
      return o;
    });
    if (sh.getMaxColumns() < w) sh.insertColumnsAfter(sh.getMaxColumns(), w - sh.getMaxColumns());
    var rg = sh.getRange(1, 1, data.length, w);
    rg.setValues(data).setNumberFormat('#,##0').setFontFamily('Arial').setFontSize(10);
    sh.getRange(1, 1).setFontSize(13).setFontWeight('bold');
    (t.dam || []).forEach(function (r) { if (r < data.length) sh.getRange(r + 1, 1, 1, w).setFontWeight('bold').setBackground('#FBEAEA'); });
    (t.rong || []).forEach(function (px, j) { try { sh.setColumnWidth(j + 1, Math.round(px * 7.5)); } catch (e) {} });
  });
  s.getSheets().forEach(function (sh) { if (ten.indexOf(sh.getName()) < 0 && s.getSheets().length > 1) s.deleteSheet(sh); });
  return { ok: true, id: s.getId(), url: s.getUrl() };
}

function readCfgViba_(ss) {
  var d = rowsOf_(ss, 'CaiDatVIBA'); if (!d.head.length) return null;
  for (var i = 0; i < d.rows.length; i++) {
    if (String(d.rows[i][0]).trim() === 'cfg') {
      try { return JSON.parse(String(d.rows[i][1])); } catch (e) { return null; }
    }
  }
  return null;
}

function ghiCfgViba_(ss, cfg, nguoi) {
  var sh = ensureSheet_(ss, 'CaiDatVIBA', CFG_HEAD);
  var val = [['cfg', JSON.stringify(cfg), Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm'), String(nguoi || '')]];
  var last = sh.getLastRow();
  if (last > 1) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === 'cfg') { sh.getRange(i + 2, 1, 1, 4).setValues(val); return { ok: true, capnhat: 1 }; }
    }
  }
  sh.getRange(last + 1, 1, 1, 4).setValues(val);
  return { ok: true, added: 1 };
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
      lydo: g(r, 'Ly do hoan') || '', ngayhoan: fmtD_(g(r, 'Ngay hoan')),
      // 01/10/2026: thong tin giao thuc te
      sophieu: String(g(r, 'So phieu VIBA') || ''), nv: String(g(r, 'NV giao') || ''),
      sosp: num_(g(r, 'So SP')), vothu: num_(g(r, 'Vo thu')),
      tang: String(g(r, 'Len tang') || '') !== '', ngoai: String(g(r, 'Ngoai thanh') || '') !== '',
      cong: num_(g(r, 'Cong giao')), cotTt: String(g(r, 'NV giao') || '') !== ''
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
    var dongMoi = timDongDau_(shHD, idMoi, cotTheoTen_(shHD, 'ID hoa don'));
    var dau = cu.rows[0];
    // Mang sang phieu moi: Ngay CK + Link bill da doi soat cua phieu cu (neu co)
    if (dongMoi > 0 && dau[19]) shHD.getRange(dongMoi, cotTheoTen_(shHD, 'Ngày CK')).setNumberFormat('@').setValue(ngayChu_(dau[19]));
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

// So dong (>=2) dau tien cua ma id trong cot col (mac dinh cot 1), -1 neu khong co.
// 13/10/2026: tab HoaDon khuon MISA -> ID hoa don KHONG o cot 1, phai truyen col.
function timDongDau_(sh, id, col) {
  if (!sh || sh.getLastRow() < 2) return -1;
  var ids = sh.getRange(2, col || 1, sh.getLastRow() - 1, 1).getValues();
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
      cHt = cot('Hinh thuc TT'), cGc = cot('Ghi chu'), cCk = cot('Ngày CK'),
      cKh = cot('Tên khách hàng') || cot('Khach'), cId = cot('ID hoa don'); // 13/10/2026: khuon MISA
  if (cId < 1 || cTong < 1 || cThu < 1 || cTt < 1 || cHt < 1 || cGc < 1 || cCk < 1) return { ok: false, error: 'HoaDon thieu cot can thiet' };

  var ids = shHD.getRange(2, cId, Math.max(shHD.getLastRow() - 1, 1), 1).getValues();
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
    if (p.giuHt) note = 'Nộp TK công ty ' + d.ngaynop + ' · đợt ' + idDot +
                        (d.nguoinop ? ' · ' + d.nguoinop : '') + (x.ghichu ? ', ' + x.ghichu : '');
    ketqua.push({ id: id, khach: cKh > 0 ? shHD.getRange(dong, cKh).getValue() : '', tong: tong,
                  thuCu: thuCu, thuMoi: thuMoi, ttCu: ttCu, ttMoi: ttMoi, tien: tien,
                  lech: (ttCu !== 'Đã thanh toán' && thuCu + tien !== tong) ? (thuCu + tien - tong) : 0 });
    if (!ghi) return;
    shHD.getRange(dong, cThu).setValue(thuMoi);
    shHD.getRange(dong, cTt).setValue(ttMoi);
    if (!p.giuHt) shHD.getRange(dong, cHt).setValue(ht); // 16/10/2026: nop tien ve TK cong ty -> giu Hinh thuc TT goc
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
// CHAY TAY 1 LAN (13/10/2026): chuyen HoaDon / HoaDonCho / HoaDonHuy sang KHUON MISA.
// CACH CHAY: dan code -> Deploy New version -> chon ham chuyenKhuonMISA_20261013 -> Run
//            (lam luc khong ai dang lap hoa don). Xem ket qua o Nhat ky thuc thi.
// AN TOAN:
//  - Tab cu KHONG bi xoa/sua: doi ten thanh '<ten>_truocMISA_20261013' (ban sao luu).
//  - Du lieu moi ghi vao tab moi, so dong + tong Thanh tien / Tong TT / Da thu phai KHOP
//    voi tab cu thi moi doi ten, lech 1 dong la DUNG, khong doi gi.
//  - Chay lai lan 2: tab da o khuon MISA thi bo qua.
// ============================================================
function chuyenKhuonMISA_20261013() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // chan app ghi xen vao giua luc chuyen
  var bao = [];
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var mm = docMaMisa_(ss);
    ['HoaDon', 'HoaDonCho', 'HoaDonHuy'].forEach(function (name) {
      var sh = ss.getSheetByName(name);
      if (!sh) { bao.push(name + ': khong co tab, bo qua'); return; }
      var v = sh.getDataRange().getValues();
      var m = cotMap_(v[0] || []);
      if (m['Số chứng từ (*)'] != null) { bao.push(name + ': da o khuon MISA, bo qua'); return; }
      if (m['ID hoa don'] == null) throw new Error(name + ': khong thay cot "ID hoa don" - DUNG, khong doi gi');
      var head = (name === 'HoaDonHuy') ? HUY_HEAD : HD_HEAD;
      var rows = [], daGap = {}, cu = { dong: 0, tien: 0, tong: 0, thu: 0 };
      for (var i = 1; i < v.length; i++) {
        var id = String(v[i][m['ID hoa don']] || '').trim(); if (!id) continue;
        var r = hdCu_(m, v[i], !daGap[id]); daGap[id] = 1;
        rows.push(r);
        cu.dong++; cu.tien += num_(r[10]); cu.tong += num_(r[13]); cu.thu += num_(r[14]);
      }
      var data = rows.length ? hdSangMisa_(rows, head, mm) : [];
      // Kiem tra khop truoc khi doi
      var c = cotMap_(head), moi = { dong: data.length, tien: 0, tong: 0, thu: 0 };
      data.forEach(function (x) {
        moi.tien += num_(x[c['Thành tiền']]); moi.tong += num_(x[c['Tong thanh toan']]); moi.thu += num_(x[c['Da thu']]);
      });
      if (cu.dong !== moi.dong || Math.round(cu.tien) !== Math.round(moi.tien) ||
          Math.round(cu.tong) !== Math.round(moi.tong) || Math.round(cu.thu) !== Math.round(moi.thu)) {
        throw new Error(name + ': LECH sau khi chuyen ' + JSON.stringify({ cu: cu, moi: moi }) + ' - DUNG, khong doi gi');
      }
      var tam = ss.insertSheet(name + '_MISA_' + Date.now());
      if (tam.getMaxColumns() < head.length) tam.insertColumnsAfter(tam.getMaxColumns(), head.length - tam.getMaxColumns());
      tam.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
      tam.getRange(1, 1, 1, MISA_HEAD.length).setBackground('#DDEBF7');               // cot MISA
      tam.getRange(1, MISA_HEAD.length + 1, 1, head.length - MISA_HEAD.length).setBackground('#FCE4D6'); // cot app
      MM_COT.forEach(function (x) { tam.getRange(1, c[x[0]] + 1).setBackground('#FFF2CC'); }); // cot ma ke toan dien
      tam.setFrozenRows(1);
      if (data.length) {
        HD_COT_CHU.forEach(function (n) { if (c[n] != null) tam.getRange(2, c[n] + 1, data.length, 1).setNumberFormat('@'); });
        tam.getRange(2, 1, data.length, head.length).setValues(data);
      }
      sh.setName(name + '_truocMISA_20261013');
      tam.setName(name);
      if (name === 'HoaDon' && rows.length) themKhoaMaMisa_(ss, mm, rows);
      bao.push(name + ': OK ' + moi.dong + ' dong, Thanh tien ' + Math.round(moi.tien) +
               ', Tong TT ' + Math.round(moi.tong) + ', Da thu ' + Math.round(moi.thu) +
               ' (ban cu luu o tab ' + name + '_truocMISA_20261013)');
    });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
    Logger.log(bao.join('\n'));
  }
  return bao;
}
