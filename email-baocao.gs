// ============================================================
// GUI BAO CAO NGAY QUA EMAIL - Kho Nuoc Ion Kiem DA05
// Dan code nay vao CUNG Apps Script project cua Sheet kho
// (project da co backend-apps-script.gs). Cac ham dat tien to bc_
// de khong trung ten voi backend.
// ------------------------------------------------------------
// CACH DUNG:
//  1. Dan code -> Luu.
//  2. Chay ham taoTabChiPhi 1 lan (tao tab ChiPhi de ghi thue/luong/ship).
//  3. Chay ham guiThuNgay 1 lan -> cho phep quyen gui mail -> kiem tra hop thu.
//  4. Chay ham caiLichGuiBaoCao 1 lan -> tu dong gui 21h moi ngay.
//     (Tat: chay tatLichBaoCao.)
// ============================================================

var EMAIL_TO = 'tnv9090@gmail.com';
var GIO_GUI = 21; // gio gui hang ngay (0-23)

function bc_ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function bc_ymd(v){ if (v instanceof Date) return Utilities.formatDate(v,'GMT+7','yyyy-MM-dd'); return String(v==null?'':v).slice(0,10); }
function bc_n(v){ var s=String(v==null?'':v).replace(/[.,\s]/g,''); var x=parseFloat(s); return isNaN(x)?0:x; }
function bc_fmt(n){ return Math.round(Number(n)||0).toLocaleString('vi-VN'); }
function bc_today(){ return Utilities.formatDate(new Date(),'GMT+7','yyyy-MM-dd'); }

function bc_rows(name){
  var sh=bc_ss().getSheetByName(name); if(!sh||sh.getLastRow()<2) return [];
  return sh.getDataRange().getValues().slice(1);
}

function bc_cat(){
  var out={}, v=bc_rows('Data');
  v.forEach(function(r){ if(!r[0])return; out[String(r[0]).trim()]={ten:String(r[1]||'').trim(),gv:Number(r[3])||0,dau:Number(r[9])||0}; });
  return out;
}

// Doc XuatKho + KhoHN: dedup LineID, phan loai Ban/Tang/ThuHoi
function bc_xuat(){
  var out=[];
  ['XuatKho','KhoHN'].forEach(function(name){
    var v=bc_rows(name), seen={};
    v.forEach(function(r){
      var lid=String(r[0]||''); if(lid){ if(seen[lid])return; seen[lid]=1; }
      if(!lid && !r[10] && !r[12]) return;
      var ten=String(r[8]||'').toUpperCase(), ghi=String(r[15]||'').toLowerCase(), tien=bc_n(r[12]);
      var isTang = ten.indexOf('[TANG]')>=0 || ((ghi.indexOf('tang')>=0||ghi.indexOf('tặng')>=0)&&tien===0);
      var lt = ten.indexOf('[THUHOI]')>=0 ? 'ThuHoi' : (isTang?'Tang':'Ban');
      out.push({ngay:bc_ymd(r[2]),ma:String(r[7]||'').trim(),sl:bc_n(r[10]),tien:tien,lt:lt,khach:String(r[5]||r[8]||''),no:bc_n(r[14])});
    });
  });
  return out;
}
function bc_nhap(){
  var out=[], v=bc_rows('NhapKho'), seen={};
  v.forEach(function(r){ var lid=String(r[0]||''); if(!lid)return; if(seen[lid])return; seen[lid]=1;
    out.push({ngay:bc_ymd(r[2]),ma:String(r[4]||'').trim(),sl:bc_n(r[7]),tien:bc_n(r[9])}); });
  return out;
}
function bc_thu(){
  var out=[], v=bc_rows('ThuTien');
  v.forEach(function(r){ if(!r[0])return; out.push({ngay:bc_ymd(r[1]),dl:String(r[2]||''),tien:bc_n(r[3])}); });
  return out;
}
function bc_chiphi(){
  // tab ChiPhi: Ngay | Khoan | So tien | Ghi chu
  var out=[], v=bc_rows('ChiPhi');
  v.forEach(function(r){ if(!r[0]&&!r[2])return; out.push({ngay:bc_ymd(r[0]),khoan:String(r[1]||''),tien:bc_n(r[2])}); });
  return out;
}

function guiBaoCaoNgay(){
  var ngay=bc_today(), thang=ngay.slice(0,7);
  var cat=bc_cat(), xuat=bc_xuat(), nhap=bc_nhap(), thu=bc_thu(), chi=bc_chiphi();

  // Hom nay
  var dtToday=0, slToday=0, phieuToday={}, nhapToday=0, chiToday=0;
  xuat.forEach(function(l){ if(l.ngay===ngay && l.lt==='Ban'){ dtToday+=l.tien; slToday+=l.sl; } });
  nhap.forEach(function(l){ if(l.ngay===ngay) nhapToday+=l.tien; });
  chi.forEach(function(c){ if(c.ngay===ngay) chiToday+=c.tien; });

  // Thang nay
  var dtThang=0, nhapThang=0, chiThang=0;
  xuat.forEach(function(l){ if(l.ngay.slice(0,7)===thang && l.lt==='Ban') dtThang+=l.tien; });
  nhap.forEach(function(l){ if(l.ngay.slice(0,7)===thang) nhapThang+=l.tien; });
  chi.forEach(function(c){ if(c.ngay.slice(0,7)===thang) chiThang+=c.tien; });

  // Ton + gia tri + cong no
  var mv={}; // ma -> {nhap,ban,tang,thuhoi}
  nhap.forEach(function(l){ var o=mv[l.ma]||(mv[l.ma]={nhap:0,ban:0,tang:0,thuhoi:0}); o.nhap+=l.sl; });
  xuat.forEach(function(l){ var o=mv[l.ma]||(mv[l.ma]={nhap:0,ban:0,tang:0,thuhoi:0}); if(l.lt==='ThuHoi')o.thuhoi+=l.sl; else if(l.lt==='Tang')o.tang+=l.sl; else o.ban+=l.sl; });
  var nhapThangSL={}; xuat.forEach(function(){}); nhap.forEach(function(l){ if(l.ngay.slice(0,7)===thang) nhapThangSL[l.ma]=(nhapThangSL[l.ma]||0)+l.sl; });

  var giaTriTon=0, tonRows='', canhBao='';
  Object.keys(cat).forEach(function(ma){
    var c=cat[ma], o=mv[ma]||{nhap:0,ban:0,tang:0,thuhoi:0};
    var ton=(c.dau||0)+o.nhap-o.ban-o.tang+o.thuhoi, gt=ton*c.gv; giaTriTon+=gt;
    tonRows+='<tr><td>'+ma+'</td><td>'+c.ten+'</td><td style="text-align:right">'+bc_fmt(ton)+'</td><td style="text-align:right">'+bc_fmt(gt)+'</td></tr>';
    var nt=nhapThangSL[ma]||0, nguong=Math.round(nt*0.2);
    if((nt>0&&ton<nguong)||ton<=0) canhBao+='<li>'+ma+' - '+c.ten+': ton '+bc_fmt(ton)+' (nguong 20% nhap thang = '+bc_fmt(nguong)+')</li>';
  });

  // Cong no
  var no={}; xuat.forEach(function(l){ if(l.no>0) no[l.khach]=(no[l.khach]||0)+l.no; });
  thu.forEach(function(t){ no[t.dl]=(no[t.dl]||0)-t.tien; });
  var tongNo=0, noRows=''; Object.keys(no).forEach(function(k){ if(no[k]>0.5){ tongNo+=no[k]; noRows+='<tr><td>'+k+'</td><td style="text-align:right">'+bc_fmt(no[k])+'</td></tr>'; } });

  var ngayVN=ngay.split('-').reverse().join('/');
  var loiToday = dtToday - nhapToday - chiToday;
  var html=''
   +'<div style="font-family:Arial,sans-serif;font-size:14px;color:#1a2430;max-width:640px">'
   +'<h2 style="color:#c8102e;margin:0 0 4px">BAO CAO NGAY '+ngayVN+'</h2>'
   +'<div style="color:#777;font-size:12px;margin-bottom:12px">Kho Nuoc Ion Kiem DA05</div>'
   +'<h3 style="margin:14px 0 6px">Hom nay</h3>'
   +'<table style="border-collapse:collapse;width:100%;font-size:13px">'
   +bc_kv('Doanh thu ban', bc_fmt(dtToday), '#0f9d58')
   +bc_kv('So luong ban', bc_fmt(slToday)+' sp', '')
   +bc_kv('Tien nhap hang (mua vao)', bc_fmt(nhapToday), '#c8102e')
   +bc_kv('Chi phi van hanh', bc_fmt(chiToday), '#c8102e')
   +bc_kv('Lai gop tam tinh (DT - nhap - chi phi)', bc_fmt(loiToday), loiToday>=0?'#0f9d58':'#c8102e')
   +'</table>'
   +'<h3 style="margin:16px 0 6px">Thang nay ('+thang.split('-').reverse().join('/')+')</h3>'
   +'<table style="border-collapse:collapse;width:100%;font-size:13px">'
   +bc_kv('Doanh thu', bc_fmt(dtThang), '')
   +bc_kv('Tien nhap hang', bc_fmt(nhapThang), '')
   +bc_kv('Chi phi van hanh', bc_fmt(chiThang), '')
   +'</table>'
   +'<h3 style="margin:16px 0 6px">Tong quan kho</h3>'
   +'<table style="border-collapse:collapse;width:100%;font-size:13px">'
   +bc_kv('Gia tri ton kho (von)', bc_fmt(giaTriTon), '')
   +bc_kv('Cong no dai ly (dang thu)', bc_fmt(tongNo), '#e08a00')
   +'</table>'
   +'<h3 style="margin:16px 0 6px">Ton kho theo san pham</h3>'
   +'<table style="border-collapse:collapse;width:100%;font-size:12.5px" border="1" cellpadding="5">'
   +'<tr style="background:#f1efe8"><th align="left">Ma</th><th align="left">Ten</th><th>Ton</th><th>Gia tri</th></tr>'+tonRows+'</table>'
   +(noRows?('<h3 style="margin:16px 0 6px">Cong no theo dai ly</h3><table style="border-collapse:collapse;width:100%;font-size:12.5px" border="1" cellpadding="5"><tr style="background:#f1efe8"><th align="left">Dai ly</th><th>Con no</th></tr>'+noRows+'</table>'):'')
   +(canhBao?('<h3 style="margin:16px 0 6px;color:#e08a00">Canh bao ton thap (&lt;20% nhap thang)</h3><ul style="color:#8a5a00">'+canhBao+'</ul>'):'')
   +'<div style="color:#999;font-size:11px;margin-top:18px">Email tu dong tu Google Apps Script. Chi phi van hanh lay tu tab ChiPhi.</div>'
   +'</div>';

  MailApp.sendEmail({ to: EMAIL_TO, subject: 'Bao cao ban hang + ton kho ngay '+ngayVN+' - DA05', htmlBody: html });
}

function bc_kv(k,v,color){
  return '<tr><td style="padding:4px 8px;color:#555">'+k+'</td><td style="padding:4px 8px;text-align:right;font-weight:bold'+(color?';color:'+color:'')+'">'+v+'</td></tr>';
}

// Gui thu ngay bay gio (test)
function guiThuNgay(){ guiBaoCaoNgay(); }

// Tao tab ChiPhi de ghi chi phi van hanh
function taoTabChiPhi(){
  var ss=bc_ss(), sh=ss.getSheetByName('ChiPhi');
  if(!sh){ sh=ss.insertSheet('ChiPhi'); sh.appendRow(['Ngay (yyyy-mm-dd)','Khoan muc','So tien','Ghi chu']); sh.setFrozenRows(1); }
}

// Cai lich tu dong gui 21h moi ngay (chay 1 lan)
function caiLichGuiBaoCao(){
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='guiBaoCaoNgay') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('guiBaoCaoNgay').timeBased().everyDays(1).atHour(GIO_GUI).create();
}
function tatLichBaoCao(){
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='guiBaoCaoNgay') ScriptApp.deleteTrigger(t); });
}
