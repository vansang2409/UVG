import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const dir = path.resolve('..');
const support = process.cwd();
const wb = Workbook.create();
const names = ['Tong_hop','Don_hang','Giao_dich_san','Chuyen_tien','Ngan_hang','So_du','Huong_dan'];
const sh = Object.fromEntries(names.map(n => [n, wb.worksheets.add(n)]));
const C = {ink:'#172B4D',navy:'#243B64',teal:'#356B70',blue:'#1565C0',green:'#23734E',muted:'#62748B',input:'#FFF4D5',calc:'#F1F4F8',red:'#B42318',pink:'#FEE4E2',rule:'#D6DFEA'};
const money = '#,##0;(#,##0);"–"';
const diffFmt = '#,##0.00;(#,##0.00);"0.00"';
const dt = s => new Date(`${s}T00:00:00.000Z`);
const first = 7, last = 106;
const q = n => `'${n}'`;
const put = (s,a,v) => sh[s].getRange(a).values=[[v]];
const fx = (s,a,f) => sh[s].getRange(a).formulas=[[f]];
function letters(n){let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;}
function base(s,endCol=15,endRow=last){
  const x=sh[s]; x.showGridLines=false;
  x.getRange(`A1:${letters(endCol)}${endRow}`).format.font={name:'Arial',size:10,color:C.ink};
  x.getRange(`A1:${letters(endCol)}${endRow}`).format.verticalAlignment='center';
  x.getRange(`A1:${letters(endCol)}${endRow}`).format.rowHeight=22;
  x.getRange(`A1:${letters(endCol)}${endRow}`).format.columnWidth=18;
  x.getRange('A2').format.font={name:'Arial',size:15,bold:true,color:C.ink};
  x.getRange(`A3:${letters(endCol)}3`).format.borders={bottom:{style:'thin',color:C.rule}};
  x.getRange('A:A').format.columnWidth=16;
  x.getRange('B:B').format.columnWidth=20;
}
function table(s,tableName,headers,inputLastCol){
  const x=sh[s], ec=letters(headers.length);
  base(s,headers.length);
  x.getRange(`A6:${ec}6`).values=[headers];
  const t=x.tables.add(`A6:${ec}${last}`,true,tableName);
  t.style='TableStyleLight9';t.showFilterButton=true;
  x.getRange(`A6:${ec}6`).format={fill:C.navy,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:58,wrapText:true,horizontalAlignment:'center',verticalAlignment:'center',borders:{insideVertical:{style:'thin',color:'#FFFFFF'}}};
  x.getRange(`A${first}:${letters(inputLastCol)}${last}`).format.fill=C.input;
  x.getRange(`A${first}:${letters(inputLastCol)}${last}`).format.font.color=C.blue;
  if(inputLastCol<headers.length)x.getRange(`${letters(inputLastCol+1)}${first}:${ec}${last}`).format.fill=C.calc;
  x.getRange(`A${first}:A${last}`).dataValidation={rule:{type:'list',values:['TikTok Shop','Shopee','Lazada']}};
  x.freezePanes.freezeRows(6);x.freezePanes.freezeColumns(2);
  return t;
}
function colform(s,c,f){fx(s,`${c}${first}`,f(first));sh[s].getRange(`${c}${first}:${c}${last}`).fillDown();}
function errCF(s,range,expression){sh[s].getRange(range).conditionalFormats.addCustom(expression,{fill:C.pink,font:{color:C.red,bold:true}});}
function rows(s,cols,data){sh[s].getRange(`A7:${letters(cols)}${6+data.length}`).values=data;}
const START="'Tong_hop'!$C$4", END="'Tong_hop'!$F$4", TOL="'Tong_hop'!$I$4";
const validPlatform = r=>`OR(A${r}="TikTok Shop",A${r}="Shopee",A${r}="Lazada")`;
const shopRegistered = r=>`COUNTIFS(T_SoDu[Khóa shop],A${r}&"|"&B${r})=1`;

// Main view: one authoritative reporting month and tolerance.
base('Tong_hop',15,51);sh.Tong_hop.tabColor=C.navy;
put('Tong_hop','A2','ĐỐI SOÁT CÔNG NỢ SÀN THEO THÁNG');
put('Tong_hop','A4','Tháng chốt');put('Tong_hop','C4',dt('2026-09-01'));
put('Tong_hop','E4','Đầu tháng sau');fx('Tong_hop','F4','=EDATE(C4,1)');
put('Tong_hop','H4','Dung sai VND');put('Tong_hop','I4',0);
sh.Tong_hop.getRange('C4').setNumberFormat('mm/yyyy');sh.Tong_hop.getRange('F4').setNumberFormat('dd/mm/yyyy');
sh.Tong_hop.getRange('C4').format.fill=C.input;sh.Tong_hop.getRange('I4').format.fill=C.input;
sh.Tong_hop.getRange('C4').format.font.color=C.blue;sh.Tong_hop.getRange('I4').format.font.color=C.blue;
put('Tong_hop','A6','DỮ LIỆU MINH HỌA tháng 09/2026. Thay dữ liệu mẫu trước khi dùng để chốt sổ.');
sh.Tong_hop.getRange('A6').format.font={name:'Arial',size:10,color:'#995D00',bold:true};
put('Tong_hop','A8','Công nợ theo sàn');

const orderH=['Sàn','Shop','Mã đơn','Ngày ghi nhận bán','Tiền hàng gồm VAT','Bù giá sàn chưa cộng','Hoàn tiền lũy kế','Phí sàn dự kiến','Ship shop chịu dự kiến','Điều chỉnh dự kiến +/-','QT trước kỳ','Hạn quyết toán','Nguồn / ghi chú','Khóa đơn','Dự kiến nhận lũy kế','QT trong kỳ','QT lũy kế','Chênh QT','Tình trạng','Kiểm tra dữ liệu','Phải thu gộp'];
table('Don_hang','T_Don',orderH,13);put('Don_hang','A2','ĐƠN HÀNG VÀ SỐ TIỀN DỰ KIẾN NHẬN');
put('Don_hang','A4','Một dòng / đơn. Cộng SKU trước khi nhập. Số tiền đã gồm VAT; bù giá chưa nằm trong tiền hàng mới nhập cột F.');
sh.Don_hang.getRange('C7:C106').setNumberFormat('@');
sh.Don_hang.getRange('D7:D106').setNumberFormat('dd/mm/yyyy');sh.Don_hang.getRange('L7:L106').setNumberFormat('dd/mm/yyyy');
sh.Don_hang.getRange('E7:K106').setNumberFormat(money);sh.Don_hang.getRange('O7:R106').setNumberFormat(money);sh.Don_hang.getRange('U7:U106').setNumberFormat(money);
sh.Don_hang.getRange('M:M').format.columnWidth=44;sh.Don_hang.getRange('N:N').format.columnWidth=38;sh.Don_hang.getRange('S:T').format.columnWidth=26;

const finH=['Sàn','Shop','Mã giao dịch','Mã đơn','Ngày giao dịch sàn','Nhóm giao dịch','Tiền hàng ghi có','Hỗ trợ ghi có','Hoàn tiền shop chịu','Phí sàn cấn trừ','Ship shop chịu','Điều chỉnh khác +/-','Nạp vào số dư','Thực ghi số dư +/-','Tính lại số dư','Chênh giao dịch','Khóa giao dịch','Khóa đơn','Mã sao kê / nguồn','Kiểm tra dữ liệu'];
table('Giao_dich_san','T_GD',finH,14);put('Giao_dich_san','A2','GIAO DỊCH TÀI CHÍNH SÀN');
put('Giao_dich_san','A4','Một dòng / giao dịch. G:K và M nhập dương; L,N có dấu. N lấy độc lập từ số thực ghi số dư của sàn.');
put('Giao_dich_san','A5','Cột S là nguồn/ghi chú được nhập tay. QT = tiền sàn ghi vào số dư, chưa phải tiền về ngân hàng.');
sh.Giao_dich_san.getRange('S7:S106').format.fill=C.input;sh.Giao_dich_san.getRange('S7:S106').format.font.color=C.blue;
sh.Giao_dich_san.getRange('F7:F106').dataValidation={rule:{type:'list',values:['Quyết toán','Số dư']}};
sh.Giao_dich_san.getRange('C7:D106').setNumberFormat('@');sh.Giao_dich_san.getRange('E7:E106').setNumberFormat('dd/mm/yyyy');sh.Giao_dich_san.getRange('G7:P106').setNumberFormat(money);
sh.Giao_dich_san.getRange('Q:R').format.columnWidth=39;sh.Giao_dich_san.getRange('S:S').format.columnWidth=38;sh.Giao_dich_san.getRange('T:T').format.columnWidth=27;

const payH=['Sàn','Shop','Mã đợt chuyển','Ngày trừ số dư','Tiền trừ số dư','Phí chuyển đã xác nhận','NH đã nhận trước kỳ','Nguồn / ghi chú','Khóa đợt chuyển','NH nhận trong kỳ','NH nhận lũy kế','Còn đang chuyển','Tình trạng','Kiểm tra dữ liệu','Dự kiến NH nhận'];
table('Chuyen_tien','T_Chuyen',payH,8);put('Chuyen_tien','A2','ĐỢT CHUYỂN TIỀN TỪ SÀN');
put('Chuyen_tien','A4','Nhập cả đợt cũ còn treo đầu kỳ. Phí cột F chỉ nhập khi đã xác nhận và chưa nằm trong Giao_dich_san.');
sh.Chuyen_tien.getRange('C7:C106').setNumberFormat('@');sh.Chuyen_tien.getRange('D7:D106').setNumberFormat('dd/mm/yyyy');sh.Chuyen_tien.getRange('E7:G106').setNumberFormat(money);sh.Chuyen_tien.getRange('J7:L106').setNumberFormat(money);sh.Chuyen_tien.getRange('O7:O106').setNumberFormat(money);
sh.Chuyen_tien.getRange('H:H').format.columnWidth=44;sh.Chuyen_tien.getRange('I:I').format.columnWidth=39;sh.Chuyen_tien.getRange('M:N').format.columnWidth=25;

const bankH=['Sàn','Shop','Mã giao dịch NH','Mã đợt chuyển','Ngày báo Có','Tiền thực nhận','Tài khoản NH','Diễn giải / nguồn','Khóa đợt chuyển','Khóa giao dịch NH','Phạm vi kỳ','Kiểm tra dữ liệu'];
table('Ngan_hang','T_NH',bankH,8);put('Ngan_hang','A2','TIỀN NGÂN HÀNG THỰC NHẬN');
put('Ngan_hang','A4','Một dòng / báo Có. Gắn mã đợt chuyển trước khi chốt. Dòng chưa gắn hoặc trùng mã sẽ được cảnh báo.');
sh.Ngan_hang.getRange('C7:D106').setNumberFormat('@');sh.Ngan_hang.getRange('E7:E106').setNumberFormat('dd/mm/yyyy');sh.Ngan_hang.getRange('F7:F106').setNumberFormat(money);sh.Ngan_hang.getRange('G7:G106').setNumberFormat('@');
sh.Ngan_hang.getRange('H:H').format.columnWidth=44;sh.Ngan_hang.getRange('I:J').format.columnWidth=40;sh.Ngan_hang.getRange('L:L').format.columnWidth=28;

const balH=['Sàn','Shop','Ngày chốt số dư','Công nợ đầu kỳ','Số dư ví đầu kỳ','Chờ QT theo BC net','Phí ước tính cộng lại','Khả dụng cuối kỳ','Giữ lại cuối kỳ','Đang chuyển theo BC','Công nợ sổ kế toán','Tổng QT theo BC sàn','Tổng rút theo BC sàn','Đủ nguồn dữ liệu','Nguồn / giải trình','Khóa shop','Công nợ theo nguồn','Kiểm tra dữ liệu','Bán gộp trong kỳ','Hoàn tiền trong kỳ','Phí trong kỳ','Điều chỉnh và nạp','NH nhận trong kỳ','Công nợ tính lại','QT thực tế trong kỳ','Số dư ví tính lại','Đang chuyển tính lại','Rút trong kỳ','Số lỗi đầu vào'];
table('So_du','T_SoDu',balH,15);put('So_du','A2','SỐ DƯ NGUỒN VÀ TÍNH CÔNG NỢ TỪNG SHOP');
put('So_du','A4','Một dòng / shop cho tháng đang chốt. D:M bắt buộc nhập số, kể cả 0. F:J là nguồn độc lập, không lấy từ kết quả tính lại.');
put('So_du','A5','F là phần đã ghi nhận bán nhưng chưa QT; G cộng lại phí ước tính chưa cấn trừ. H,I,J không trùng nhau.');
sh.So_du.getRange('C7:C106').setNumberFormat('dd/mm/yyyy');sh.So_du.getRange('D7:M106').setNumberFormat(money);sh.So_du.getRange('Q7:Q106').setNumberFormat(money);sh.So_du.getRange('S7:AB106').setNumberFormat(money);sh.So_du.getRange('AC7:AC106').setNumberFormat('0');
sh.So_du.getRange('N7:N106').dataValidation={rule:{type:'list',values:['Chưa','Đã đủ']}};
sh.So_du.getRange('O:O').format.columnWidth=50;sh.So_du.getRange('P:P').format.columnWidth=31;sh.So_du.getRange('R:R').format.columnWidth=29;

// Independent sample values, not company records.
const od=(s,b,id,date,e,f,g,h,i,j,k,due,note='Minh họa')=>[s,b,id,dt(date),e,f,g,h,i,j,k,dt(due),note];
rows('Don_hang',13,[
od('TikTok Shop','DEMO_TT','TT-OLD','2026-08-28',600000,0,0,60000,0,0,0,'2026-09-02','Đơn cũ chưa QT đầu kỳ'),
od('TikTok Shop','DEMO_TT','TT-001','2026-09-05',1000000,0,100000,100000,20000,0,0,'2026-09-10','Đã hoàn 100.000 trong tháng'),
od('TikTok Shop','DEMO_TT','TT-002','2026-09-28',500000,20000,0,50000,10000,0,0,'2026-10-05','Chưa đến hạn cuối T9'),
od('Shopee','DEMO_SP','SP-OLD','2026-08-29',400000,0,0,40000,0,0,0,'2026-09-03','Đơn cũ chưa QT đầu kỳ'),
od('Shopee','DEMO_SP','SP-001','2026-09-07',2000000,50000,0,205000,20000,0,0,'2026-09-15','Cố ý lệch ship 2.000 để kiểm tra'),
od('Shopee','DEMO_SP','SP-002','2026-09-27',700000,0,0,70000,0,0,0,'2026-10-03','Chưa đến hạn cuối T9'),
od('Lazada','DEMO_LZ','LZ-OLD','2026-08-29',350000,0,0,35000,0,0,0,'2026-09-02','Đơn cũ chưa QT đầu kỳ'),
od('Lazada','DEMO_LZ','LZ-001','2026-09-04',1500000,0,200000,150000,30000,20000,0,'2026-09-12','Hoàn tiền 200.000, hoàn phí 20.000'),
od('Lazada','DEMO_LZ','LZ-002','2026-09-28',800000,0,0,80000,0,0,0,'2026-10-06','Chưa đến hạn cuối T9')
]);
const gd=(s,b,id,ord,date,group,g,h,i,j,k,l,m,n,src='Minh họa')=>[s,b,id,ord,dt(date),group,g,h,i,j,k,l,m,n,null,null,null,null,src];
rows('Giao_dich_san',19,[
gd('TikTok Shop','DEMO_TT','TT-F01','TT-OLD','2026-09-02','Quyết toán',600000,0,0,60000,0,0,0,540000),
gd('TikTok Shop','DEMO_TT','TT-F02','TT-001','2026-09-10','Quyết toán',1000000,0,0,100000,20000,0,0,880000),
gd('TikTok Shop','DEMO_TT','TT-F03','','2026-09-15','Số dư',0,0,0,30000,0,0,0,-30000,'Trừ số dư cho quảng cáo'),
gd('TikTok Shop','DEMO_TT','TT-F04','TT-001','2026-09-20','Quyết toán',0,0,100000,0,0,0,0,-100000),
gd('TikTok Shop','DEMO_TT','TT-F05','','2026-09-18','Số dư',0,0,0,0,0,0,50000,50000,'Công ty nạp thêm, không phải doanh thu'),
gd('TikTok Shop','DEMO_TT','TT-F06','TT-002','2026-10-05','Quyết toán',500000,20000,0,50000,10000,0,0,460000,'Ngoài T9, không cộng vào T9'),
gd('Shopee','DEMO_SP','SP-F01','SP-OLD','2026-09-03','Quyết toán',400000,0,0,40000,0,0,0,360000),
gd('Shopee','DEMO_SP','SP-F02','SP-001','2026-09-15','Quyết toán',2000000,50000,0,205000,22000,0,0,1823000,'Ship thực tế cao hơn dự kiến 2.000'),
gd('Shopee','DEMO_SP','SP-F03','','2026-09-20','Số dư',0,0,0,0,0,25000,0,25000,'Bồi thường ngoài đơn'),
gd('Shopee','DEMO_SP','SP-F04','SP-002','2026-10-03','Quyết toán',700000,0,0,70000,0,0,0,630000),
gd('Lazada','DEMO_LZ','LZ-F01','LZ-OLD','2026-09-02','Quyết toán',350000,0,0,35000,0,0,0,315000),
gd('Lazada','DEMO_LZ','LZ-F02','LZ-001','2026-09-12','Quyết toán',1500000,0,0,150000,30000,0,0,1320000),
gd('Lazada','DEMO_LZ','LZ-F03','LZ-001','2026-09-22','Quyết toán',0,0,200000,0,0,20000,0,-180000,'Hoàn tiền và hoàn phí'),
gd('Lazada','DEMO_LZ','LZ-F04','LZ-002','2026-10-06','Quyết toán',800000,0,0,80000,0,0,0,720000)
]);
rows('Chuyen_tien',8,[
['TikTok Shop','DEMO_TT','TT-P0',dt('2026-08-31'),100000,0,0,'Đang chuyển đầu tháng'],
['TikTok Shop','DEMO_TT','TT-P1',dt('2026-09-15'),1000000,10000,0,'Phí chuyển 10.000 chưa nằm trong GD sàn'],
['TikTok Shop','DEMO_TT','TT-P2',dt('2026-09-30'),400000,0,0,'Ngân hàng nhận ngày 02/10'],
['Shopee','DEMO_SP','SP-P0',dt('2026-08-31'),100000,0,0,'Đang chuyển đầu tháng'],
['Shopee','DEMO_SP','SP-P1',dt('2026-09-20'),1800000,0,0,'Đã nhận đủ'],
['Shopee','DEMO_SP','SP-P2',dt('2026-09-30'),500000,0,0,'Ngân hàng nhận ngày 02/10'],
['Lazada','DEMO_LZ','LZ-P0',dt('2026-08-31'),100000,0,0,'Đang chuyển đầu tháng'],
['Lazada','DEMO_LZ','LZ-P1',dt('2026-09-15'),1000000,0,0,'Cố ý thiếu báo Có 1.000'],
['Lazada','DEMO_LZ','LZ-P2',dt('2026-09-30'),300000,0,0,'Ngân hàng nhận ngày 02/10']
]);
rows('Ngan_hang',8,[
['TikTok Shop','DEMO_TT','NH-TT01','TT-P0',dt('2026-09-02'),100000,'DEMO112','Thu khoản đang chuyển đầu kỳ'],
['TikTok Shop','DEMO_TT','NH-TT02','TT-P1',dt('2026-09-16'),990000,'DEMO112','Thực nhận sau phí chuyển'],
['TikTok Shop','DEMO_TT','NH-TT03','TT-P2',dt('2026-10-02'),400000,'DEMO112','Ngoài T9'],
['Shopee','DEMO_SP','NH-SP01','SP-P0',dt('2026-09-01'),100000,'DEMO112','Thu đầu kỳ'],
['Shopee','DEMO_SP','NH-SP02','SP-P1',dt('2026-09-21'),1800000,'DEMO112','Thu đủ'],
['Shopee','DEMO_SP','NH-SP03','SP-P2',dt('2026-10-02'),500000,'DEMO112','Ngoài T9'],
['Lazada','DEMO_LZ','NH-LZ01','LZ-P0',dt('2026-09-01'),100000,'DEMO112','Thu đầu kỳ'],
['Lazada','DEMO_LZ','NH-LZ02','LZ-P1',dt('2026-09-16'),999000,'DEMO112','Thiếu 1.000 chưa có giải trình'],
['Lazada','DEMO_LZ','NH-LZ03','LZ-P2',dt('2026-10-02'),300000,'DEMO112','Ngoài T9']
]);
rows('So_du',15,[
['TikTok Shop','DEMO_TT',dt('2026-09-30'),1000000,300000,460000,60000,200000,40000,400000,1160000,1320000,1400000,'Đã đủ','Số giả lập; cả ba lớp đối chiếu khớp'],
['Shopee','DEMO_SP',dt('2026-09-30'),800000,300000,630000,70000,208000,0,500000,1408000,2183000,2300000,'Đã đủ','Số giả lập; đơn SP-001 lệch phí 2.000'],
['Lazada','DEMO_LZ',dt('2026-09-30'),600000,150000,720000,80000,305000,0,300000,1406000,1455000,1300000,'Đã đủ','Số giả lập; đợt LZ-P1 thiếu tiền NH 1.000']
]);

// Row keys and validation stay with their owning inputs.
colform('So_du','P',r=>`=IF(AND(A${r}="",B${r}=""),"",A${r}&"|"&B${r})`);
colform('So_du','R',r=>`=IF(COUNTA(A${r}:O${r})=0,"",IF(OR(NOT(${validPlatform(r)}),B${r}="",NOT(ISNUMBER(C${r})),COUNT(D${r}:M${r})<>10),"Thiếu/sai số dư",IF(C${r}<>${END}-1,"Sai ngày chốt",IF(COUNTIFS(T_SoDu[Khóa shop],P${r})<>1,"Trùng shop",IF(AND(N${r}<>"Đã đủ",N${r}<>"Chưa"),"Thiếu xác nhận nguồn","")))))`);
colform('So_du','Q',r=>`=IF(P${r}="","",IF(COUNT(F${r}:J${r})=5,SUM(F${r}:J${r}),"n.a."))`);
colform('Don_hang','N',r=>`=IF(C${r}="","",A${r}&"|"&B${r}&"|"&C${r})`);
colform('Don_hang','T',r=>`=IF(COUNTA(A${r}:M${r})=0,"",IF(OR(NOT(${validPlatform(r)}),B${r}="",C${r}="",NOT(ISNUMBER(D${r})),COUNT(E${r}:K${r})<>7,NOT(ISNUMBER(L${r}))),"Thiếu/sai dữ liệu",IF(NOT(${shopRegistered(r)}),"Shop chưa khai báo",IF(COUNTIFS(T_Don[Khóa đơn],N${r})<>1,"Trùng mã đơn",""))))`);
colform('Giao_dich_san','Q',r=>`=IF(C${r}="","",A${r}&"|"&B${r}&"|"&C${r})`);
colform('Giao_dich_san','R',r=>`=IF(D${r}="","",A${r}&"|"&B${r}&"|"&D${r})`);
colform('Giao_dich_san','T',r=>`=IF(COUNTA(A${r}:N${r})=0,"",IF(OR(NOT(${validPlatform(r)}),B${r}="",C${r}="",NOT(ISNUMBER(E${r})),AND(F${r}<>"Quyết toán",F${r}<>"Số dư"),COUNT(G${r}:N${r})<>8),"Thiếu/sai dữ liệu",IF(NOT(${shopRegistered(r)}),"Shop chưa khai báo",IF(COUNTIFS(T_GD[Khóa giao dịch],Q${r})<>1,"Trùng giao dịch",IF(AND(D${r}<>"",COUNTIFS(T_Don[Khóa đơn],R${r})<>1),"Không có đơn duy nhất","")))))`);
colform('Giao_dich_san','O',r=>`=IF(C${r}="","",IF(COUNT(G${r}:M${r})=7,SUM(G${r}:H${r})-SUM(I${r}:K${r})+SUM(L${r}:M${r}),"n.a."))`);
colform('Giao_dich_san','P',r=>`=IF(C${r}="","",IF(COUNT(N${r}:O${r})=2,N${r}-O${r},"n.a."))`);
colform('Chuyen_tien','I',r=>`=IF(C${r}="","",A${r}&"|"&B${r}&"|"&C${r})`);
colform('Chuyen_tien','N',r=>`=IF(COUNTA(A${r}:H${r})=0,"",IF(OR(NOT(${validPlatform(r)}),B${r}="",C${r}="",NOT(ISNUMBER(D${r})),COUNT(E${r}:G${r})<>3,E${r}<0,F${r}<0,F${r}>E${r},G${r}<0),"Thiếu/sai dữ liệu",IF(NOT(${shopRegistered(r)}),"Shop chưa khai báo",IF(COUNTIFS(T_Chuyen[Khóa đợt chuyển],I${r})<>1,"Trùng đợt chuyển",""))))`);
colform('Ngan_hang','I',r=>`=IF(D${r}="","",A${r}&"|"&B${r}&"|"&D${r})`);
colform('Ngan_hang','J',r=>`=IF(C${r}="","",G${r}&"|"&C${r})`);
colform('Ngan_hang','L',r=>`=IF(COUNTA(A${r}:H${r})=0,"",IF(OR(NOT(${validPlatform(r)}),B${r}="",C${r}="",D${r}="",NOT(ISNUMBER(E${r})),NOT(ISNUMBER(F${r})),F${r}<0,G${r}=""),"Thiếu/sai dữ liệu",IF(NOT(${shopRegistered(r)}),"Shop chưa khai báo",IF(COUNTIFS(T_NH[Khóa giao dịch NH],J${r})<>1,"Trùng giao dịch NH",IF(COUNTIFS(T_Chuyen[Khóa đợt chuyển],I${r})<>1,"Chưa ghép đợt chuyển","")))))`);
colform('Ngan_hang','K',r=>`=IF(C${r}="","",IF(AND(E${r}>=${START},E${r}<${END}),"Trong kỳ","Ngoài kỳ"))`);

// Business calculations; transaction source cash is independent of reconstructed components.
colform('Don_hang','U',r=>`=IF(C${r}="","",IF(COUNT(E${r}:F${r})=2,SUM(E${r}:F${r}),"n.a."))`);
colform('Don_hang','O',r=>`=IF(C${r}="","",IF(COUNT(E${r}:K${r})=7,U${r}-SUM(G${r}:I${r})+J${r},"n.a."))`);
colform('Don_hang','P',r=>`=IF(C${r}="","",SUMIFS(T_GD[Thực ghi số dư +/-],T_GD[Khóa đơn],N${r},T_GD[Nhóm giao dịch],"Quyết toán",T_GD[Ngày giao dịch sàn],">="&${START},T_GD[Ngày giao dịch sàn],"<"&${END}))`);
colform('Don_hang','Q',r=>`=IF(C${r}="","",IF(ISNUMBER(K${r}),K${r}+P${r},"n.a."))`);
colform('Don_hang','R',r=>`=IF(C${r}="","",IF(AND(ISNUMBER(O${r}),ISNUMBER(Q${r})),Q${r}-O${r},"n.a."))`);
colform('Don_hang','S',r=>`=IF(C${r}="","",IF(T${r}<>"","Sửa dữ liệu",IF(D${r}>=${END},"Ngoài kỳ",IF(ABS(R${r})<=${TOL},"Khớp",IF(R${r}>${TOL},"QT vượt dự kiến",IF(L${r}<${END},"Quá hạn / thiếu QT",IF(Q${r}=0,"Chờ QT","Chờ QT một phần")))))))`);
colform('Chuyen_tien','O',r=>`=IF(C${r}="","",IF(COUNT(E${r}:F${r})=2,E${r}-F${r},"n.a."))`);
colform('Chuyen_tien','J',r=>`=IF(C${r}="","",SUMIFS(T_NH[Tiền thực nhận],T_NH[Khóa đợt chuyển],I${r},T_NH[Ngày báo Có],">="&${START},T_NH[Ngày báo Có],"<"&${END}))`);
colform('Chuyen_tien','K',r=>`=IF(C${r}="","",IF(ISNUMBER(G${r}),G${r}+J${r},"n.a."))`);
colform('Chuyen_tien','L',r=>`=IF(C${r}="","",IF(D${r}>=${END},0,IF(AND(ISNUMBER(O${r}),ISNUMBER(K${r})),O${r}-K${r},"n.a.")))`);
colform('Chuyen_tien','M',r=>`=IF(C${r}="","",IF(N${r}<>"","Sửa dữ liệu",IF(D${r}>=${END},"Ngoài kỳ",IF(ABS(L${r})<=${TOL},"Đã nhận đủ",IF(L${r}<0,"Nhận vượt",IF(K${r}>0,"Nhận thiếu","Đang chuyển"))))))`);

const sumShop=(table,amount, date, r, extra='')=>`SUMIFS(${table}[${amount}],${table}[Sàn],A${r},${table}[Shop],B${r}${date?`,${table}[${date}],">="&${START},${table}[${date}],"<"&${END}`:''}${extra})`;
colform('So_du','AC',r=>`=IF(P${r}="","",IF(R${r}="",0,1)+COUNTIFS(T_Don[Sàn],A${r},T_Don[Shop],B${r},T_Don[Kiểm tra dữ liệu],"?*")+COUNTIFS(T_GD[Sàn],A${r},T_GD[Shop],B${r},T_GD[Kiểm tra dữ liệu],"?*")+COUNTIFS(T_Chuyen[Sàn],A${r},T_Chuyen[Shop],B${r},T_Chuyen[Kiểm tra dữ liệu],"?*")+COUNTIFS(T_NH[Sàn],A${r},T_NH[Shop],B${r},T_NH[Kiểm tra dữ liệu],"?*"))`);
const guarded=(r,body)=>`=IF(P${r}="","",IF(AC${r}>0,"n.a.",${body}))`;
colform('So_du','S',r=>guarded(r,sumShop('T_Don','Phải thu gộp','Ngày ghi nhận bán',r)));
colform('So_du','T',r=>guarded(r,sumShop('T_GD','Hoàn tiền shop chịu','Ngày giao dịch sàn',r)));
colform('So_du','U',r=>guarded(r,`${sumShop('T_GD','Phí sàn cấn trừ','Ngày giao dịch sàn',r)}+${sumShop('T_GD','Ship shop chịu','Ngày giao dịch sàn',r)}+${sumShop('T_Chuyen','Phí chuyển đã xác nhận','Ngày trừ số dư',r)}`));
colform('So_du','V',r=>guarded(r,`${sumShop('T_GD','Điều chỉnh khác +/-','Ngày giao dịch sàn',r)}+${sumShop('T_GD','Nạp vào số dư','Ngày giao dịch sàn',r)}`));
colform('So_du','W',r=>guarded(r,sumShop('T_NH','Tiền thực nhận','Ngày báo Có',r)));
colform('So_du','X',r=>guarded(r,`D${r}+S${r}-T${r}-U${r}+V${r}-W${r}`));
colform('So_du','Y',r=>guarded(r,sumShop('T_GD','Thực ghi số dư +/-','Ngày giao dịch sàn',r,',T_GD[Nhóm giao dịch],"Quyết toán"')));
colform('So_du','Z',r=>guarded(r,`E${r}+${sumShop('T_GD','Thực ghi số dư +/-','Ngày giao dịch sàn',r)}-${sumShop('T_Chuyen','Tiền trừ số dư','Ngày trừ số dư',r)}`));
colform('So_du','AA',r=>guarded(r,sumShop('T_Chuyen','Còn đang chuyển','',r)));
colform('So_du','AB',r=>guarded(r,sumShop('T_Chuyen','Tiền trừ số dư','Ngày trừ số dư',r)));

// Dashboard: business results first; independent checks below do not feed the business calculations.
const mh=['Sàn','Đầu kỳ','Bán gộp','Hoàn tiền','Phí cấn trừ','Điều chỉnh / nạp','NH thực nhận','Cuối kỳ tính lại','Cuối kỳ nguồn','Cuối kỳ sổ KT','Lệch nguồn','Lệch sổ KT','Nguồn dữ liệu'];
sh.Tong_hop.getRange('A9:M9').values=[mh];
sh.Tong_hop.getRange('A10:A12').values=[['TikTok Shop'],['Shopee'],['Lazada']];
const mc=['Công nợ đầu kỳ','Bán gộp trong kỳ','Hoàn tiền trong kỳ','Phí trong kỳ','Điều chỉnh và nạp','NH nhận trong kỳ','Công nợ tính lại','Công nợ theo nguồn','Công nợ sổ kế toán'];
function agg(r,c){return `SUMIFS(T_SoDu[${c}],T_SoDu[Sàn],$A${r})`;}
function agguard(r,c){return `=IF(COUNTIFS(T_SoDu[Sàn],$A${r})=0,"n.a.",IF(COUNTIFS(T_SoDu[Sàn],$A${r},T_SoDu[Số lỗi đầu vào],">0")>0,"n.a.",${agg(r,c)}))`;}
for(let r=10;r<=12;r++){
  for(let i=0;i<mc.length;i++)fx('Tong_hop',`${letters(i+2)}${r}`,agguard(r,mc[i]));
  fx('Tong_hop',`K${r}`,`=IF(COUNT(H${r}:I${r})=2,H${r}-I${r},"n.a.")`);
  fx('Tong_hop',`L${r}`,`=IF(AND(ISNUMBER(H${r}),ISNUMBER(J${r})),H${r}-J${r},"n.a.")`);
  fx('Tong_hop',`M${r}`,`=IF(COUNTIFS(T_SoDu[Sàn],A${r})=0,"Chưa có shop",IF(COUNTIFS(T_SoDu[Sàn],A${r},T_SoDu[Số lỗi đầu vào],">0")>0,"Sửa dữ liệu",IF(COUNTIFS(T_SoDu[Sàn],A${r},T_SoDu[Đủ nguồn dữ liệu],"<>Đã đủ")>0,"Chưa đủ nguồn","Đã xác nhận đủ")))`);
}
put('Tong_hop','A13','TỔNG');
for(let c=2;c<=12;c++)fx('Tong_hop',`${letters(c)}13`,`=IF(COUNT(${letters(c)}10:${letters(c)}12)=3,SUM(${letters(c)}10:${letters(c)}12),"n.a.")`);
put('Tong_hop','A15','Đơn vị VND. Công nợ tính lại = đầu kỳ + bán gộp − hoàn tiền − phí + điều chỉnh/nạp − NH nhận.');
put('Tong_hop','A16','Phí dự kiến ở Đơn hàng chỉ để kiểm tra quyết toán; chỉ phí đã cấn trừ mới giảm công nợ ở bảng trên.');
put('Tong_hop','A18','Đối chiếu nguồn độc lập');
const ch=['Sàn','QT chi tiết','QT theo BC','Lệch QT','Ví tính lại','Ví theo BC','Lệch ví','Đang chuyển tính lại','Đang chuyển BC','Lệch đang chuyển','Rút chi tiết','Rút theo BC','Lệch rút'];
sh.Tong_hop.getRange('A19:M19').values=[ch];
for(let r=20;r<=22;r++){
  fx('Tong_hop',`A${r}`,`=A${r-10}`);
  fx('Tong_hop',`B${r}`,agguard(r,'QT thực tế trong kỳ'));fx('Tong_hop',`C${r}`,agguard(r,'Tổng QT theo BC sàn'));
  fx('Tong_hop',`D${r}`,`=IF(COUNT(B${r}:C${r})=2,B${r}-C${r},"n.a.")`);
  fx('Tong_hop',`E${r}`,agguard(r,'Số dư ví tính lại'));
  fx('Tong_hop',`F${r}`,`=IF(COUNTIFS(T_SoDu[Sàn],$A${r})=0,"n.a.",IF(COUNTIFS(T_SoDu[Sàn],$A${r},T_SoDu[Số lỗi đầu vào],">0")>0,"n.a.",${agg(r,'Khả dụng cuối kỳ')}+${agg(r,'Giữ lại cuối kỳ')}))`);
  fx('Tong_hop',`G${r}`,`=IF(COUNT(E${r}:F${r})=2,E${r}-F${r},"n.a.")`);
  fx('Tong_hop',`H${r}`,agguard(r,'Đang chuyển tính lại'));fx('Tong_hop',`I${r}`,agguard(r,'Đang chuyển theo BC'));
  fx('Tong_hop',`J${r}`,`=IF(COUNT(H${r}:I${r})=2,H${r}-I${r},"n.a.")`);
  fx('Tong_hop',`K${r}`,agguard(r,'Rút trong kỳ'));fx('Tong_hop',`L${r}`,agguard(r,'Tổng rút theo BC sàn'));
  fx('Tong_hop',`M${r}`,`=IF(COUNT(K${r}:L${r})=2,K${r}-L${r},"n.a.")`);
}
put('Tong_hop','A25','Các khoản cần kiểm tra');
put('Tong_hop','A26','Nội dung');put('Tong_hop','B26','Số dòng');put('Tong_hop','C26','Nơi xem');put('Tong_hop','E26','Cách xử lý');
const issues=[
['Lỗi dữ liệu số dư','=COUNTIFS(T_SoDu[Kiểm tra dữ liệu],"?*")','So_du / cột R','Nhập số 0 khi không phát sinh; đúng ngày chốt; shop không trùng.'],
['Lỗi dữ liệu đơn hàng','=COUNTIFS(T_Don[Kiểm tra dữ liệu],"?*")','Don_hang / cột T','Bổ sung mã/ngày/giá trị; gộp các SKU thành một đơn.'],
['Lỗi dữ liệu giao dịch','=COUNTIFS(T_GD[Kiểm tra dữ liệu],"?*")','Giao_dich_san / cột T','Kiểm tra mã trùng, shop, đơn gốc và số tiền.'],
['Lỗi đợt chuyển tiền','=COUNTIFS(T_Chuyen[Kiểm tra dữ liệu],"?*")','Chuyen_tien / cột N','Kiểm tra mã đợt, ngày trừ số dư và khoản nhận trước kỳ.'],
['Lỗi ghép ngân hàng','=COUNTIFS(T_NH[Kiểm tra dữ liệu],"?*")','Ngan_hang / cột L','Gắn đúng mã đợt chuyển; không ghi trùng báo Có.'],
['Giao dịch tài chính lệch',`=COUNTIFS(T_GD[Chênh giao dịch],">"&${TOL},T_GD[Ngày giao dịch sàn],">="&${START},T_GD[Ngày giao dịch sàn],"<"&${END})+COUNTIFS(T_GD[Chênh giao dịch],"<"&-${TOL},T_GD[Ngày giao dịch sàn],">="&${START},T_GD[Ngày giao dịch sàn],"<"&${END})`,'Giao_dich_san / cột P','So thực ghi số dư với tổng cấu phần.'],
['Đơn quá hạn / thiếu QT','=COUNTIFS(T_Don[Tình trạng],"Quá hạn / thiếu QT")','Don_hang / cột S','Kiểm tra phí, hoàn tiền, hạn QT và làm việc với sàn.'],
['Đơn QT vượt dự kiến','=COUNTIFS(T_Don[Tình trạng],"QT vượt dự kiến")','Don_hang / cột S','Kiểm tra hỗ trợ, hoàn phí, giao dịch trùng và số dự kiến.'],
['Đợt chuyển nhận thiếu','=COUNTIFS(T_Chuyen[Tình trạng],"Nhận thiếu")','Chuyen_tien / cột M','Đối chiếu báo Có và phí chuyển đã được xác nhận.'],
['Đợt chuyển nhận vượt','=COUNTIFS(T_Chuyen[Tình trạng],"Nhận vượt")','Chuyen_tien / cột M','Kiểm tra ghép mã và giao dịch ngân hàng trùng.'],
['Shop chưa xác nhận đủ nguồn','=COUNTIFS(T_SoDu[Khóa shop],"?*",T_SoDu[Đủ nguồn dữ liệu],"<>Đã đủ")','So_du / cột N','Chỉ xác nhận sau khi đã kiểm tra đủ phạm vi dữ liệu.']
];
issues.forEach((a,i)=>{let r=27+i;put('Tong_hop',`A${r}`,a[0]);fx('Tong_hop',`B${r}`,a[1]);put('Tong_hop',`C${r}`,a[2]);put('Tong_hop',`E${r}`,a[3]);});
for(let r=26;r<=37;r++){
  sh.Tong_hop.getRange(`C${r}:D${r}`).merge();
  sh.Tong_hop.getRange(`E${r}:M${r}`).merge();
}
put('Tong_hop','A40','Đọc kết quả mẫu');
put('Tong_hop','A41','TikTok: khớp. Shopee: SP-001 thiếu QT 2.000 do ship. Lazada: LZ-P1 thiếu NH 1.000.');
put('Tong_hop','A43','Dùng thực tế: thay ô vàng, giữ ô xám có công thức. Xem Huong_dan trước khi dán file sàn.');
put('Tong_hop','A44','Mẫu không tự nhập file gốc. Cần ánh xạ cột và cộng gộp theo đơn/giao dịch trước khi dán.');
for(const r of [9,19,26])sh.Tong_hop.getRange(`A${r}:M${r}`).format={fill:C.navy,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:r===26?30:50,wrapText:true,horizontalAlignment:'center',verticalAlignment:'center'};
sh.Tong_hop.getRange('B10:L13').setNumberFormat(money);sh.Tong_hop.getRange('B20:M22').setNumberFormat(money);
for(const r of ['K10:L13','D20:D22','G20:G22','J20:J22','M20:M22'])sh.Tong_hop.getRange(r).setNumberFormat(diffFmt);
sh.Tong_hop.getRange('A13:M13').format={fill:'#E8EEF6',font:{name:'Arial',size:10,bold:true,color:C.ink},rowHeight:28,borders:{top:{style:'thin',color:C.navy}}};
sh.Tong_hop.getRange('A10:M12').format.rowHeight=30;sh.Tong_hop.getRange('A20:M22').format.rowHeight=30;
sh.Tong_hop.getRange('A:A').format.columnWidth=31;sh.Tong_hop.getRange('B:L').format.columnWidth=17;sh.Tong_hop.getRange('M:M').format.columnWidth=25;
sh.Tong_hop.getRange('A27:M37').format.rowHeight=32;
sh.Tong_hop.getRange('A27:A37').format.wrapText=true;
sh.Tong_hop.getRange('E27:M37').format.wrapText=true;
for(const r of ['K10:L13','D20:D22','G20:G22','J20:J22','M20:M22']){const a=r.split(':')[0];errCF('Tong_hop',r,`AND(ISNUMBER(${a}),ABS(${a})>$I$4)`);}
errCF('Tong_hop','B27:B37','B27>0');
errCF('Tong_hop','M10:M12','M10<>"Đã xác nhận đủ"');
for(const s of ['Don_hang','Giao_dich_san','Chuyen_tien','Ngan_hang','So_du']){
  const lastCol={Don_hang:'T',Giao_dich_san:'T',Chuyen_tien:'N',Ngan_hang:'L',So_du:'R'}[s];
  errCF(s,`${lastCol}7:${lastCol}${last}`,`${lastCol}7<>""`);
}
errCF('Don_hang','R7:S106',`AND(ISNUMBER($R7),ABS($R7)>${TOL},OR($S7="Quá hạn / thiếu QT",$S7="QT vượt dự kiến"))`);
errCF('Giao_dich_san','P7:P106',`AND(ISNUMBER(P7),ABS(P7)>${TOL})`);
errCF('Chuyen_tien','L7:M106','OR($M7="Nhận thiếu",$M7="Nhận vượt")');

// Longer instructions need a separate final sheet for the recurring join/cutoff workflow.
base('Huong_dan',5,55);sh.Huong_dan.tabColor='#8293A6';
sh.Huong_dan.getRange('A:A').format.columnWidth=27;sh.Huong_dan.getRange('B:B').format.columnWidth=105;sh.Huong_dan.getRange('C:C').format.columnWidth=4;sh.Huong_dan.getRange('D:D').format.columnWidth=30;sh.Huong_dan.getRange('E:E').format.columnWidth=80;
put('Huong_dan','A2','CÁCH DÙNG MẪU ĐỐI SOÁT THÁNG');
const guide=[
['Phạm vi','Một file cho một kỳ chốt. Có thể nhập nhiều shop thuộc TikTok Shop, Shopee, Lazada. Đồng tiền VND.'],
['Dữ liệu đang có','Toàn bộ số liệu DEMO là giả lập, không phải số liệu UVGREEN. Mẫu cố ý để 2 chênh lệch chưa xử lý.'],
['Màu ô','Vàng: nhập hoặc dán giá trị. Xám: công thức. Ô đỏ: lỗi/chênh cần kiểm tra. Không dán đè công thức.'],
['Bước 1','Lưu bản riêng theo tháng. Đặt Tong_hop!C4 là ngày 01 của tháng chốt; I4 là dung sai theo chứng cứ làm tròn, mặc định 0.'],
['Bước 2','Xóa dữ liệu DEMO ở các cột nhập liệu của 5 bảng. Giao_dich_san còn có cột S nhập nguồn; không xóa cột công thức.'],
['Bước 3','Khai báo So_du cho từng shop. Chụp/xuất số dư đúng cuối kỳ; nhập công nợ đầu kỳ đã đối chiếu kỳ trước.'],
['Bước 4','Chuẩn hóa các file đơn hàng: một dòng / mã đơn. Gộp SKU trước. Mã đơn phải là Text; không làm tròn mã dài.'],
['Bước 5','Chuẩn hóa giao dịch tài chính: một dòng / mã giao dịch duy nhất. Một đơn có thể có nhiều dòng và nhiều ngày.'],
['Bước 6','Dán đợt rút tiền và báo Có ngân hàng. Giữ mã đợt chuyển để nối nhiều đơn với một đợt và nhiều báo Có.'],
['Bước 7','Kiểm tra Tong_hop: số công nợ, các chênh lệch độc lập và danh sách vấn đề. Lọc cột trạng thái ở bảng chi tiết.'],
['Bước 8','Sau khi đủ phạm vi, đánh dấu So_du cột N = Đã đủ. Không xác nhận chỉ vì các ô chênh tổng đang bằng 0.'],
['Sang tháng sau','Tạo bản sao file; cập nhật toàn bộ ảnh chụp số dư và dữ liệu của tháng mới, không chỉ thay ô tháng.'],
['QT trước kỳ','Don_hang cột K và Chuyen_tien cột G là lũy kế TRƯỚC tháng đang chốt. Không nhập lại vào giao dịch của tháng.'],
['Phạm vi đơn cũ','Bao gồm tất cả đơn còn treo đầu kỳ và đơn cũ có hoàn tiền/điều chỉnh trong tháng; loại đơn cũ đã đóng không còn biến động.'],
['Phạm vi giao dịch','Lấy theo ngày ghi số dư/giao dịch tài chính, không theo ngày đặt đơn. Có thể giữ dữ liệu tháng sau để kiểm tra thu tiếp.'],
['Ngày bán hàng','Nhập ngày đủ điều kiện ghi nhận theo hồ sơ kế toán. Mẫu không tự suy ra ngày này từ trạng thái Đã hoàn thành của sàn.'],
['Tiền hàng và VAT','Don_hang E là sau giảm giá do shop chịu, gồm VAT. F chỉ là bù giá chưa nằm trong E. Doanh thu 511 phải tách VAT khi hạch toán.'],
['Phí dự kiến','Don_hang H:I là lũy kế dự kiến đến cuối kỳ, dùng để so với tiền QT. Không trực tiếp giảm công nợ tính lại.'],
['Hoàn tiền','Don_hang G là lũy kế để tính dự kiến; Giao_dich_san I là khoản thực tế shop chịu phát sinh trong kỳ, dùng giảm công nợ.'],
['Quyết toán và số dư','Giao_dich_san nhóm Quyết toán gắn với khoản sàn QT; nhóm Số dư là biến động riêng như nạp tiền, quảng cáo, bồi thường.'],
['Nạp thêm tiền','Nạp từ công ty không phải doanh thu. Mẫu cộng vào tài sản còn tại sàn; cần theo dõi đúng tài khoản trên sổ để đối chiếu cùng cơ sở.'],
['Phí chuyển tiền','Chuyen_tien F chỉ là phí trừ thẳng từ khoản rút, đã xác nhận và chưa ghi tại Giao_dich_san. Không ghi phí hai lần.'],
['Số dư độc lập','So_du F:J lấy từ báo cáo/số dư nguồn, không sao chép số tính lại. Khả dụng, giữ lại, đang chuyển phải loại trừ trùng.'],
['Net và gross','So_du F là đơn đã ghi nhận nhưng chưa QT, có thể net phí ước tính. G chỉ cộng lại phần phí dự kiến chưa thực sự cấn trừ.'],
['Sổ kế toán','K là tổng tài sản/công nợ tương ứng với cùng phạm vi, gồm tiền đang chuyển nếu sổ đã tách 113. Không đối chiếu một TK thiếu phạm vi.'],
['Tài khoản phí phải trả','Phí đã trích trước nhưng chưa cấn trừ không tự giảm khoản phải thu. Đối chiếu khoản phải trả riêng, tránh trừ hai lần.'],
['Thời điểm khóa sổ','Lọc thời gian theo >= đầu tháng và < đầu tháng sau; giao dịch báo Có tháng sau chưa làm giảm công nợ tháng này.'],
['Dòng vượt 100','Mỗi bảng có sẵn 100 dòng. Thêm dòng TRONG Excel Table, mở rộng bảng đúng toàn bộ dữ liệu, sao chép cột công thức từ dòng trên.'],
['Sau khi thêm dòng','Công thức SUMIFS dùng tên bảng nên theo phạm vi bảng mở rộng. Kiểm tra công thức, định dạng và validation ở dòng mới trước khi chốt.'],
['Dữ liệu trống','Các số tiền bắt buộc phải nhập số, kể cả 0. Mẫu cảnh báo dữ liệu thiếu; n.a. là chưa tính được, không phải không phát sinh.'],
['Giới hạn tự động','Không có macro, kết nối API hay tự nhận diện cột file sàn. Đầu vào là dữ liệu đã chuẩn hóa; giữ file gốc để truy vết.'],
['Các khoản chưa hỗ trợ','Trường hợp sàn ứng trước, cho vay, thu hộ ngoài mẫu hoặc nhiều đồng tiền cần bổ sung lịch đối chiếu riêng trước khi dùng kết quả.'],
['Báo Có gộp nhiều đợt','Phân bổ báo Có theo đợt, đặt mã dòng NH riêng; lưu mã báo Có gốc tại H. Kiểm tra tổng phân bổ bằng sao kê, không nhập toàn bộ tiền cho từng đợt.'],
['Mã giao dịch tài chính','Nếu sàn không có mã duy nhất từng dòng, tạo mã ổn định từ mã báo cáo + số dòng gốc. Không xóa dòng hoàn tiền chỉ vì trùng mã đơn.'],
['Ví dụ TikTok','Cuối kỳ 1.160.000. Các đối chiếu tổng khớp; TT-002 chờ QT đúng hạn, P2 đang chuyển 400.000.'],
['Ví dụ Shopee','Cuối kỳ 1.408.000, tổng vẫn khớp. SP-001 thiếu 2.000 so dự kiến do phí ship 22.000 thay vì 20.000.'],
['Ví dụ Lazada','Cuối kỳ tính lại 1.406.000, nguồn 1.405.000. LZ-P1 mới nhận 999.000 trên số phải nhận 1.000.000.'],
['Tiêu chí hoàn tất','Đủ nguồn, hết lỗi dữ liệu, mọi chênh đã có chứng cứ xử lý và từng khoản treo xác định được. Không dùng điều chỉnh để ép khớp.']
];
guide.forEach((r,i)=>{let rr=i+5;sh.Huong_dan.getRange(`A${rr}:B${rr}`).values=[r];sh.Huong_dan.getRange(`A${rr}`).format.font.bold=true;sh.Huong_dan.getRange(`B${rr}`).format.wrapText=true;sh.Huong_dan.getRange(`A${rr}:B${rr}`).format.rowHeight=r[1].length>130?48:36;});
put('Huong_dan','D5','Nguồn tham khảo');
const sources=[
['TikTok: giao dịch đã QT','https://seller-vn.tiktok.com/university/essay?knowledge_id=2824323762210561&lang=vi-VN'],
['TikTok: giao dịch chờ QT','https://seller-vn.tiktok.com/university/essay?knowledge_id=2844649149974273&lang=vi-VN'],
['Shopee: báo cáo thu nhập','https://banhang.shopee.vn/edu/article/2072'],
['Shopee: báo cáo doanh thu','https://banhang.shopee.vn/edu/article/2066'],
['Lazada: dữ liệu tài chính','https://open.lazada.com/apps/doc/doc?docId=108249&nodeId=10556']
];
sh.Huong_dan.getRange('D6:E10').values=sources;sh.Huong_dan.getRange('E6:E10').format.wrapText=true;sh.Huong_dan.getRange('D6:E10').format.rowHeight=46;
put('Huong_dan','D12','Nguồn số liệu mẫu');put('Huong_dan','E12','Số giả lập do người lập mẫu tạo, không sử dụng dữ liệu thực tế của doanh nghiệp.');sh.Huong_dan.getRange('E12').format.wrapText=true;
put('Huong_dan','D14','Tên cột dữ liệu gốc');put('Huong_dan','E14','Có thể khác theo phiên bản báo cáo của từng sàn. Ánh xạ theo ý nghĩa khoản tiền, ngày, shop và mã giao dịch.');sh.Huong_dan.getRange('E14').format.wrapText=true;
const sourceGuide=[
['Hồ sơ cần lấy','Tải file gốc, lưu kỳ/ngày tải/shop. Giữ báo cáo đơn hàng, giao dịch tài chính, rút tiền, sao kê NH, số dư cuối kỳ và sổ chi tiết.'],
['TikTok Shop','Lấy đơn hàng; tài chính đã quyết toán và chưa quyết toán; đợt rút tiền. Tách giao dịch vào số dư khỏi báo Có ngân hàng.'],
['Shopee','Lấy báo cáo đơn hàng, thu nhập/chi tiết quyết toán, giao dịch số dư và rút tiền. Theo dõi hoàn tiền/điều chỉnh cả đơn kỳ trước.'],
['Lazada','Lấy báo cáo đơn hàng, sao kê tài chính/chi tiết giao dịch và lịch sử chi trả. Giữ mã đơn, giao dịch và đợt chi trả để nối dữ liệu.'],
['Ngân hàng / kế toán','Sao kê báo Có tháng chốt, báo Có kỳ sau cho khoản treo; sổ chi tiết khoản phải thu/số dư sàn và tiền đang chuyển cùng phạm vi.'],
['Kiểm tra đủ dữ liệu','Đối chiếu tổng theo file gốc và từng shop: quyết toán, rút tiền, báo Có và số dư. Kiểm tra đủ ngày, không trùng lần tải, không sót hoàn tiền.']
];
sourceGuide.forEach((r,i)=>{const rr=17+i;sh.Huong_dan.getRange(`D${rr}:E${rr}`).values=[r];sh.Huong_dan.getRange(`D${rr}`).format.font.bold=true;sh.Huong_dan.getRange(`E${rr}`).format.wrapText=true;sh.Huong_dan.getRange(`D${rr}:E${rr}`).format.rowHeight=54;});

wb.recalculate();
// Export an independent compact audit report, not extra user deliverables.
const report={};
for(const [name,range] of [['Tong_hop','A9:M37'],['So_du','P6:AC9'],['Don_hang','N6:U15'],['Chuyen_tien','I6:O15'],['Giao_dich_san','N6:T20']]){
  const v=await wb.inspect({kind:'table',range:`${name}!${range}`,include:'values,formulas',tableMaxRows:35,tableMaxCols:16});report[`${name}!${range}`]=v.ndjson;
}
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30},summary:'formula error scan'});
report.errors=errors.ndjson;
await fs.writeFile(path.join(support,'inspection.json'),JSON.stringify(report,null,2));
console.log('SUMMARY_VALUES',JSON.stringify(sh.Tong_hop.getRange('A10:M13').values));
console.log('ISSUE_COUNTS',JSON.stringify(sh.Tong_hop.getRange('A27:B37').values));
console.log('ERROR_SCAN',errors.ndjson);

// Representative edits in this disposable in-memory state, restoring all cells afterward.
const tests=[];
function assert(condition,msg){if(!condition)throw new Error(msg);tests.push(msg);}
function num(s,a){return sh[s].getRange(a).values[0][0];}
assert(num('Tong_hop','H10')===1160000,'TikTok closing balance 1,160,000');
assert(num('Tong_hop','H11')===1408000,'Shopee closing balance 1,408,000');
assert(num('Tong_hop','H12')===1406000,'Lazada closing balance 1,406,000');
assert(num('Tong_hop','K12')===1000,'Lazada independent balance mismatch 1,000');
assert(num('Don_hang','R11')===-2000,'Shopee order fee variance -2,000');
assert(num('Chuyen_tien','L14')===1000,'Lazada payout short receipt 1,000');
assert(num('Tong_hop','B33')===1,'Overdue/under-settled orders visible');
put('Ngan_hang','F14',1000000);wb.recalculate();
assert(num('Tong_hop','K12')===0,'Changing bank input removes Lazada difference');
put('Ngan_hang','F14',999000);
put('Giao_dich_san','K14',20000);put('Giao_dich_san','N14',1825000);wb.recalculate();
assert(num('Don_hang','R11')===0,'Corrected Shopee shipping recalculates order');
put('Giao_dich_san','K14',22000);put('Giao_dich_san','N14',1823000);
put('So_du','D7',null);wb.recalculate();
assert(num('Tong_hop','H10')==='n.a.','Missing opening balance is not a valid zero');
put('So_du','D7',1000000);
put('Ngan_hang','D8','UNKNOWN');wb.recalculate();
assert(num('Ngan_hang','L8')==='Chưa ghép đợt chuyển','Unmatched bank reference is flagged');
put('Ngan_hang','D8','TT-P1');
put('Giao_dich_san','C8','TT-F01');wb.recalculate();
assert(num('Giao_dich_san','T8')==='Trùng giao dịch','Duplicate transaction is flagged');
put('Giao_dich_san','C8','TT-F02');
put('Ngan_hang','E9',dt('2026-09-30'));wb.recalculate();
assert(num('Chuyen_tien','L9')===0,'Last-day credit is included in selected month');
put('Ngan_hang','E9',dt('2026-10-02'));wb.recalculate();
assert(num('Chuyen_tien','L9')===400000,'Next-month credit remains in transit');
put('Tong_hop','C4',dt('2026-10-01'));wb.recalculate();
assert(num('So_du','R7')==='Sai ngày chốt','Changing month does not validate stale snapshots');
put('Tong_hop','C4',dt('2026-09-01'));

// Next table row: structured references must grow with table expansion.
const bankTable=sh.Ngan_hang.tables.items[0];
bankTable.rows.add(null,[['TikTok Shop','DEMO_TT','NH-TEST','TT-P1',dt('2026-09-20'),1000,'DEMO112','Temporary extension test',null,null,null,null]]);
for(let c=9;c<=12;c++){
  sh.Ngan_hang.getRange(`${letters(c)}107`).copyFrom(sh.Ngan_hang.getRange(`${letters(c)}106`),'formulas');
}
wb.recalculate();
// Refresh the engine's parsed structured reference after a structural table edit.
// Excel's independent calculation test also verifies table growth without this refresh.
const payoutFormula=sh.Chuyen_tien.getRange('J8').formulas[0][0];
put('Chuyen_tien','J8',null);fx('Chuyen_tien','J8',payoutFormula);wb.recalculate();
assert(num('Chuyen_tien','J8')===991000,'Appended table record feeds SUMIFS after formula copy');
sh.Ngan_hang.getRange('A107:L107').copyFrom(sh.Ngan_hang.getRange('A106:L106'),'all');
wb.recalculate();
assert(num('Tong_hop','H10')===1160000,'All mutation tests restored');
await fs.writeFile(path.join(support,'verification.json'),JSON.stringify({tests},null,2));
console.log('VERIFIED',tests.length);

const renderRanges={Tong_hop:'A2:M22',Don_hang:'A6:M15',Giao_dich_san:'G6:T15',Chuyen_tien:'A6:O15',Ngan_hang:'A6:L15',So_du:'A6:N9',Huong_dan:'A2:B17'};
for(const [sheetName,range] of Object.entries(renderRanges)){
  const blob=await wb.render({sheetName,range,scale:1.4,format:'png'});
  await fs.writeFile(path.join(support,`${sheetName}.png`),new Uint8Array(await blob.arrayBuffer()));
  console.log('RENDERED',sheetName);
}
for(const [sheetName,range,file] of [
  ['Tong_hop','A25:M44','Tong_hop_kiem_tra'],
  ['Don_hang','N6:U15','Don_hang_ket_qua'],
  ['So_du','P6:AC9','So_du_ket_qua'],
  ['Huong_dan','A18:B42','Huong_dan_chi_tiet'],
  ['Huong_dan','D14:E22','Nguon_du_lieu']
]){
  const blob=await wb.render({sheetName,range,scale:1.3,format:'png'});
  await fs.writeFile(path.join(support,`${file}.png`),new Uint8Array(await blob.arrayBuffer()));
}
const out=await SpreadsheetFile.exportXlsx(wb);
const outputPath=path.join(dir,'Mau_doi_soat_cong_no_san_theo_thang.xlsx');
await out.save(outputPath);
console.log('SAVED',outputPath);
