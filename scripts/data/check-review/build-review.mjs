import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const root='D:/UVG', helper=path.join(root,'scripts/data/check-review'), out=path.join(root,'outputs/kiem-tra-20261008');
await fs.mkdir(out,{recursive:true});
const data=JSON.parse(await fs.readFile(path.join(helper,'review.json'),'utf8'));
const w=Workbook.create(), s=w.worksheets.add('Can kiem tra'), t=w.worksheets.add('So lieu doi chieu');
data.items.forEach(x => x.status=x.status.replaceAll(',', ' -'));
const start=7, last=start+data.items.length-1;
function base(sh,cols){sh.showGridLines=false;sh.getRangeByIndexes(0,0,last,cols).format.font={name:'Arial',size:10};sh.getRangeByIndexes(0,0,last,cols).format.verticalAlignment='center';sh.freezePanes.freezeRows(6);sh.freezePanes.freezeColumns(4);}
base(s,11);base(t,12);
s.getRange('A2').values=[['UVG — Các trường hợp cần kiểm tra']];s.getRange('A2').format.font={name:'Arial',size:15,bold:true};
s.getRange('A3').values=[['19 chuỗi: 17 ngoài ngoại lệ, 2 ngoại lệ. Kiểm tra các điểm đã ghi nhận; chưa phải rà toàn bộ hóa đơn.']];
s.getRange('A4').values=[['Điền các cột H–J. Chỉ chọn Đã xác minh khi đã ghi kết quả và bằng chứng; không thay số liệu nguồn ở tab đối chiếu.']];
s.getRange('A5').values=[['Phạm vi F0 01/01/2024–30/06/2025, liên quan đến 31/12/2025. Số tiền bằng đồng; ô trống được giữ nguyên.']];
const headers=['STT','Nhóm xử lý','Hóa đơn gốc F0','Hóa đơn cần kiểm tra','Vấn đề','Nội dung đã ghi nhận','Cần kiểm tra / xác nhận','Trạng thái','Kết quả kiểm tra của bạn','File / căn cứ bổ sung','Nguồn đã dùng'];
s.getRange('A6:K6').values=[headers];
s.getRange(`A${start}:K${last}`).values=data.items.map((x,i)=>[i+1,x.group,x.root,x.target,x.category,x.issue,x.action,x.status,null,null,x.provenance]);
s.tables.add(`A6:K${last}`,true,'DanhSachKiemTra');
s.getRange(`H${start}:H${last}`).dataValidation={rule:{type:'list',values:['Chờ kiểm tra','Có căn cứ 0 ở chi tiết - chờ chốt','Đã chốt tạm - chờ căn cứ gốc','Đã xác nhận ngày - chờ XML','Đang kiểm tra','Đã xác minh','Chưa đủ căn cứ']}};
s.getRange(`H${start}:J${last}`).format.fill='#FFF4CF';
s.getRange(`H${start}:H${last}`).conditionalFormats.add('containsText',{text:'Đã xác minh',format:{fill:'#DCF0E1',font:{color:'#245733'}}});
s.getRange(`H${start}:H${last}`).conditionalFormats.add('containsText',{text:'Chưa đủ căn cứ',format:{fill:'#FCE1DE',font:{color:'#9C3025'}}});
t.getRange('A2').values=[['UVG — Số liệu để đối chiếu']];t.getRange('A2').format.font={name:'Arial',size:15,bold:true};
t.getRange('A3').values=[['E–G: số của hóa đơn cần kiểm tra trong bảng tổng; I: cộng dồn cả chuỗi theo report hiện có.']];
t.getRange('A4').values=[['H = thanh toán − (tiền hàng + thuế), chỉ tính khi đủ 3 số. Ô trống không đổi thành 0.']];
t.getRange('A5').values=[['Các nguồn và ghi chú gốc ở K–L; kết quả kiểm tra nhập tại tab Can kiem tra.']];
t.getRange('A6:L6').values=[['STT','Hóa đơn gốc F0','Hóa đơn cần kiểm tra','Ngày theo report','Tiền hàng bảng tổng','Thuế bảng tổng','Thanh toán bảng tổng','Lệch số học','Cộng dồn thanh toán chuỗi','Chuỗi hiện có','Ghi chú report gốc','Nguồn đã dùng']];
t.getRange(`A${start}:L${last}`).values=data.items.map((x,i)=>[i+1,x.root,x.target,new Date(x.day+'Z'),x.goods,x.tax,x.payment,null,x.cumulative,x.chain,x.original_note,x.provenance]);
t.getRange(`H${start}:H${last}`).formulas=data.items.map((x,i)=>[`=IF(COUNT(E${start+i}:G${start+i})=3,G${start+i}-SUM(E${start+i}:F${start+i}),"")`]);
t.tables.add(`A6:L${last}`,true,'SoLieuDoiChieu');
t.getRange(`D${start}:D${last}`).setNumberFormat('dd/mm/yyyy');t.getRange(`E${start}:I${last}`).setNumberFormat('#,##0;[Red](#,##0);0');
t.getRange(`H${start}:H${last}`).conditionalFormats.add('cellIs',{operator:'notEqual',formula:0,format:{fill:'#FCE1DE',font:{color:'#9C3025',bold:true}}});
const widthsS=[6,27,24,24,29,65,69,32,46,40,85], widthsT=[6,24,24,18,23,20,25,19,27,58,65,85];
for(const [sh,widths] of [[s,widthsS],[t,widthsT]]){
 widths.forEach((v,i)=>sh.getRangeByIndexes(5,i,last-5,1).format.columnWidth=v);
 sh.getRangeByIndexes(5,0,1,widths.length).format={fill:'#243746',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,horizontalAlignment:'center',verticalAlignment:'center',rowHeight:42};
 sh.getRangeByIndexes(6,0,data.items.length,widths.length).format.wrapText=true;
 sh.getRangeByIndexes(6,0,data.items.length,widths.length).format.rowHeight=100;
 for(let i=0;i<3;i++)sh.getRangeByIndexes(2+i,0,1,1).format.font={name:'Arial',size:10,italic:true};
}
w.recalculate();
console.log((await w.inspect({kind:'table',range:'Can kiem tra!A6:E10',include:'values',tableMaxRows:5,tableMaxCols:5,maxChars:1800})).ndjson);
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#SPILL!',options:{useRegex:true,maxResults:20},maxChars:1000})).ndjson);
for(const [sheetName,range,file] of [['Can kiem tra','A2:G9','review-left.png'],['Can kiem tra','H6:K9','review-input.png'],['So lieu doi chieu','A6:I10','review-values.png']]){
 const preview=await w.render({sheetName,range,scale:1,format:'png'});await fs.writeFile(path.join(helper,file),new Uint8Array(await preview.arrayBuffer()));
}
const dest=path.join(out,'CAC_TRUONG_HOP_CAN_KIEM_TRA.xlsx');await (await SpreadsheetFile.exportXlsx(w)).save(dest);
console.log(JSON.stringify({saved:dest,cases:data.items.length}));

