import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const root='D:/UVG',h=`${root}/scripts/data/untreated`;
const d=JSON.parse(await fs.readFile(`${h}/untreated.json`,'utf8'));
const w=Workbook.create(),s=w.worksheets.add('Chua dieu chinh'),g=w.worksheets.add('Huong dan');
const last=d.groups.length+2;
const headers=['Ký hiệu','Số hóa đơn','Ngày','Tiền hàng','Thuế','Thanh toán','Tên người mua','MST/CCCD nguyên bản','Trạng thái nguồn','Số dòng chi tiết gốc','Dòng thành tiền khác 0','Dòng thành tiền bằng 0','Phân loại số tiền','Cần kiểm tra','Nguồn bảng tổng','Nguồn chi tiết'];
s.getRange('A2:P2').values=[headers];
function date(t){const [dd,mm,yy]=t.split('/').map(Number);return new Date(Date.UTC(yy,mm-1,dd));}
s.getRange(`A3:P${last}`).values=d.groups.map(x=>{
 const inv=x.invoices[0],raw=x.detail[0],c=x.line_counts;
 return [inv[0],inv[1],date(inv[2]),...inv.slice(4,7),raw[5],raw[6],raw[8],c.lines,c.nonzero_goods||0,c.zero_goods||0,x.money_class,x.notes||null,raw[12],x.detail_source];
});
const t=s.tables.add(`A2:P${last}`,true,'ChuaDieuChinhChuoi');t.style='TableStyleMedium2';t.showFilterButton=true;
s.showGridLines=false;s.getRange(`A1:P${last}`).format.font={name:'Arial',size:10};s.getRange(`A1:P${last}`).format.verticalAlignment='center';
s.getRange('A1:P2').format={fill:'#243746',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center',wrapText:true};
for(const [range,title] of [['A1:F1','F0 — Hóa đơn gốc chưa điều chỉnh trong kỳ'],['G1:H1','Thông tin người mua'],['I1:M1','Trạng thái và dòng chi tiết'],['N1:P1','Ghi chú và căn cứ nguồn']]){s.mergeCells(range);s.getRange(range).values=[[title]];}
s.getRange('A1:P1').format.rowHeight=28;s.getRange('A2:P2').format.rowHeight=42;s.getRange(`A3:P${last}`).format.rowHeight=38;
const widths=[16,15,16,22,20,24,58,28,33,20,22,22,30,105,105,90];
widths.forEach((v,i)=>s.getRangeByIndexes(0,i,last,1).format.columnWidth=v);
s.getRange(`B3:B${last}`).setNumberFormat('00000000');s.getRange(`H3:H${last}`).setNumberFormat('@');s.getRange(`C3:C${last}`).setNumberFormat('dd/mm/yyyy');s.getRange(`D3:F${last}`).setNumberFormat('#,##0;[Red](#,##0);0');s.getRange(`J3:L${last}`).setNumberFormat('0');
s.getRange(`D3:F${last}`).format.horizontalAlignment='right';s.getRange(`G3:I${last}`).format.wrapText=true;s.getRange(`M3:P${last}`).format.wrapText=true;
for(let i=0;i<d.groups.length;i++){
 const x=d.groups[i];const height=Math.max(38,Math.ceil(x.name.length/60)*14+10,Math.ceil((x.notes||'').length/135)*14+10,Math.ceil(x.detail_source.length/115)*14+10);
 if(height>38)s.getRange(`A${i+3}:P${i+3}`).format.rowHeight=height;
}
s.freezePanes.freezeRows(2);s.freezePanes.freezeColumns(2);
s.getRange(`M3:M${last}`).conditionalFormats.add('containsText',{text:'Thiếu số tiền',format:{fill:'#FFF0C2',font:{color:'#8A5200'}}});
const notes=[
 'UVG — Nhóm ngoài ngoại lệ chưa có điều chỉnh trong kỳ',
 'F0 có ngày 01/01/2024–30/06/2025; kiểm liên kết điều chỉnh/thay thế đến 31/12/2025.',
 '8.410 F0: 62 năm 2024 và 8.348 tháng 1–6/2025. Không trùng nhóm ngoại lệ và nhóm đã điều chỉnh.',
 'Đã loại 6 chuỗi thay thế và 10 hóa đơn gốc có trạng thái đã hủy khỏi nhóm này.',
 'Chưa điều chỉnh là chưa có liên kết điều chỉnh trong bộ nguồn đến 31/12/2025; chưa xác minh các lần xử lý năm 2026.',
 'Tiền hàng/thuế/thanh toán lấy trực tiếp cột P/Q/R của 10 phần bảng kê tổng theo ký hiệu + số hóa đơn; không cộng các dòng chi tiết để thay số tổng.',
 '6 F0 trống thanh toán bảng tổng; cả 6 có dòng chi tiết thanh toán 0. Giữ trống bảng tổng và ghi ở cột N, chưa áp dụng bổ sung số 0.',
 'Đã đối chiếu 13.570 dòng chi tiết; 952 F0 có ít nhất một dòng thành tiền bằng 0. Theo quy tắc người dùng, chỉ dòng thành tiền gốc = 0 mới nhập 1 cột Hàng khuyến mãi khi lập mẫu.',
 'Lọc M = Thiếu số tiền bảng tổng hoặc N khác trống để kiểm 6 trường hợp; lọc L > 0 để xem hóa đơn có dòng thành tiền bằng 0.',
 'Ngoại lệ giữ nguyên quy tắc tên công ty/doanh nghiệp hoặc định danh dạng MST trong nguồn, bảo vệ theo liên kết cùng chuỗi; chưa xác thực MST.',
 'Giữ nguyên bộ nguồn và hai report đang dùng. Bảng này chưa có kế hoạch hóa đơn mới, chưa tạo import, ký hoặc phát hành.',
 'Nguồn tổng và nguồn chi tiết của mỗi hóa đơn ghi ngay tại O–P. Số hóa đơn và định danh giữ dạng văn bản; đơn vị tiền là đồng.'
];
g.showGridLines=false;g.getRange(`A1:A${notes.length}`).values=notes.map(v=>[v]);g.getRange(`A1:A${notes.length}`).format={font:{name:'Arial',size:10},wrapText:true,verticalAlignment:'center',columnWidth:145,rowHeight:38};g.getRange('A1').format.font.bold=true;
w.recalculate();
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#SPILL!',options:{useRegex:true,maxResults:20},maxChars:600})).ndjson);
for(const [sheetName,range,name] of [['Chua dieu chinh','A1:F6','untreated-values.png'],['Chua dieu chinh','I1:P6','untreated-notes.png'],['Huong dan','A1:A12','untreated-guide.png']]){
 const b=await w.render({sheetName,range,scale:1,format:'png'});await fs.writeFile(`${h}/${name}`,new Uint8Array(await b.arrayBuffer()));
}
const p=`${root}/outputs/HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx`;
await (await SpreadsheetFile.exportXlsx(w)).save(p);console.log(JSON.stringify({saved:p,rows:d.groups.length}));

