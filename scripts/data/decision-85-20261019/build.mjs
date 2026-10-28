import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';

const here=path.dirname(fileURLToPath(import.meta.url));
const root=path.resolve(here,'../../..');
const d=JSON.parse(await fs.readFile(path.join(here,'data.json'),'utf8'));
const w=Workbook.create();
const summary=w.worksheets.add('Tong hop');
const detail=w.worksheets.add('Khong can dieu chinh');
const navy='#22364D', ink='#243142', muted='#596778', pale='#EDF2F7';
const money='#,##0;[Red](#,##0);0';
const put=(s,c,v)=>{s.getRange(c).values=[[v]];};
for(const s of [summary,detail]){
  s.showGridLines=false;
  s.getRange(s===summary?'A1:D24':'A1:H78').format={font:{name:'Arial',size:10,color:ink},rowHeight:25,verticalAlignment:'center'};
}
summary.tabColor=navy;
summary.getRange('A1:A24').format.columnWidth=55;
summary.getRange('B1:B24').format.columnWidth=18;
summary.getRange('C1:C24').format.columnWidth=3;
summary.getRange('D1:D24').format.columnWidth=75;
put(summary,'A2','SỐ HÓA ĐƠN CẦN XUẤT SAU ĐỐI CHIẾU');
summary.getRange('A2').format.font={name:'Arial',size:14,bold:true,color:navy};
summary.getRange('A2:D2').format.borders={bottom:{style:'thin',color:navy}};
const headlineRows=[
 ['Chỉ tiêu','Số hóa đơn','Diễn giải'],
 ['Tập kế toán sau loại trùng',d.source_invoices,'10.442 hóa đơn trong bộ hiện tại + 85 hóa đơn đối chiếu.'],
 ['Không cần xuất điều chỉnh',null,'62 đã khớp + 6 cùng chuỗi sản phẩm + 4 chỉ đổi thông tin.'],
 ['Cần xuất theo phương án mới',null,'Số hóa đơn riêng biệt trừ số không cần điều chỉnh.'],
 ['Đã có trong bộ import hiện tại',d.baseline_to_issue,'8.410 chưa điều chỉnh + 2.032 đã điều chỉnh.'],
 ['Cần bổ sung vào bộ import',null,'13 hóa đơn trên 8 chuỗi; chưa bổ sung vào các file import.']
];
summary.getRange('A4:B9').values=headlineRows.map(r=>r.slice(0,2));
summary.getRange('D4:D9').values=headlineRows.map(r=>[r[2]]);
summary.getRange('B6').formulas=[["=COUNTA('Khong can dieu chinh'!D7:D78)"]];
summary.getRange('B7').formulas=[['=B5-B6']];
summary.getRange('B9').formulas=[['=B7-B8']];
summary.getRange('B5:B9').setNumberFormat('#,##0');
summary.getRange('B5:B9').format.horizontalAlignment='right';
summary.getRange('D5:D9').format.wrapText=true;
summary.getRange('A5:D9').format.rowHeight=38;
for(const r of [7,9]){
  summary.getRange(`A${r}:B${r}`).format.fill=pale;
  summary.getRange(`A${r}:B${r}`).format.font={name:'Arial',size:10,bold:true,color:ink};
}
const groupRows=[
 ['Nhóm không cần xuất điều chỉnh','Số hóa đơn','Lý do'],
 ...d.groups.map((g,i)=>[g.label,null,[
  'Chi tiết và số tiền đã khớp; tổng các hóa đơn liên quan = 0.',
  'Cùng chuỗi sản phẩm, chi tiết và số tiền đã khớp; tổng chuỗi = 0.',
  '3 đổi thông tin người mua + 1 đổi tên hàng; không ảnh hưởng tiền hàng, tiền thuế.'
 ][i]]),
 ['Tổng không cần xuất',null,'Danh sách từng hóa đơn tại sheet Khong can dieu chinh.']
];
summary.getRange('A12:B16').values=groupRows.map(r=>r.slice(0,2));
summary.getRange('D12:D16').values=groupRows.map(r=>[r[2]]);
for(let i=0;i<3;i++)summary.getRange(`B${13+i}`).formulas=[[`=COUNTIFS('Khong can dieu chinh'!B7:B78,A${13+i})`]];
summary.getRange('B16').formulas=[['=SUM(B13:B15)']];
summary.getRange('B13:B16').setNumberFormat('#,##0');
summary.getRange('A13:D16').format.rowHeight=38;
summary.getRange('D13:D16').format.wrapText=true;
summary.getRange('A16:B16').format.borders={top:{style:'thin',color:'#B5C1CD'}};
summary.getRange('A16:B16').format.font={name:'Arial',size:10,bold:true,color:ink};
for(const range of ['A4:B4','D4','A12:B12','D12'])summary.getRange(range).format={fill:navy,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center',rowHeight:31};
const notes=[
 'Giữ 00006108, 00004203, 00004204 vì chỉ đổi thông tin người mua; vẫn xử lý hóa đơn gốc.',
 'Cộng 00009186 chỉ đổi tên hàng: tổng 4 hóa đơn không đổi tiền hàng, tiền thuế.',
 '13 hóa đơn bổ sung = 9 thuộc 6 cặp trước đây giữ riêng + 4 hóa đơn dùng MST cá nhân.',
 '10.455 là số theo phương án mới. Các file import hiện tại vẫn có 10.442 hóa đơn.'
];
for(let i=0;i<notes.length;i++)put(summary,`A${18+i}`,notes[i]);
summary.getRange('A18:A21').format.font={name:'Arial',size:10,color:muted};
put(summary,'A23','Nguồn: danh sách 85 hóa đơn đối chiếu kế toán và xác nhận mới ngày 19/10/2026.');
summary.getRange('A23').format.font={name:'Arial',size:10,italic:true,color:muted};

put(detail,'A2','72 HÓA ĐƠN KHÔNG CẦN XUẤT ĐIỀU CHỈNH');
detail.getRange('A2').format.font={name:'Arial',size:14,bold:true,color:navy};
detail.getRange('A2:H2').format.borders={bottom:{style:'thin',color:navy}};
put(detail,'A3','Trong nhóm 85 hóa đơn đối chiếu; các hóa đơn chỉ đổi thông tin không đổi tiền được giữ nguyên, hóa đơn gốc tiếp tục xử lý.');
detail.getRange('A3').format.font={name:'Arial',size:10,color:muted};
const header=['STT','Nhóm','Ký hiệu','Số hóa đơn','Ngày hóa đơn','Hóa đơn gốc của chuỗi','Tên người mua','Lý do không cần xuất điều chỉnh'];
detail.getRange('A6:H6').values=[header];
const rows=d.records.map((r,i)=>[i+1,r.group,r.key[0],r.key[1],new Date(r.date+'T00:00:00Z'),r.root,r.buyer,r.reason]);
detail.getRange('A7:H78').values=rows;
const widths=[6,31,12,14,14,26,40,67];
for(let i=0;i<widths.length;i++)detail.getRange(`${String.fromCharCode(65+i)}1:${String.fromCharCode(65+i)}78`).format.columnWidth=widths[i];
detail.getRange('A6:H6').format={fill:navy,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center',rowHeight:42,wrapText:true};
detail.getRange('A7:H78').format.rowHeight=47;
for(const col of ['B','F','G','H'])detail.getRange(`${col}7:${col}78`).format.wrapText=true;
detail.getRange('D7:D78').setNumberFormat('00000000');
detail.getRange('E7:E78').setNumberFormat('dd/mm/yyyy');
detail.getRange('A7:A78').setNumberFormat('0');
detail.getRange('A7:A78').format.horizontalAlignment='right';
detail.getRange('C7:F78').format.horizontalAlignment='left';
detail.getRange('E7:E78').format.horizontalAlignment='right';
const table=detail.tables.add('A6:H78',true,'HoaDonKhongCanDieuChinh');
table.style='TableStyleMedium2';table.showFilterButton=true;
detail.freezePanes.freezeRows(6);detail.freezePanes.freezeColumns(4);

w.recalculate();
const check=await w.inspect({kind:'region',sheetId:'Tong hop',range:'A4:C16',maxChars:2600,tableMaxRows:13,tableMaxCols:3});
await fs.writeFile(path.join(here,'summary-inspect.ndjson'),check.ndjson);
const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30}});
await fs.writeFile(path.join(here,'formula-errors.ndjson'),errors.ndjson);
for(const [name,range,file] of [['Tong hop','A1:D24','summary'],['Khong can dieu chinh','A1:H12','detail-first'],['Khong can dieu chinh','A67:H78','detail-last']]){
  const b=await w.render({sheetName:name,range,scale:1,format:'png'});
  await fs.writeFile(path.join(here,file+'.png'),new Uint8Array(await b.arrayBuffer()));
}
const output=path.join(root,d.output);
await(await SpreadsheetFile.exportXlsx(w)).save(output);
console.log('EXPORTED',d.records.length,'invoices; planned total',d.new_to_issue);
