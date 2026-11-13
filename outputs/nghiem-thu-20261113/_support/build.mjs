import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const out=path.dirname(fileURLToPath(import.meta.url));
const records=JSON.parse(await fs.readFile(path.join(out,'mapping.json'),'utf8'));
const expected=JSON.parse(await fs.readFile(path.join(out,'summary.json'),'utf8'));
const wb=Workbook.create();
const cover=wb.worksheets.add('Tổng hợp');
const headers=['Chứng từ AMIS','Ngày AMIS','Số HĐ','Ký hiệu HĐ','Ngày HĐ','Kết quả đối chiếu','Tiền hàng AMIS (đ)','VAT AMIS (đ)','Thanh toán AMIS (đ)','Thanh toán HĐ/nhóm (đ)','VAT HĐ/nhóm (đ)','Lệch thanh toán nhóm (đ)','Lệch VAT nhóm (đ)','CT chung HĐ','HĐ đi kèm (ký hiệu/số)','HĐ khác và liên kết cũ','Ghi chú cần kiểm tra','Căn cứ map','Nguồn dòng AMIS','Nguồn dòng HĐ'];
const fmt='#,##0;[Red](#,##0);0';
const col=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s};
const asDate=s=>s?new Date(s+'T00:00:00Z'):null;
const joined=a=>a.join('; ');
function matrix(r){return [r.doc,asDate(r.day),r.invoice,r.symbol,asDate(r.invoice_day),r.status,...r.amis_money,r.mi_money[2],r.mi_money[1],r.delta[2],r.delta[1],r.group_docs.length>1?joined(r.group_docs):null,joined(r.related),joined([...new Set([...r.history,...(!r.invoice?r.candidates:[])])]),joined(r.issues),joined(r.evidence),r.source_amis,r.source_invoice]}
const specs=[];
const review=records.filter(r=>!['Khớp tiền và VAT','Khớp theo nhóm chứng từ'].includes(r.status));
for(const [name,rs] of [['Map 2024',records.filter(r=>r.year===2024)],['Map 2025',records.filter(r=>r.year===2025)],['Cần kiểm tra',review]]){
  const s=wb.worksheets.add(name),end=rs.length+6;specs.push({name,count:rs.length,end});
  const compact=name!=='Cần kiểm tra',ncols=compact?4:20,last=compact?'D':'T';
  s.showGridLines=false;
  s.getRange('A2').values=[[name==='Cần kiểm tra'?'Các chứng từ cần kiểm tra':`Map số và ký hiệu HĐ — ${name.slice(-4)}`]];
  s.getRange('A3').values=[['Nguồn 13/11/2026. Mỗi dòng là một chứng từ AMIS; số HĐ lưu dạng chuỗi 8 chữ số.']];
  s.getRange('A4').values=[[compact?'Lọc Kết quả để tìm dòng cần kiểm tra; chi tiết ngoại lệ ở tab Cần kiểm tra.':'Cột J:M tính theo nhóm CT ở cột N và lặp theo từng CT. Chênh lệch = AMIS − HĐ.']];
  s.getRange(`A6:${last}6`).values=[compact?[headers[0],headers[2],headers[3],headers[5]]:headers];
  console.log(JSON.stringify({stage:'writing',sheet:name,rows:rs.length}));
  s.getRangeByIndexes(6,0,rs.length,ncols).values=rs.map(r=>compact?[r.doc,r.invoice,r.symbol,r.status]:matrix(r));
  console.log(JSON.stringify({stage:'styling',sheet:name}));
  s.tables.add(`A6:${last}${end}`,true,'MapTable'+specs.length);
  s.getRange(`A1:${last}${Math.min(end,80)}`).format.font={name:'Arial',size:10,color:'#202124'};
  s.getRange(`A1:${last}${Math.min(end,80)}`).format.verticalAlignment='center';
  s.getRange(`A6:A${Math.min(end,80)}`).format.rowHeight=23;
  s.getRange('A2').format.font={name:'Arial',size:15,bold:true,color:'#202124'};
  s.getRange('A3:A4').format.font={name:'Arial',size:10,italic:true,color:'#59636F'};
  s.getRange(`A6:${last}6`).format={fill:'#253B53',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:40,horizontalAlignment:'center',verticalAlignment:'center'};
  const widths=compact?[185,115,115,285]:[185,105,105,110,105,240,155,145,165,170,155,175,155,235,260,300,420,440,510,470];
  widths.forEach((w,i)=>s.getRange(`${col(i+1)}1`).format.columnWidthPx=w);
  if(compact)s.getRange(`B7:C${end}`).setNumberFormat('@');
  else{s.getRange(`B7:B${end}`).setNumberFormat('dd/mm/yyyy');s.getRange(`E7:E${end}`).setNumberFormat('dd/mm/yyyy');s.getRange(`C7:D${end}`).setNumberFormat('@');s.getRange(`G7:M${end}`).setNumberFormat(fmt);}
  s.freezePanes.freezeRows(6);s.freezePanes.freezeColumns(compact?1:4);
  const st=compact?'D':'F';
  s.getRange(`${st}7:${st}${end}`).conditionalFormats.addCustom(`OR($${st}7="Chưa map",$${st}7="Nhiều ứng viên",$${st}7="Đã map - lệch tiền/VAT")`,{fill:'#FCE7E7',font:{color:'#9C2626'}});
  s.getRange(`${st}7:${st}${end}`).conditionalFormats.addCustom(`OR($${st}7="Đã map - thiếu số tiền",$${st}7="Đã map - cần kiểm tra nguồn",$${st}7="Lệch làm tròn 1-2đ")`,{fill:'#FFF1D6',font:{color:'#875D0F'}});
  if(name==='Cần kiểm tra')s.tabColor='#C08226';
}
cover.showGridLines=false;cover.tabColor='#253B53';
cover.getRange('A2').values=[['Map số và ký hiệu hóa đơn với AMIS']];
cover.getRange('A3').values=[['Dữ liệu AMIS 2024–2025 và meInvoice từ 01/01/2024 đến 13/11/2026. Đơn vị: đồng.']];
cover.getRange('A5:F5').values=[['Năm AMIS','Số chứng từ','Đã có số/ký hiệu HĐ','Chưa chọn được HĐ','Tiền hàng AMIS','VAT AMIS']];
for(const [i,y] of [2024,2025].entries()){
  const row=i+6,sp=specs[i],ref=`'${sp.name}'`;
  cover.getRange(`A${row}`).values=[[y]];
  cover.getRange(`B${row}:D${row}`).formulas=[[`=COUNTA(${ref}!A7:A${sp.end})`,`=COUNTA(${ref}!B7:B${sp.end})`,`=B${row}-C${row}`]];
  cover.getRange(`E${row}:F${row}`).values=[[expected.amis_totals[y][0],expected.amis_totals[y][1]]];
}
cover.getRange('A8').values=[['Tổng']];cover.getRange('B8:F8').formulas=[['=SUM(B6:B7)','=SUM(C6:C7)','=SUM(D6:D7)','=SUM(E6:E7)','=SUM(F6:F7)']];
cover.getRange('A11:C11').values=[['Kết quả đối chiếu','2024','2025']];
const statuses=[...new Set(records.map(r=>r.status))];
statuses.forEach((status,i)=>{
  const rr=i+12;cover.getRange(`A${rr}`).values=[[status]];
  cover.getRange(`B${rr}:C${rr}`).formulas=[specs.slice(0,2).map(sp=>`=COUNTIFS('${sp.name}'!D7:D${sp.end},A${rr})`)];
});
const notes=[
  'Số/ký hiệu HĐ đã map không đồng nghĩa toàn bộ nghiệp vụ đã được nghiệm thu.',
  'Liên kết chứng từ lấy từ các bảng đã đối soát 26/10–11/11, kiểm lại HĐ và số tiền bằng nguồn 13/11.',
  'HĐ đã bị thay thế/hủy không cộng vào số tiền hiện hành. Phiếu BTL/BGG map riêng HĐ điều chỉnh.',
  'BTL00475 và BHHT2025031047 cùng map HĐ thay thế về 0: 00186384 – 1C26MUV; đối chiếu theo nhóm.',
  'Ô tiền trống được giữ trống. Chỉ bổ sung từ bảng chi tiết/liên kết khi có giá trị nguồn và ghi rõ ở cột Q.',
  'Tab Cần kiểm tra: cột J:M lặp số tiền nhóm, không cộng dồn. Cột G:I là tiền riêng từng CT.',
  'Map theo ký hiệu + số HĐ và liên kết chứng từ. Chưa đối chiếu từng SKU/SL, XML hoặc CQT trực tiếp.',
  'Tab Cần kiểm tra là bản trích các dòng ngoại lệ của hai tab Map, không cộng thêm vào tổng AMIS.'
];
cover.getRange('A22').values=[['Cách đọc và phạm vi']];notes.forEach((n,i)=>cover.getRange(`A${i+24}`).values=[[n]]);
cover.getRange('A34').values=[['Nguồn AMIS']];cover.getRange('B34').values=[['So_chi_tiet_ban_hang (2024).xlsx; So_chi_tiet_ban_hang (2025).xlsx — thư mục data/1311/amis']];
cover.getRange('A35').values=[['Nguồn meInvoice']];cover.getRange('B35').values=[['4 ZIP bảng kê tổng, chi tiết, điều chỉnh, thay thế — thư mục data/1311/meinvoice']];
cover.getRange('A36').values=[['Số HĐ bảng kê tổng']];cover.getRange('B36').values=[[expected.ledger]];
cover.getRange('A37').values=[['Số khóa HĐ có chi tiết']];cover.getRange('B37').values=[[expected.detail_keys]];
cover.getRange('A38').values=[['Số liên kết mới']];cover.getRange('B38').values=[[expected.edges]];
cover.getRange('A39').values=[['Số dòng tab Cần kiểm tra']];cover.getRange('B39').values=[[review.length]];
cover.getRange('A1:H39').format.font={name:'Arial',size:10,color:'#202124'};
cover.getRange('A1:H39').format.verticalAlignment='center';cover.getRange('A1:A39').format.rowHeight=24;
cover.getRange('A2').format.font={name:'Arial',size:15,bold:true,color:'#202124'};
cover.getRange('A3').format.font.italic=true;
for(const range of ['A5:F5','A11:C11'])cover.getRange(range).format={fill:'#253B53',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',rowHeight:38,wrapText:true};
[330,145,175,175,180,180,90,90].forEach((w,i)=>cover.getRange(`${col(i+1)}1`).format.columnWidthPx=w);
cover.getRange('B6:D8').setNumberFormat('#,##0');cover.getRange('E6:F8').setNumberFormat(fmt);cover.getRange('B12:C20').setNumberFormat('#,##0');
cover.getRange('A8:F8').format.font.bold=true;cover.getRange('A22').format.font.bold=true;
wb.recalculate();
const total=cover.getRange('B8:D8').values[0];
if(total[0]!==expected.docs||total[1]!==expected.mapped||total[2]!==expected.unmapped)throw new Error('Counts do not reconcile '+JSON.stringify(total));
for(const [i,y] of [2024,2025].entries()){
  const totals=cover.getRange(`E${i+6}:F${i+6}`).values[0];
  if(totals[0]!==expected.amis_totals[y][0]||totals[1]!==expected.amis_totals[y][1])throw new Error('AMIS totals mismatch');
}
console.log((await wb.inspect({kind:'region',sheetId:'Tổng hợp',range:'A5:F8',maxChars:2500,tableMaxRows:4,tableMaxCols:6})).ndjson);
for(const name of ['Tổng hợp',...specs.map(s=>s.name)]){
  const blob=await wb.render({sheetName:name,range:name==='Tổng hợp'?'A1:F19':name==='Cần kiểm tra'?'A1:I10':'A1:D10',scale:1,format:'png'});
  await fs.writeFile(path.join(out,`preview-${name}.png`),new Uint8Array(await blob.arrayBuffer()));
}
const output=await SpreadsheetFile.exportXlsx(wb);await output.save(path.join(out,'MAP_SO_KY_HIEU_HD_AMIS_1311.xlsx'));
await fs.writeFile(path.join(out,'workbook-specs.json'),JSON.stringify({specs,review:review.length,counts:total},null,2));
console.log(JSON.stringify({stage:'exported',records:records.length,review:review.length}));
