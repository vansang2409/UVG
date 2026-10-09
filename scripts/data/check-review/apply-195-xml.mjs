import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const h='D:/UVG/scripts/data/check-review';
const d=JSON.parse(await fs.readFile(`${h}/195-xml-data.json`,'utf8'));
console.log('IMPORT BASE',new Date().toISOString());
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${h}/before-195-xml.xlsx`));
console.log('BASE LOADED',new Date().toISOString());
const main=w.worksheets.getItem('Dieu chinh'),imp=w.worksheets.getItem('Import dieu chinh');
const sample=d.changes[0].row;
let b=await w.render({sheetName:'Dieu chinh',range:`AC${sample}:AN${sample+1}`,scale:1,format:'png'});
await fs.writeFile(`${h}/195-before-plan.png`,new Uint8Array(await b.arrayBuffer()));
for(const c of d.changes){
 for(const [col,value] of Object.entries(c.values))main.getRange(`${col}${c.row}`).values=[[['AE','AK'].includes(col)?new Date(`${value}T00:00:00Z`):value]];
 main.getRange(`AB${c.row}`).format.wrapText=true;
 main.getRange(`AB${c.row}`).format.rowHeight=Math.max(180,main.getRange(`AB${c.row}`).format.rowHeight||0);
}
const rows=d.rows.map(r=>{const v=[...r];v[1]=new Date(`${v[1]}T00:00:00Z`);return v;});
for(let i=0;i<rows.length;i++)imp.getRange(`A${d.start_row+i}:X${d.start_row+i}`).copyFrom(imp.getRange('A40786:X40786'),'all');
imp.getRange(`A${d.start_row}:X${d.last_row}`).values=rows;
imp.getRange(`A${d.start_row}:X${d.last_row}`).format.rowHeight=62;
imp.getRange(`B${d.start_row}:B${d.last_row}`).setNumberFormat('dd/mm/yyyy');
imp.getRange(`E${d.start_row}:E${d.last_row}`).setNumberFormat('@');
imp.getRange(`K${d.start_row}:K${d.last_row}`).setNumberFormat('00000000');
imp.getRange(`C${d.start_row}:C${d.last_row}`).format.wrapText=true;
imp.getRange(`N${d.start_row}:O${d.last_row}`).format.wrapText=true;
imp.getRange(`Q${d.start_row}:W${d.last_row}`).setNumberFormat('#,##0;[Red](#,##0);0');
imp.getRange('A3').values=[['- Ngày hóa đơn: nhóm cũ 08/10/2026; nhóm làm lại 195 chuỗi (STT 1601–1990) ngày 09/10/2026. Có thể sửa đồng nhất các dòng cùng STT trước khi import.']];
imp.getRange('A4').values=[[`- 986 chuỗi: 1.990 hóa đơn mới, ${d.total_lines.toLocaleString('en-US').replaceAll(',','.')} dòng. Gồm 791 chuỗi cũ và 195 chuỗi làm lại đã chốt theo XML.`]];
imp.getRange('A7').values=[['- 19 chuỗi chờ KT giữ riêng; 42 chuỗi đã 0 còn lại chưa thêm import. Nhóm làm lại 195 chuỗi: đảo tổng XML cũ rồi giảm chi tiết gốc. Thành tiền gốc 0 thì Hàng khuyến mãi = 1.']];
const guide=w.worksheets.getItem('Huong dan');
guide.getRange('A13:A16').values=[[
 `Import dieu chinh: 986 chuỗi / 1.990 hóa đơn mới / ${d.total_lines.toLocaleString('en-US').replaceAll(',','.')} dòng. Đã thêm 195 chuỗi / 390 hóa đơn / 491 dòng làm lại theo XML (STT 1601–1990).`
],['Cho KT xac nhan: 19 chuỗi giữ riêng. Trong 237 chuỗi đã 0, đã mở 195 chuỗi cùng dạng 00000508 theo chốt của người dùng; 42 chuỗi còn lại chưa thêm import.'],['F4 đảo khoản điều chỉnh cũ theo tổng XML; F5 giảm đầy đủ chi tiết F0 cho 195 chuỗi. Ngày dự thảo nhóm mới 09/10/2026. Tên khách hàng bằng Người mua hàng.'],['195 chuỗi làm lại phải giữ tiền hàng, thuế và thanh toán cộng dồn bằng 0. Nguồn chi tiết gốc giữ số tiền thực tế, không thay bằng phép nhân SL × đơn giá. Chưa import, ký hoặc phát hành.']];
w.recalculate();console.log('UPDATED',new Date().toISOString());
console.log((await w.inspect({kind:'table',range:`Dieu chinh!BA${sample}:BC${sample}`,include:'values,formulas',tableMaxRows:1,tableMaxCols:3,maxChars:1000})).ndjson);
for(const [sheetName,range,name] of [['Dieu chinh',`AC${sample}:AN${sample+1}`,'195-after-plan.png'],['Dieu chinh',`AB${sample}:AB${sample}`,'195-note.png'],['Import dieu chinh',`J${d.start_row}:X${d.start_row+2}`,'195-import.png']]){
 b=await w.render({sheetName,range,scale:1,format:'png'});await fs.writeFile(`${h}/${name}`,new Uint8Array(await b.arrayBuffer()));
}
console.log('EXPORT',new Date().toISOString());
await(await SpreadsheetFile.exportXlsx(w)).save(`${h}/with-195-xml.xlsx`);
console.log('STAGED',JSON.stringify({invoices:d.invoices,roots:d.roots,lines:d.total_lines}),new Date().toISOString());
