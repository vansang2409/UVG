import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile,Workbook} from '@oai/artifact-tool';
const root='D:/UVG',work=`${root}/scripts/data/replacement-zero-20261029`;
const p=JSON.parse(await fs.readFile(`${work}/prepared.json`,'utf8'));
const out=`${root}/outputs/amis-2024-theo-btl-bhh-20261026/IMPORT_THAY_THE_VE_0_6_CHUOI_2024.xlsx`;
const template=await SpreadsheetFile.importXlsx(await FileBlob.load(p.template));
const ts=template.worksheets.getItem('Hóa đơn GTGT');
const w=Workbook.create(),s=w.worksheets.add('Hóa đơn GTGT'),last=p.rows.length+9;
// Xem bố cục mẫu trước khi điền, giữ nguyên tên sheet và thứ tự 22 cột.
await fs.writeFile(`${work}/template-preview.png`,new Uint8Array(await (await template.render({sheetName:ts.name,range:'L9:V11',scale:1,format:'png'})).arrayBuffer()));
// Dựng vùng nhập sạch để không mang theo công thức ví dụ rỗng của mẫu.
s.getRange('A1:V9').values=ts.getRange('A1:V9').values;
s.getRange(`A1:V${last}`).format={font:{name:'Times New Roman',size:11},verticalAlignment:'center',columnWidth:22};
for(let row=1;row<=7;row++){
 s.mergeCells(`A${row}:V${row}`);
 s.getRange(`A${row}:V${row}`).format={rowHeight:23,wrapText:true};
}
s.getRange('A1').format.font={name:'Times New Roman',size:16,bold:true};
s.getRange('A2').format.font={name:'Times New Roman',size:11,bold:true,color:'#FF0000'};
s.getRange('A5').format.font={name:'Times New Roman',size:11,color:'#FF0000'};
s.getRange('A9:V9').values=[p.headers];
s.getRange('A9:V9').format.fill='#CCCCFF';
s.getRange('P9:T9').format.fill='#FFFF00';
s.getRange(`A10:V${last}`).values=p.rows;
s.getRange(`A10:V${last}`).format={font:{name:'Times New Roman',size:11},verticalAlignment:'center',rowHeight:58,wrapText:true,borders:{preset:'all',style:'thin',color:'#D9D9D9'}};
s.getRange('A9:V9').format={rowHeight:70,wrapText:true,font:{name:'Times New Roman',size:11,bold:true}};
s.getRange(`A10:A${last}`).setNumberFormat('0');
for(const c of ['D','H','I','L','M','O'])s.getRange(`${c}10:${c}${last}`).setNumberFormat('@');
for(const c of ['B','N'])s.getRange(`${c}10:${c}${last}`).setNumberFormat('dd/mm/yyyy');
s.getRange(`R10:V${last}`).setNumberFormat('#,##0;[Red]-#,##0;0');
s.getRange(`U10:U${last}`).setNumberFormat('0');
s.getRange(`B10:B${last}`).format.fill='#FFF2CC';
s.getRange(`F10:F${last}`).format.columnWidth=30;
s.getRange(`K9:K${last}`).format.columnWidth=23;
s.getRange(`L9:O${last}`).format.columnWidth=22;
s.getRange(`P9:P${last}`).format.columnWidth=90;
s.getRange(`Q9:Q${last}`).format.columnWidth=12;
s.getRange(`R9:V${last}`).format.columnWidth=18;
s.freezePanes.freezeRows(9);s.freezePanes.freezeColumns(1);
w.recalculate();
console.log((await w.inspect({kind:'table',range:`'Hóa đơn GTGT'!L9:V12`,include:'values,formulas',tableMaxRows:4,tableMaxCols:11,maxChars:2400})).ndjson);
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!',options:{useRegex:true,maxResults:10},maxChars:1000})).ndjson);
for(const [range,name] of [['J9:O12','references'],['P9:V12','zero-money'],['P185:V190','last-lines']]){
 const png=await w.render({sheetName:s.name,range,scale:1,format:'png'});
 await fs.writeFile(`${work}/${name}.png`,new Uint8Array(await png.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(w)).save(out);
console.log(JSON.stringify({output:out,invoices:p.audit.length,lines:p.rows.length}));
