import fs from 'node:fs/promises';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';
const dir='D:/UVG/outputs/bo-sung-13-20261029';
const p=JSON.parse(await fs.readFile(`${dir}/prepared.json`,'utf8'));
const w=Workbook.create(),s=w.worksheets.add('Import dieu chinh');
const last=9+p.rows.length;
s.showGridLines=false;
s.getRange(`A1:X${last}`).format={font:{name:'Times New Roman',size:11},verticalAlignment:'center',columnWidth:20};
s.getRange('A1:X9').values=p.headers;
for(let row=1;row<=7;row++){
  s.mergeCells(`A${row}:X${row}`);
  s.getRange(`A${row}:X${row}`).format={rowHeight:24,wrapText:true};
}
s.getRange('A1').format.font={name:'Times New Roman',size:15,bold:true};
s.getRange('A2').format.font={name:'Times New Roman',size:11,bold:true,color:'#B91C1C'};
s.getRange('A9:X9').format={fill:'#DBEAFE',font:{name:'Times New Roman',size:11,bold:true},rowHeight:66,wrapText:true,horizontalAlignment:'center'};
// Giữ nguyên dữ liệu 24 cột của 206 dòng; không thêm dòng tổng trong vùng nhập.
s.getRange(`A10:X${last}`).values=p.rows;
s.getRange(`A10:X${last}`).format={rowHeight:62,wrapText:true};
s.getRange(`A10:A${last}`).setNumberFormat('0');
for(const c of ['E','I','J','K','M'])s.getRange(`${c}10:${c}${last}`).setNumberFormat('@');
s.getRange(`B10:B${last}`).setNumberFormat('dd/mm/yyyy');
s.getRange(`B10:B${last}`).format.fill='#FFF2CC';
s.getRange(`Q10:S${last}`).setNumberFormat('#,##0;[Red]-#,##0;0');
for(const c of ['T','V','W'])s.getRange(`${c}10:${c}${last}`).setNumberFormat('#,##0;[Red]-#,##0;0');
s.getRange(`U10:U${last}`).setNumberFormat('0');
s.getRange('A:A').format.columnWidth=12;
for(const c of ['B','J','K','L','P','U','X'])s.getRange(`${c}:${c}`).format.columnWidth=16;
s.getRange('C:D').format.columnWidth=36;
s.getRange('F:F').format.columnWidth=36;
s.getRange('N:O').format.columnWidth=80;
s.freezePanes.freezeRows(9);s.freezePanes.freezeColumns(1);
w.recalculate();
console.log((await w.inspect({kind:'table',range:`'Import dieu chinh'!Q9:W12`,include:'values,formulas',tableMaxRows:4,tableMaxCols:7,maxChars:1400})).ndjson);
for(const [range,name] of [['J9:O12','references'],['O9:W12','money'],['O210:W215','last-lines']]){
  const png=await w.render({sheetName:s.name,range,scale:1,format:'png'});
  await fs.writeFile(`${dir}/${name}.png`,new Uint8Array(await png.arrayBuffer()));
}
const out=`${dir}/IMPORT_DIEU_CHINH_BO_SUNG_13_HD.xlsx`;
await (await SpreadsheetFile.exportXlsx(w)).save(out);
console.log(JSON.stringify({output:out,...p.summary}));
