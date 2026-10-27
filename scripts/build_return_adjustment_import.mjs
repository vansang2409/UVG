import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const work='D:/UVG/scripts/data/amis-2024-linked-20261026';
const p=JSON.parse(await fs.readFile(`${work}/return-adjustment-import.json`,'utf8'));
const w=Workbook.create(),s=w.worksheets.add('Import dieu chinh');
s.showGridLines=false;s.getRange('A1:Z225').format={font:{name:'Times New Roman',size:11},verticalAlignment:'center',rowHeight:42};
for(let i=0;i<7;i++){s.mergeCells(`A${i+1}:Z${i+1}`);s.getRange(`A${i+1}`).values=[[p.instructions[i]]];s.getRange(`A${i+1}:Z${i+1}`).format={rowHeight:i===4||i===6?36:23,wrapText:true};}
s.getRange('A1').format.font={name:'Times New Roman',size:16,bold:true};
s.getRange('A9:Z9').values=[p.headers];s.getRange('A9:Z9').format={fill:'#CCCCFF',font:{name:'Times New Roman',size:11,bold:true},rowHeight:66,wrapText:true};
s.getRange('A10:Z225').values=p.rows;
const widths=[16,18,38,50,20,38,25,24,23,23,26,22,25,76,90,12,16,23,24,26,20,26,28,20,29,29];widths.forEach((x,i)=>s.getRangeByIndexes(0,i,1,1).format.columnWidth=x);
for(const c of ['B','L'])s.getRange(`${c}10:${c}225`).setNumberFormat('dd/mm/yyyy');
for(const c of ['E','J','K','M','Y','Z'])s.getRange(`${c}10:${c}225`).setNumberFormat('@');
s.getRange('Q10:W225').setNumberFormat('#,##0.########;[Red](#,##0.########);0');
s.getRange('U10:U225').setNumberFormat('0');s.getRange('N10:O225').format.wrapText=true;s.getRange('B10:B225').format.fill='#FFF2CC';s.getRange('J10:M225').format.fill='#FFF2CC';s.freezePanes.freezeRows(9);s.freezePanes.freezeColumns(1);
const t=w.worksheets.add('Cho KT tra lai');t.getRange('A1:S1').values=[p.pending_headers];t.getRange('A2:S4').values=p.pending.map(r=>r.map((v,i)=>i===2?new Date(v):v));t.getRange('A1:S4').format={font:{name:'Arial',size:10},rowHeight:58,wrapText:true,verticalAlignment:'center',columnWidth:24};t.getRange('A1:S1').format={fill:'#24435C',font:{bold:true,color:'#FFFFFF'},rowHeight:60};t.getRange('C2:C4').setNumberFormat('dd/mm/yyyy');t.getRange('H2:L4').setNumberFormat('#,##0.########');t.getRange('E1:E4').format.columnWidth=70;t.getRange('S1:S4').format.columnWidth=110;t.freezePanes.freezeRows(1);
await(await SpreadsheetFile.exportXlsx(w)).save(`${work}/_return-import.tmp.xlsx`);
for(const [range,name]of[['N8:W12','return-import-money'],['X8:Z12','return-import-map']]){const png=await w.render({sheetName:'Import dieu chinh',range,scale:1,format:'png'});await fs.writeFile(`${work}/${name}.png`,new Uint8Array(await png.arrayBuffer()));}
console.log('Prepared 161 adjustments / 216 rows; two BTL held separately.');
