import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile,Workbook} from '@oai/artifact-tool';

const root='D:/UVG';
const work=`${root}/scripts/data/replacement-duplicate-2000-20261101`;
const outputDir=`${root}/outputs/thay-the-trung-2000-20261101`;
const p=JSON.parse(await fs.readFile(`${work}/prepared.json`,'utf8'));
await fs.mkdir(outputDir,{recursive:true});
const template=await SpreadsheetFile.importXlsx(await FileBlob.load(p.template));
const templatePreview=await template.render({sheetName:'Hóa đơn GTGT',range:'J9:V11',scale:1,format:'png'});
await fs.writeFile(`${work}/template-preview.png`,new Uint8Array(await templatePreview.arrayBuffer()));
const w=Workbook.create();
const s=w.worksheets.add('Hóa đơn GTGT');
const last=p.rows.length+9;
s.showGridLines=false;
s.getRange(`A1:X${last}`).format={font:{name:'Times New Roman',size:11},verticalAlignment:'center',rowHeight:44};
const widths=[16,16,34,20,45,38,24,20,22,22,24,22,22,22,34,90,12,16,20,22,19,25,90,26];
widths.forEach((v,i)=>s.getRangeByIndexes(0,i,1,1).format.columnWidth=v);
for(let i=0;i<7;i++){
 s.mergeCells(`A${i+1}:X${i+1}`);
 s.getRange(`A${i+1}`).values=[[p.notes[i]]];
 s.getRange(`A${i+1}:X${i+1}`).format={rowHeight:i===6?30:23,wrapText:true};
}
s.getRange('A1').format.font={name:'Times New Roman',size:16,bold:true};
s.getRange('A2').format.font={name:'Times New Roman',size:11,bold:true,color:'#B00020'};
s.getRange('A7').format.font={name:'Times New Roman',size:11,color:'#B00020'};
s.getRange('A8:X8').format.rowHeight=12;
s.getRange('A9:X9').values=[p.headers];
s.getRange('A9:X9').format={fill:'#CCCCFF',font:{name:'Times New Roman',size:11,bold:true},rowHeight:74,wrapText:true,horizontalAlignment:'center',verticalAlignment:'center'};
s.getRange('P9:T9').format.fill='#FFFF00';
for(const c of ['D','H','I','L','M','O'])s.getRange(`${c}10:${c}${last}`).setNumberFormat('@');
for(const c of ['B','N'])s.getRange(`${c}10:${c}${last}`).setNumberFormat('dd/mm/yyyy');
for(let i=0;i<p.rows.length;i+=500)s.getRangeByIndexes(i+9,0,Math.min(500,p.rows.length-i),24).values=p.rows.slice(i,i+500);
s.getRange(`A10:X${last}`).format.wrapText=true;
s.getRange(`A10:X${last}`).format.rowHeight=52;
s.getRange(`A10:A${last}`).setNumberFormat('0');
s.getRange(`R10:V${last}`).setNumberFormat('#,##0;[Red]-#,##0;0');
s.getRange(`U10:U${last}`).setNumberFormat('0"%"');
s.getRange(`X10:X${last}`).setNumberFormat('0');
s.freezePanes.freezeRows(9);
s.freezePanes.freezeColumns(1);
w.recalculate();
const inspection=await w.inspect({kind:'table',range:"'Hóa đơn GTGT'!L9:X12",include:'values,formulas',tableMaxRows:4,tableMaxCols:13,maxChars:3500});
const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:10},maxChars:1000});
await fs.writeFile(`${work}/inspection.ndjson`,`${inspection.ndjson}\n${errors.ndjson}`);
const filename=`${outputDir}/IMPORT_THAY_THE_VE_0_2000_HD_TRUNG_2024.xlsx`;
await(await SpreadsheetFile.exportXlsx(w)).save(filename);
for(const [range,name]of[['L9:O12','references'],['P9:V12','zero-money'],['U9:X12','reason-note'],['P2007:X2009','last-lines'],['A9:F12','buyers']]){
 const png=await w.render({sheetName:s.name,range,scale:1,format:'png'});
 await fs.writeFile(`${work}/${name}.png`,new Uint8Array(await png.arrayBuffer()));
}
console.log(inspection.ndjson);
console.log(errors.ndjson);
console.log(JSON.stringify({output:filename,invoices:p.invoices,lines:p.lines,new_date:p.new_date,one_note_line_per_invoice:true}));
