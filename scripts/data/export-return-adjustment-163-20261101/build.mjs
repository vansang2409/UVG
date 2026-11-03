import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';

const root='D:/UVG';
const work=`${root}/scripts/data/export-return-adjustment-163-20261101`;
const outputDir=`${root}/outputs/import-dieu-chinh-163-20261101`;
const source=JSON.parse(await fs.readFile(`${work}/source-cells.json`,'utf8'));
const original=source.sheets.find(s=>s.properties.sheetId===700000001);
if(original.properties.title!=='Import dieu chinh')throw new Error('Unexpected source sheet');
const cells=original.data[0].rowData;
const cellValue=c=>{
 const v=c?.effectiveValue;
 if(!v)return null;
 if(v.errorValue)throw new Error(JSON.stringify(v.errorValue));
 return v.stringValue??v.numberValue??v.boolValue??null;
};
const values=Array.from({length:228},(_,r)=>Array.from({length:26},(_,c)=>cellValue(cells[r]?.values?.[c])));
const body=values.slice(9);
const totals=[18,21,22].map(c=>body.reduce((sum,r)=>sum+(r[c]??0),0));
if(body.length!==219||new Set(body.map(r=>r[0])).size!==163||JSON.stringify(totals)!==JSON.stringify([-225865926,-18069275,-243935201]))throw new Error('Source controls failed');

const color=f=>{
 const v=f?.backgroundColorStyle?.rgbColor??f?.backgroundColor;
 return v?'#'+['red','green','blue'].map(k=>Math.round((v[k]??0)*255).toString(16).padStart(2,'0')).join(''):null;
};
const w=Workbook.create();
const s=w.worksheets.add('Import dieu chinh');
s.showGridLines=false;
s.getRange('A1:Z228').format={font:{name:'Times New Roman',size:11},verticalAlignment:'center',rowHeight:52};
const widths=[16,18,38,50,20,38,25,24,23,23,26,22,28,110,90,12,16,23,24,26,20,26,28,20,30,29];
widths.forEach((v,i)=>s.getRangeByIndexes(0,i,1,1).format.columnWidth=v);
// Write cached values from Google Sheets, including lookup results as literal text.
// This gives the import file no dependencies on the other tabs in the source.
for(const c of ['E','J','K','M','Y','Z'])s.getRange(`${c}10:${c}228`).setNumberFormat('@');
for(const c of ['B','L'])s.getRange(`${c}10:${c}228`).setNumberFormat('dd/mm/yyyy');
s.getRange('A1:Z228').values=values;
for(let r=1;r<=7;r++){
 s.mergeCells(`A${r}:Z${r}`);
 s.getRange(`A${r}:Z${r}`).format={rowHeight:r===5||r===7?32:23,wrapText:true};
 const font=cells[r-1].values[0].userEnteredFormat?.textFormat??{};
 s.getRange(`A${r}`).format.font={name:font.fontFamily??'Times New Roman',size:font.fontSize??11,bold:font.bold??false};
}
s.getRange('A8:Z8').format.rowHeight=12;
s.getRange('A9:Z9').format={font:{name:'Times New Roman',size:11,bold:true},wrapText:true,rowHeight:74,horizontalAlignment:'center',verticalAlignment:'center'};
for(let c=0;c<26;c++){
 const fill=color(cells[8].values[c].userEnteredFormat);
 if(fill)s.getRangeByIndexes(8,c,1,1).format.fill=fill;
}
s.getRange('A10:Z228').format.wrapText=true;
// Preserve source input fills and number formats, grouped by adjacent rows.
for(let c=0;c<26;c++){
 for(let start=9;start<228;){
  const fill=color(cells[start].values[c].userEnteredFormat);
  let end=start+1;
  while(end<228&&color(cells[end].values[c].userEnteredFormat)===fill)end++;
  if(fill)s.getRangeByIndexes(start,c,end-start,1).format.fill=fill;
  start=end;
 }
 const f=cells[9].values[c].userEnteredFormat?.numberFormat;
 let pattern=f?.type==='TEXT'?'@':f?.pattern;
 const numbers=body.map(r=>r[c]).filter(v=>typeof v==='number');
 if(c>=16&&c<=22&&c!==20&&numbers.length&&numbers.every(Number.isInteger))pattern='#,##0;[Red](#,##0);0';
 if(pattern)s.getRangeByIndexes(9,c,219,1).setNumberFormat(pattern);
}
s.freezePanes.freezeRows(9);
w.recalculate();
const inspection=await w.inspect({kind:'table',range:"'Import dieu chinh'!J9:N12",include:'values,formulas',tableMaxRows:4,tableMaxCols:5,maxChars:2600});
const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:10},maxChars:1000});
await fs.writeFile(`${work}/inspection.ndjson`,`${inspection.ndjson}\n${errors.ndjson}`);
await fs.mkdir(outputDir,{recursive:true});
const filename=`${outputDir}/IMPORT_DIEU_CHINH_163_PHIEU_TRA_HANG_2024_MOI_NHAT.xlsx`;
await(await SpreadsheetFile.exportXlsx(w)).save(filename);
for(const [range,name]of[['A9:B13','dates'],['J9:N12','references-reasons'],['O9:W13','return-money'],['J226:N228','last-reasons']]){
 const png=await w.render({sheetName:s.name,range,scale:1,format:'png'});
 await fs.writeFile(`${work}/${name}.png`,new Uint8Array(await png.arrayBuffer()));
}
console.log(inspection.ndjson);
console.log(errors.ndjson);
console.log(JSON.stringify({output:filename,invoices:163,lines:219,money:totals}));
