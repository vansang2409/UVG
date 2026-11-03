import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';

const work='D:/UVG/scripts/data/add-return-sku-20261101';
const p=JSON.parse(await fs.readFile(`${work}/prepared.json`,'utf8'));
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(p.backup));
const s=w.worksheets.getItem(p.sheet);
if(process.argv.includes('--help-filter')){
 console.log(w.help('worksheet.autoFilter',{include:'index,examples,notes',maxChars:3000}).ndjson);
}else if(process.argv.includes('--preview-only')){
 const png=await w.render({sheetName:s.name,range:'V9:Z13',scale:1,format:'png'});
 await fs.writeFile(`${work}/before.png`,new Uint8Array(await png.arrayBuffer()));
 console.log(JSON.stringify({preview:`${work}/before.png`}));
}else{
 s.getRange('Z9:Z228').copyFrom(s.getRange('Y9:Y228'),'all');
 s.getRange('Z9').values=[[p.header]];
 s.getRange('Z10:Z228').values=p.mapping.map(r=>[r.code]);
 s.getRange('Z10:Z228').setNumberFormat('@');
 s.getRange('Z1:Z228').format.columnWidth=32;
 w.recalculate();
 const inspected=await w.inspect({kind:'table',range:"'Import dieu chinh'!Y9:Z13",include:'values,formulas',tableMaxRows:5,tableMaxCols:2,maxChars:1400});
 const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:10},maxChars:1000});
 await fs.writeFile(`${work}/inspection.ndjson`,`${inspected.ndjson}\n${errors.ndjson}`);
 await(await SpreadsheetFile.exportXlsx(w)).save(p.staged);
 for(const [range,name]of[['V9:Z13','after'],['Y225:Z228','last-codes']]){
  const png=await w.render({sheetName:s.name,range,scale:1,format:'png'});
  await fs.writeFile(`${work}/${name}.png`,new Uint8Array(await png.arrayBuffer()));
 }
 console.log(inspected.ndjson);
 console.log(errors.ndjson);
 console.log(JSON.stringify({staged:p.staged,rows:p.rows,unique_codes:Object.keys(p.sku_counts).length}));
}
