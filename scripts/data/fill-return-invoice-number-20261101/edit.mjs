import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const work='D:/UVG/scripts/data/fill-return-invoice-number-20261101';
const p=JSON.parse(await fs.readFile(`${work}/prepared.json`,'utf8'));
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(p.backup));
const s=w.worksheets.getItem(p.sheet);
if(process.argv.includes('--preview-only')){
 const png=await w.render({sheetName:s.name,range:'I9:M15',scale:1,format:'png'});
 await fs.writeFile(`${work}/before.png`,new Uint8Array(await png.arrayBuffer()));
}else{
 for(const c of p.changes)s.getRange(`K${c.row}`).values=[[c.number]];
 w.recalculate();
 const inspect=await w.inspect({kind:'table',range:"'Import dieu chinh'!J9:K15",include:'values,formulas',tableMaxRows:7,tableMaxCols:2,maxChars:1300});
 const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!',options:{useRegex:true,maxResults:10},maxChars:700});
 await fs.writeFile(`${work}/inspection.ndjson`,`${inspect.ndjson}\n${errors.ndjson}`);
 await(await SpreadsheetFile.exportXlsx(w)).save(p.staged);
 const png=await w.render({sheetName:s.name,range:'I9:M15',scale:1,format:'png'});
 await fs.writeFile(`${work}/after.png`,new Uint8Array(await png.arrayBuffer()));
 console.log(inspect.ndjson);
 console.log(errors.ndjson);
 console.log(JSON.stringify({cells_filled:p.changes.length,staged:p.staged}));
}
