import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const h='D:/UVG/scripts/data/check-review';
const d=JSON.parse(await fs.readFile(`${h}/zero-detail-review.json`,'utf8'));
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${h}/before-zero-notes.xlsx`));
const s=w.worksheets.getItem('Dieu chinh');
const before=await w.render({sheetName:'Dieu chinh',range:'Y138:AB140',scale:1,format:'png'});
await fs.writeFile(`${h}/zero-notes-before.png`,new Uint8Array(await before.arrayBuffer()));
for(const x of d.roots){
 s.getRange(`AB${x.row}`).values=[[x.note]];
 s.getRange(`AB${x.row}`).format.wrapText=true;
 // Mở chiều cao đúng các dòng ghi chú được bổ sung.
 s.getRange(`A${x.row}:BC${x.row}`).format.rowHeight=Math.max(s.getRange(`AB${x.row}`).format.rowHeight||48,Math.ceil(x.note.length/115)*15+12);
}
w.recalculate();
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#SPILL!',options:{useRegex:true,maxResults:20},maxChars:600})).ndjson);
for(const [range,name] of [['Y138:AB140','zero-notes-mixed.png'],['Y176:AB178','zero-notes-sale.png']]){
 const b=await w.render({sheetName:'Dieu chinh',range,scale:1,format:'png'});await fs.writeFile(`${h}/${name}`,new Uint8Array(await b.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(w)).save(`${h}/with-zero-notes.xlsx`);
console.log(JSON.stringify({noted:d.roots.length,counts:d.counts}));
