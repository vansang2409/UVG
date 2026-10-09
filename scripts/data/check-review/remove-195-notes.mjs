import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const h='D:/UVG/scripts/data/check-review';
const d=JSON.parse(await fs.readFile(`${h}/remove-195-notes-data.json`,'utf8'));
console.log('IMPORT');
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${h}/before-remove-195-notes.xlsx`));
const s=w.worksheets.getItem('Dieu chinh');
let b=await w.render({sheetName:'Dieu chinh',range:`AB${d.changes[0].row}:AB${d.changes[0].row}`,scale:1,format:'png'});
await fs.writeFile(`${h}/remove-notes-before.png`,new Uint8Array(await b.arrayBuffer()));
for(const x of d.changes){s.getRange(`AB${x.row}`).values=[[x.note]];s.getRange(`AB${x.row}`).format.rowHeight=x.height;}
w.recalculate();
b=await w.render({sheetName:'Dieu chinh',range:`AB${d.changes[0].row}:AB${d.changes[0].row}`,scale:1,format:'png'});
await fs.writeFile(`${h}/remove-notes-after.png`,new Uint8Array(await b.arrayBuffer()));
console.log('EXPORT');
await(await SpreadsheetFile.exportXlsx(w)).save(`${h}/without-195-notes.xlsx`);
console.log('STAGED');
