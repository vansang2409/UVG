import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import {fileURLToPath} from 'node:url';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';
const h=path.dirname(fileURLToPath(import.meta.url));
const root=path.resolve(h,'../../..');
const p=path.join(root,'outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx');
const d=JSON.parse(await fs.readFile(path.join(h,'supplement-57-price-changes.json'),'utf8'));
const previewOnly=process.argv.includes('--preview-only');
if(!previewOnly && crypto.createHash('sha256').update(await fs.readFile(p)).digest('hex')!==d.base_sha256) throw new Error('Workbook changed since audit');
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(previewOnly?path.join(h,'supplement-57-preview.xlsx'):p));
for(const c of previewOnly?[]:d.changes) {
  const r=w.worksheets.getItem(c.sheet).getRange(c.cell);
  r.values=[[c.new]];
  r.setNumberFormat('#,##0.0;[Red](#,##0.0);0');
}
w.recalculate();
const row=Number(d.changes[0].cell.slice(1));
const preview=await w.render({sheetName:'Import dieu chinh',range:`O${row}:S${row+1}`,scale:1.5,format:'png'});
await fs.writeFile(path.join(h,'supplement-57-prices.png'),new Uint8Array(await preview.arrayBuffer()));
if(!previewOnly) await(await SpreadsheetFile.exportXlsx(w)).save(path.join(h,'supplement-57-staged.xlsx'));
console.log(previewOnly?'Rendered staged OOXML subset':`Staged ${d.changes.length} unit-price cells`);
