import fs from 'node:fs/promises';
import { createReadStream } from 'node:fs';
import readline from 'node:readline';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const dir = 'D:/UVG/outputs/merge_20261012';
const manifest = JSON.parse(await fs.readFile(`${dir}/merge_manifest.json`, 'utf8'));
const columns = manifest.columns;
const n = manifest.row_count;
const workbook = Workbook.create();
const sheet = workbook.worksheets.add('Hóa đơn gộp');
sheet.showGridLines = false;
sheet.tabColor = '#284B63';
sheet.getRangeByIndexes(0, 0, 1, columns.length).values = [columns];

let batch = [], written = 0;
const lines = readline.createInterface({ input: createReadStream(`${dir}/rows.jsonl`, { encoding: 'utf8' }), crlfDelay: Infinity });
function flush() {
  if (!batch.length) return;
  sheet.getRangeByIndexes(written + 1, 0, batch.length, columns.length).values = batch;
  written += batch.length;
  batch = [];
  if (written % 10000 === 0) console.log(`Written ${written}/${n}`);
}
for await (const line of lines) {
  const row = JSON.parse(line);
  batch.push(columns.map(k => {
    const v = row[k] ?? null;
    return typeof v === 'string' && v.startsWith('=') ? `'${v}` : v;
  }));
  if (batch.length === 1000) flush();
}
flush();
if (written !== n) throw new Error(`Count mismatch: ${written} vs ${n}`);
console.log(`Data complete: ${written} rows, ${columns.length} columns`);

const header = sheet.getRangeByIndexes(0, 0, 1, columns.length);
header.format = { fill: '#284B63', font: { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' }, wrapText: true, horizontalAlignment: 'center', verticalAlignment: 'center' };
header.format.rowHeight = 56;
sheet.freezePanes.freezeRows(1);
const last = (num => { let s = ''; while(num) { const m = (num-1)%26; s=String.fromCharCode(65+m)+s;num=Math.floor((num-1)/26); } return s; })(columns.length);
console.log('Adding filter table');
const table = sheet.tables.add(`A1:${last}${n+1}`, true, 'HoaDonGop');
table.showFilterButton = true;
table.style = 'TableStyleMedium2';
console.log('Final calculation');
workbook.recalculate();
const checked = await workbook.inspect({ kind: 'table', range: 'Hóa đơn gộp!A1:H5', include: 'values', tableMaxRows: 5, tableMaxCols: 8, maxChars: 1800 });
await fs.writeFile(`${dir}/artifact_check.ndjson`, checked.ndjson);
console.log('Exporting workbook');
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${dir}/ĐỐI SOÁT ĐÃ GỘP.xlsx`);
console.log(`SAVED ${dir}/ĐỐI SOÁT ĐÃ GỘP.xlsx`);
