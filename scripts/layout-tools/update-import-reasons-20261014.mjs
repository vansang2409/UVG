import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';
const root = path.resolve('scripts/data/import-reasons-20261014');
const plan = JSON.parse(await fs.readFile(path.join(root, 'plan.json'), 'utf8'));
const index = Number(process.argv[2]);
const f = plan.files[index];
await fs.mkdir(path.join(root, 'staged'), { recursive: true });
// Chỉ dựng các ô được phép thay đổi; giữ toàn bộ package gốc ở bước ghép sau.
const w = Workbook.create();
const s = w.worksheets.add('Import dieu chinh');
for (let i = 0; i < f.changes.length; i += 1000) {
  const block = f.changes.slice(i, i + 1000);
  s.getRange(`N${block[0].row}:N${block.at(-1).row}`).values = block.map(c => [c.new]);
}
console.log('EDITED', f.changes.length);
s.getRange('N10:N11').format = {columnWidth: 100, rowHeight: 50, wrapText: true};
const preview = await w.render({sheetName: 'Import dieu chinh', range: 'N10:N11', scale: 1, format: 'png'});
await fs.writeFile(path.join(root, `preview-${index}.png`), new Uint8Array(await preview.arrayBuffer()));
await (await SpreadsheetFile.exportXlsx(w)).save(path.join(root, `authored-${index}.xlsx`));
console.log('EXPORTED', f.name);
