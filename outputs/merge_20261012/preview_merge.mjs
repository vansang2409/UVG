import fs from 'node:fs/promises';
import { Workbook, FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const dir='D:/UVG/outputs/merge_20261012';
console.log('Importing saved workbook for visual verification');
const workbook=await SpreadsheetFile.importXlsx(await FileBlob.load(`${dir}/qa_snapshot.xlsx`));
for(const [file,range] of [['preview_left.png','A1:H7'],['preview_dates.png','I1:S7'],['preview_products.png','T1:AA7'],['preview_source.png','AR1:AX7']]){
  console.log(`Rendering ${range}`);
  const preview=await workbook.render({sheetName:'Hóa đơn gộp',range,scale:1.3,format:'png'});
  await fs.writeFile(`${dir}/${file}`,new Uint8Array(await preview.arrayBuffer()));
}
console.log('Visual previews saved');
