import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const p='D:/UVG/outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx';
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(p));
console.log((await w.inspect({kind:'sheet',include:'id,name',maxChars:1000})).ndjson);
console.log(w.help('table.resize',{include:'index,examples,notes',maxChars:2400}).ndjson);
const b=await w.render({sheetName:'Dieu chinh',range:'S1:AB5',scale:1,format:'png'});
await fs.writeFile('D:/UVG/scripts/data/check-review/before-plan.png',new Uint8Array(await b.arrayBuffer()));
