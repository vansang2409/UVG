import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const work='D:/UVG/scripts/data/amis-2024-linked-20261026';
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${work}/_template-preview.tmp.xlsx`));
console.log((await w.inspect({kind:'region',sheetId:'Hóa đơn GTGT',range:'A9:S11',maxChars:1800,tableMaxRows:3,tableMaxCols:19})).ndjson);
for(const [range,name] of [['A1:K11','template-left'],['L8:S12','template-right']]){
 const png=await w.render({sheetName:'Hóa đơn GTGT',range,scale:1,format:'png'});
 await fs.writeFile(`${work}/${name}.png`,new Uint8Array(await png.arrayBuffer()));
}
await fs.unlink(`${work}/_template-preview.tmp.xlsx`);
