import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const h='D:/UVG/scripts/data/check-review';
const w=await SpreadsheetFile.importXlsx(await FileBlob.load('D:/UVG/outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx'));
if(process.argv.includes('--inspect')){
 const b=await w.render({sheetName:'Dieu chinh',range:'G902:L904',scale:1,format:'png'});
 await fs.writeFile(`${h}/date-12885-before.png`,new Uint8Array(await b.arrayBuffer()));
 console.log('Rendered before');
}else{
 const d=JSON.parse(await fs.readFile(`${h}/date-12885-confirmed.json`,'utf8'));
 const s=w.worksheets.getItem('Dieu chinh');
 s.getRange('I903').values=[[new Date(Date.UTC(2025,7,18))]];
 s.getRange('AB903').values=[[d.AB903]];
 w.worksheets.getItem('Chi tiet nguon').getRange('M1822').values=[[d.M1822]];
 w.recalculate();
 console.log((await w.inspect({kind:'table',range:'Dieu chinh!G903:L903',include:'values',tableMaxRows:1,tableMaxCols:6,maxChars:1200})).ndjson);
 const b=await w.render({sheetName:'Dieu chinh',range:'G902:L904',scale:1,format:'png'});
 await fs.writeFile(`${h}/date-12885-after.png`,new Uint8Array(await b.arrayBuffer()));
 await (await SpreadsheetFile.exportXlsx(w)).save(`${h}/date-12885-updated.xlsx`);
 console.log('Staged date edit');
}