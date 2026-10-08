import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const dir='D:/UVG/scripts/data/check-review';
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${dir}/before-rename.xlsx`));
const s=w.worksheets.getItem('Dieu chinh');
let b=await w.render({sheetName:'Dieu chinh',range:'G1:L3',scale:1,format:'png'});
await fs.writeFile(`${dir}/rename-before.png`,new Uint8Array(await b.arrayBuffer()));
// Chỉ đổi tiêu đề đang chứa từ sửa, áp dụng cho mọi khối F.
const old=s.getRange('A2:BC2').values[0];const next=old.map(v=>typeof v==='string'?v.replace(/\bsửa\s+/g,''):v);
s.getRange('A2:BC2').values=[next];
s.tables.items.find(t=>t.name==='DieuChinhChuoi').delete();
const t=s.tables.add('A2:BC1049',true,'DieuChinhChuoi');t.style='TableStyleMedium2';t.showFilterButton=true;
w.recalculate();
b=await w.render({sheetName:'Dieu chinh',range:'G1:L3',scale:1,format:'png'});
await fs.writeFile(`${dir}/rename-after.png`,new Uint8Array(await b.arrayBuffer()));
await (await SpreadsheetFile.exportXlsx(w)).save(`${dir}/renamed.xlsx`);
console.log(JSON.stringify({changed:old.filter((v,i)=>v!==next[i]).length,headers:next.slice(6,24)}));
