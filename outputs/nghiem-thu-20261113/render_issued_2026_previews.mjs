import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook} from '@oai/artifact-tool';
const out=path.dirname(fileURLToPath(import.meta.url));
const chains=process.argv.includes('--chains');const compact=process.argv.includes('--compact')||chains;
const data=JSON.parse(await fs.readFile(path.join(out,chains?'issued-2026-chain-preview-cells.json':compact?'issued-2026-compact-preview-cells.json':'issued-2026-preview-cells.json'),'utf8'));
const wb=Workbook.create();
for(const view of data){
 const sh=wb.worksheets.add(view.name);sh.showGridLines=false;
 for(let c=0;c<view.widths.length;c++)sh.getCell(0,c).format.columnWidth=view.widths[c];
 for(let r=0;r<view.rows.length;r++){
  sh.getCell(r,0).format.rowHeight=view.heights[r];
  for(let c=0;c<view.rows[r].length;c++){
   const obj=view.rows[r][c];const cell=sh.getCell(r,c);
   cell.values=[[obj.isDate?new Date(obj.value):obj.value??null]];
   if(!obj.font)continue;
   cell.format.font=obj.font;
   cell.format.verticalAlignment='center';
   if(obj.fill)cell.format.fill=obj.fill;
   if(obj.borders && Object.keys(obj.borders).length)cell.format.borders=obj.borders;
   cell.format.wrapText=obj.wrap;
   if(['left','center','right'].includes(obj.align))cell.format.horizontalAlignment=obj.align;
   cell.setNumberFormat(obj.format);
   // Renderer tự đọc chuỗi số như số; định dạng này tái hiện đúng 8 ký tự Excel thực đang lưu dạng text.
   if(view.name==='HĐ 2026' && view.source_columns[c]===4 && typeof obj.value==='string' && /^\d{8}$/.test(obj.value))cell.setNumberFormat('00000000');
   if(typeof obj.value==='string' && /^\d{15,}$/.test(obj.value))cell.setNumberFormat('"'+obj.value+'"');
   if(chains && view.source_columns[c]===3 && view.source_rows[r]>=7 && (obj.value===null||obj.value===''))cell.format.fill='#FFF1D6';
  }
 }
 for(const address of view.merges||[])sh.mergeCells(address);
}
wb.recalculate();
for(const view of data){
 const p=await wb.render({sheetName:view.name,autoCrop:'all',scale:1,format:'png'});
 await fs.writeFile(path.join(out,(chains?'chain-final-preview-':compact?'compact-final-preview-':'final-issued-preview-')+view.name+'.png'),new Uint8Array(await p.arrayBuffer()));
 console.log('RENDERED actual workbook',view.name);
}
