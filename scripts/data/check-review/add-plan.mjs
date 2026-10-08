import fs from 'node:fs/promises';
import path from 'node:path';
import {SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const root='D:/UVG', h=path.join(root,'scripts/data/check-review');
const d=JSON.parse(await fs.readFile(path.join(h,'plan-data.json'),'utf8'));
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(h,'before-plan.xlsx')));
const s=w.worksheets.getItem('Dieu chinh'), last=d.items.length+2;
function col(i){let r='';for(i++;i;i=Math.floor((i-1)/26))r=String.fromCharCode(65+(i-1)%26)+r;return r;}
const hdr=[], matrix=[], formulas=[];
for(let j=4;j<=7;j++)hdr.push(`Nội dung F${j} / HĐ tham chiếu`,`Số HĐ mới F${j}`,`Ngày HĐ mới F${j}`,`Tiền hàng F${j}`,`Thuế F${j}`,`Thanh toán F${j}`);
hdr.push('Cộng dồn tiền hàng sau dự kiến','Cộng dồn thuế sau dự kiến','Cộng dồn thanh toán sau dự kiến');
for(const x of d.items){
 const row=[], operations=[...x.old.slice(1).map(v=>({kind:'Đảo khoản điều chỉnh',v})),{kind:'Giảm toàn bộ gốc',v:x.old[0]}];
 const pending=Boolean(x.note)||x.root.join('/')==='1C24TUV/00000003';
 for(let j=0;j<4;j++){
  const op=operations[j];
  if(!op){row.push(null,null,null,null,null,null);continue;}
  const key=op.v.slice(0,2).join('/'), source=d.totals[key];
  if(JSON.stringify(op.v.slice(3,6))!==JSON.stringify(source.money))throw new Error('Report/source difference '+key);
  let label=op.kind.startsWith('Đảo') ? `Đảo khoản ${key}; tham chiếu gốc ${x.root.join('/')}` : `Giảm toàn bộ gốc: ${key}`;
  if(pending)label+=' (chờ kiểm tra)';
  else if(x.prior.every(v=>v===0))label+=' (chuỗi đã 0; rà trước khi tạo)';
  row.push(label,null,null,...source.money.map(v=>v===null?null:-v));
 }
 row.push(null,null,null);matrix.push(row);
 formulas.push([0,1,2].map(k=>{
  const cells=[0,1,2,3].map(j=>`${col(28+j*6+3+k)}${x.row}`);
  const prior=`${col(24+k)}${x.row}`;
  return `=IF(AND(ISNUMBER(${prior}),COUNT(${cells.join(',')})=${operations.length}),${prior}+SUM(${cells.join(',')}),"")`;
 }));
}
s.getRange(`AC2:BC2`).values=[hdr];s.getRange(`AC3:BC${last}`).values=matrix;
s.getRange(`BA3:BC${last}`).formulas=formulas;
// Mở rộng cùng bảng để bộ lọc bao gồm các cột mới.
s.tables.items.find(t=>t.name==='DieuChinhChuoi').delete();
const table=s.tables.add(`A2:BC${last}`,true,'DieuChinhChuoi');table.style='TableStyleMedium2';table.showFilterButton=true;
s.getRange(`AC1:BC${last}`).format.font={name:'Arial',size:10};
s.getRange(`AC1:BC${last}`).format.verticalAlignment='center';
s.getRange('AC1:BC2').format={fill:'#243746',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center',wrapText:true};
for(let j=0;j<4;j++){
 const c=28+j*6, range=`${col(c)}1:${col(c+5)}1`;
 s.mergeCells(range);s.getRange(`${col(c)}1`).values=[[`F${j+4} — Hóa đơn dự kiến tạo`]];
 s.getRangeByIndexes(0,c,2,6).format.fill=['#546C54','#657F65','#546C54','#657F65'][j];
 [55,18,20,22,20,24].forEach((v,k)=>s.getRangeByIndexes(0,c+k,last,1).format.columnWidth=v);
 s.getRangeByIndexes(2,c,d.items.length,1).format.wrapText=true;
 s.getRangeByIndexes(2,c+1,d.items.length,1).setNumberFormat('@');
 s.getRangeByIndexes(2,c+2,d.items.length,1).setNumberFormat('dd/mm/yyyy');
 s.getRangeByIndexes(2,c+3,d.items.length,3).setNumberFormat('#,##0;[Red](#,##0);0');
 s.getRangeByIndexes(2,c+3,d.items.length,3).format.horizontalAlignment='right';
}
s.mergeCells('BA1:BC1');s.getRange('BA1').values=[['Cộng dồn sau F4–F7 (dự kiến)']];
s.getRange(`BA3:BC${last}`).setNumberFormat('#,##0;[Red](#,##0);0');
s.getRange(`BA3:BC${last}`).format.horizontalAlignment='right';
s.getRange(`BA1:BC${last}`).format.columnWidth=29;
for(const x of d.items)s.getRange(`A${x.row}:BC${x.row}`).format.rowHeight=Math.max(48,s.getRange(`A${x.row}:AB${x.row}`).format.rowHeight||28);
s.freezePanes.freezeRows(2);s.freezePanes.freezeColumns(2);
w.recalculate();
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#SPILL!',options:{useRegex:true,maxResults:20},maxChars:800})).ndjson);
for(const [range,file] of [['AC1:AN5','plan-new-1.png'],['AO1:BC5','plan-new-2.png'],['AO733:BC735','plan-four-invoices.png']]){
 const b=await w.render({sheetName:'Dieu chinh',range,scale:1,format:'png'});await fs.writeFile(path.join(h,file),new Uint8Array(await b.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(w)).save(path.join(h,'with-plan.xlsx'));
console.log(JSON.stringify({staged:true,chains:d.items.length,planned:d.items.reduce((a,x)=>a+x.old.length,0)}));

