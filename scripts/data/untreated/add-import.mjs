import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const h='D:/UVG/scripts/data/untreated';
const excludedKeys=new Set(['1C24TUV/00000774','1C24TUV/00000783','1C24TUV/00000784','1C24TUV/00000786','1C25TUV/00002796','1C25TUV/00002797']);
if(process.argv.includes('--exclude-zero')){
 console.log('READ CURRENT');const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${h}/before-zero-exclusion.xlsx`));
 const s=w.worksheets.getItem('Import dieu chinh'),review=w.worksheets.add('Kiem tra SL 0');
 const old=s.getRange('A10:X13579').values,keep=[],removed=[],seen=new Map(),offset=new Map();
 const d=JSON.parse(await fs.readFile(`${h}/import-data.json`,'utf8'));const src=new Map(d.roots.map(x=>[x.key.join('/'),x]));
 for(let i=0;i<old.length;i++){
  const v=old[i],key=v[9]+'/'+v[10];
  if(excludedKeys.has(key)){
   const index=offset.get(key)||0;const l=src.get(key).lines[index];offset.set(key,index+1);
   const first=index===0;
   const note='Đã tách toàn bộ hóa đơn khỏi import theo yêu cầu người dùng ngày 08/10/2026 vì có dòng số lượng gốc = 0.'+(v[16]===0?' Dòng này: số lượng, đơn giá, thành tiền gốc = 0.':' Dòng này đi kèm trong cùng hóa đơn cần kiểm tra.');
   removed.push([v[9],v[10],v[11],v[2],v[3],v[5],v[7],v[14],v[15],v[16]===null?null:-v[16],v[17],v[18]===null?null:-v[18],first?-v[19]:null,first?v[20]:null,first?-v[21]:null,first?-v[22]:null,note,l.source+' | dòng '+l.row,i+10]);
  }else{if(!seen.has(key))seen.set(key,seen.size+1);const copy=[...v];copy[0]=seen.get(key);keep.push(copy);}
 }
 if(removed.length!==27||offset.size!==6||seen.size!==8404||keep.length!==13543)throw new Error('Unexpected exclusion counts');
 console.log('WRITE');for(let i=0;i<keep.length;i+=1000){const part=keep.slice(i,i+1000);s.getRange(`A${i+10}:X${i+9+part.length}`).values=part;}
 s.getRange('A13553:X13579').clear({applyTo:'contents'});
 s.getRange('A7').values=[['- Thành tiền dòng gốc = 0: Hàng khuyến mãi = 1. Đã tách 6 hóa đơn có dòng số lượng gốc 0 sang sheet Kiem tra SL 0; import còn 8.404 hóa đơn, 13.543 dòng.']];
 review.mergeCells('A1:S1');review.getRange('A1').values=[['6 hóa đơn có dòng số lượng gốc = 0 — đã loại toàn bộ khỏi import, chờ kiểm tra']];
 const headers=['Ký hiệu hóa đơn gốc','Số hóa đơn gốc','Ngày hóa đơn gốc','Tên khách hàng','Địa chỉ','Người mua hàng','Hình thức thanh toán','Tên hàng hóa/dịch vụ','ĐVT','Số lượng gốc','Đơn giá gốc','Thành tiền gốc','Tổng tiền hàng gốc','Thuế suất GTGT (%)','Tổng thuế gốc','Tổng thanh toán gốc','Ghi chú kiểm tra','Nguồn chi tiết','Dòng import trước khi tách'];
 review.getRange('A2:S2').values=[headers];review.getRange('A3:S29').values=removed;
 review.showGridLines=false;review.getRange('A1:S29').format={font:{name:'Arial',size:10},verticalAlignment:'center',rowHeight:65};
 review.getRange('A1:S2').format={fill:'#243746',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true};review.getRange('A1').format.rowHeight=32;review.getRange('A2:S2').format.rowHeight=48;
 const widths=[19,20,21,48,55,48,23,100,12,17,23,24,24,18,24,26,105,95,22];widths.forEach((v,i)=>review.getRangeByIndexes(0,i,1,1).format.columnWidth=v);
 review.getRange('D3:H29').format.wrapText=true;review.getRange('Q3:R29').format.wrapText=true;review.getRange('B3:B29').setNumberFormat('00000000');review.getRange('J3:P29').setNumberFormat('#,##0;[Red](#,##0);0');
 for(let i=0;i<removed.length;i++)if(removed[i][9]===0)review.getRange(`A${i+3}:S${i+3}`).format.fill='#FFF0C2';
 const t=review.tables.add('A2:S29',true,'KiemTraSoLuong0');t.style='TableStyleMedium2';t.showFilterButton=true;review.freezePanes.freezeRows(2);review.freezePanes.freezeColumns(2);
 const g=w.worksheets.getItem('Huong dan');g.getRange('A13').values=[['Sheet Import dieu chinh: còn 8.404 hóa đơn điều chỉnh giảm với 13.543 dòng theo mẫu người dùng; đã loại 6 hóa đơn có dòng số lượng gốc 0.']];
 g.getRange('A16').values=[['Sheet Kiem tra SL 0 giữ đủ 27 dòng chi tiết gốc của 6 hóa đơn đã tách; tổng tiền gốc chỉ ghi ở dòng đầu mỗi hóa đơn, không sửa nguồn. Import đã đánh lại số thứ tự liên tục.']];g.getRange('A16').format={font:{name:'Arial',size:10},wrapText:true,rowHeight:48,columnWidth:145};
 w.recalculate();console.log('PREVIEW');for(const [sheetName,range,name] of [['Kiem tra SL 0','H1:P5','zero-review-values.png'],['Kiem tra SL 0','Q1:S5','zero-review-notes.png'],['Import dieu chinh','A9:C13','zero-excluded-import.png']]){const b=await w.render({sheetName,range,scale:1,format:'png'});await fs.writeFile(`${h}/${name}`,new Uint8Array(await b.arrayBuffer()));}
 console.log('EXPORT');await(await SpreadsheetFile.exportXlsx(w)).save(`${h}/with-zero-exclusion.xlsx`);
 await fs.writeFile(`${h}/zero-exclusion-meta.json`,JSON.stringify({keys:[...excludedKeys],keep,removed}));console.log(JSON.stringify({invoices:seen.size,keptLines:keep.length,removedLines:removed.length,staged:true}));process.exit(0);
}

if(process.argv.includes('--sync-buyer')){
 console.log('READ CURRENT');
 const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${h}/before-buyer-sync.xlsx`));
 const s=w.worksheets.getItem('Import dieu chinh');
 if(s.getRange('C9').values[0][0]!=='Tên khách hàng'||s.getRange('F9').values[0][0]!=='Người mua hàng')throw new Error('Unexpected import headers');
 const buyers=s.getRange('F10:F13579').values;
 s.getRange('C10:C13579').values=buyers;s.getRange('C10:C13579').format.wrapText=true;
 w.recalculate();console.log('PREVIEW');
 const b=await w.render({sheetName:'Import dieu chinh',range:'A9:I12',scale:1,format:'png'});await fs.writeFile(`${h}/buyer-sync.png`,new Uint8Array(await b.arrayBuffer()));
 console.log('EXPORT');await(await SpreadsheetFile.exportXlsx(w)).save(`${h}/with-buyer-sync.xlsx`);
 console.log(JSON.stringify({synced:buyers.length,staged:true}));process.exit(0);
}

if(process.argv.includes('--set-date')){
 const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${h}/with-import.xlsx`));
 const s=w.worksheets.getItem('Import dieu chinh');
 s.getRange('B10:B13579').values=[[new Date(Date.UTC(2026,9,8))]];s.getRange('B10:B13579').setNumberFormat('dd/mm/yyyy');
 s.getRange('A3').values=[['- Ngày hóa đơn mới hiện tại là 08/10/2026; có thể sửa đồng nhất các dòng cùng số thứ tự hóa đơn trước khi import.']];
 w.worksheets.getItem('Huong dan').getRange('A14').values=[['Ngày hóa đơn mới 08/10/2026 tại B; có thể sửa đồng nhất trước khi import. Tổng tiền chỉ ghi ở dòng đầu; 6 F0 đã chốt thanh toán 0.']];
 w.recalculate();
 const b=await w.render({sheetName:'Import dieu chinh',range:'A9:I12',scale:1,format:'png'});await fs.writeFile(`${h}/import-buyer.png`,new Uint8Array(await b.arrayBuffer()));
 await(await SpreadsheetFile.exportXlsx(w)).save(`${h}/with-import.xlsx`);
 console.log('DATE UPDATED 08/10/2026');process.exit(0);
}

const d=JSON.parse(await fs.readFile(`${h}/import-data.json`,'utf8'));
console.log('IMPORT BASE');const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${h}/before-import.xlsx`));console.log('BASE READY');
const s=w.worksheets.add('Import dieu chinh');
const headers=['Số thứ tự hóa đơn (*)','Ngày hóa đơn','Tên khách hàng','Địa chỉ','Mã số thuế','Người mua hàng','Email','Hình thức thanh toán (*)','HĐ bị điều chỉnh thuộc hệ thống khác','Ký hiệu HĐ bị điều chỉnh (*)','Số hóa đơn bị điều chỉnh (*)','Ngày hóa đơn bị điều chỉnh','Mã của CQT trên HĐ bị điều chỉnh','Lý do điều chỉnh','Tên hàng hóa/dịch vụ (*)','ĐVT','Số lượng','Đơn giá','Thành tiền','Tổng tiền hàng','Thuế suất GTGT (%)','Tổng tiền thuế GTGT','Tổng tiền thanh toán','Hàng khuyến mãi'];
const instructions=['File mẫu danh sách hóa đơn để nhập vào phần mềm','Hướng dẫn:','- Ngày hóa đơn mới hiện tại là 08/10/2026; có thể sửa đồng nhất các dòng cùng số thứ tự hóa đơn trước khi import.','- Các cột có dấu (*) là những cột bắt buộc.','- Tổng tiền hàng, thuế suất, tổng thuế và tổng thanh toán chỉ ghi ở dòng đầu mỗi số thứ tự hóa đơn; số tổng lấy từ bảng tổng đã chốt.','- Mỗi số thứ tự tương ứng một hóa đơn điều chỉnh giảm, tham chiếu đúng F0; giữ từng dòng hàng, đảo dấu số lượng/thành tiền.','- Thành tiền dòng gốc = 0: Hàng khuyến mãi = 1. Dữ liệu bên dưới là dữ liệu chuẩn bị thực tế.'];
for(let i=0;i<instructions.length;i++){s.mergeCells(`A${i+1}:X${i+1}`);s.getRange(`A${i+1}`).values=[[instructions[i]]];}
s.getRange('A9:X9').values=[headers];
const rows=[],meta=[];let seq=0;
const neg=v=>v===null||v===''?null:-Number(v);
const text=v=>v===null||v===''?null:String(v);
const rate=v=>v===null||v===''?null:/^\d+(\.\d+)?%?$/.test(String(v))?Number(String(v).replace('%','')):v;
for(const x of d.roots){if(excludedKeys.has(x.key.join('/')))continue;seq++;let first=true;
 for(const l of x.lines){const v=l.v,m=l.mt;const taxrate=v[m?25:18];
  const row=[seq,new Date(Date.UTC(2026,9,8)),text(v[m?9:7]),text(v[m?5:5]),text(v[m?8:6]),text(v[m?9:7]),m?text(v[10]):null,text(v[m?13:8]),null,x.key[0],x.key[1],v[m?3:2],null,`Điều chỉnh giảm hàng bán HĐ số ${x.key[1]} ngày ${v[m?3:2]} do xuất sai thông tin`,text(v[m?18:11]),text(v[m?19:12]),neg(v[m?20:13]),v[m?21:14],neg(v[m?22:15]),first?neg(x.values[3]):null,first?rate(taxrate):null,first?neg(x.values[4]):null,first?neg(x.values[5]):null,v[m?22:15]===0?1:null];
  if(!row[7]||!row[14])throw new Error(`Missing required field ${x.key}`);
  rows.push(row);meta.push({row:rows.length+9,key:x.key,source:l.source,source_row:l.row,first});first=false;
 }
}
console.log('WRITE DATA',rows.length);for(let start=0;start<rows.length;start+=1000){const part=rows.slice(start,start+1000);s.getRange(`A${start+10}:X${start+9+part.length}`).values=part;}console.log('FORMAT');
s.showGridLines=false;s.getRange('A1:X9').format.font={name:'Arial',size:10};
s.getRange('A1:X7').format={wrapText:true,rowHeight:26};s.getRange('A1').format.font.bold=true;
s.getRange('A9:X9').format={fill:'#243746',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:54};
s.getRange(`A10:X${rows.length+9}`).format.rowHeight=62;
const widths=[14,16,42,52,23,42,30,24,18,21,24,20,25,85,100,12,18,22,24,24,19,24,26,18];
widths.forEach((v,i)=>s.getRangeByIndexes(0,i,1,1).format.columnWidth=v);
s.getRange(`N10:O${rows.length+9}`).format.wrapText=true;
s.getRange(`B10:B${rows.length+9}`).setNumberFormat('dd/mm/yyyy');s.getRange(`E10:E${rows.length+9}`).setNumberFormat('@');s.getRange(`K10:K${rows.length+9}`).setNumberFormat('00000000');
s.getRange(`Q10:W${rows.length+9}`).setNumberFormat('#,##0.########;[Red](#,##0.########);0');s.getRange(`A10:A${rows.length+9}`).setNumberFormat('0');s.getRange(`X10:X${rows.length+9}`).setNumberFormat('0');
s.freezePanes.freezeRows(9);s.freezePanes.freezeColumns(1);
const g=w.worksheets.getItem('Huong dan');g.getRange('A13:A15').values=[['Sheet Import dieu chinh: 8.410 hóa đơn điều chỉnh giảm với 13.570 dòng theo mẫu người dùng; các tổng tiền đã đối chiếu khớp chi tiết.'],['Ngày hóa đơn mới 08/10/2026 tại B; có thể sửa đồng nhất trước khi import. Tổng tiền chỉ ghi ở dòng đầu; 6 F0 đã chốt thanh toán 0.'],['4 F0 1C24TUV/00000774, 00000783, 00000784, 00000786 có dòng thành tiền 0/thuế suất trống; giữ nguyên nguồn, Hàng khuyến mãi = 1. Chưa thực hiện import hoặc phát hành.']];g.getRange('A13:A15').format={font:{name:'Arial',size:10},wrapText:true,rowHeight:48,columnWidth:145};
console.log('RECALC');w.recalculate();console.log('PREVIEW');
for(const [range,file] of [['J9:X13','import-money.png'],['A9:I12','import-buyer.png']]){const b=await w.render({sheetName:'Import dieu chinh',range,scale:1,format:'png'});await fs.writeFile(`${h}/${file}`,new Uint8Array(await b.arrayBuffer()));}
console.log('EXPORT');await(await SpreadsheetFile.exportXlsx(w)).save(`${h}/with-import.xlsx`);
await fs.writeFile(`${h}/import-meta.json`,JSON.stringify({rows,meta}));console.log(JSON.stringify({invoices:seq,lines:rows.length,staged:true}));
