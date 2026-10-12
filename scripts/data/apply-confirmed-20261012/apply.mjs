import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const here=path.dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/,'$1'));
const d=JSON.parse(await fs.readFile(path.join(here,'edit-plan.json'),'utf8'));
const inspectOnly=process.argv.includes('--inspect');
const reviewFirst=process.argv.includes('--review');
const repairFields=process.argv.includes('--repair-fields');
const filter=process.argv.find(v=>v.startsWith('--kind='))?.split('=')[1];
const col=n=>{let x='';for(n++;n;n=Math.floor((n-1)/26))x=String.fromCharCode(65+(n-1)%26)+x;return x;};
const day=x=>x?new Date(`${x.slice(0,10)}T00:00:00Z`):null;
const money='#,##0;[Red](#,##0);0';
await fs.mkdir(path.join(here,'previews'),{recursive:true});
await fs.mkdir(path.join(here,'staged'),{recursive:true});
function put(s,cell,value){s.getRange(cell).values=[[value]];}
function replaceTable(s,end){const t=s.tables.items[0];if(!t)return;const name=t.name,style=t.style;t.delete();const nt=s.tables.add(end,true,name);if(style)nt.style=style;nt.showFilterButton=true;}
function wrap(s,range,height){s.getRange(range).format={font:{name:'Arial',size:10},wrapText:true,rowHeight:height,verticalAlignment:'center'};}
async function render(w,f,when){
  const selection=f.kind==='untreated'?[['Import dieu chinh',when==='before'?'J9:X12':`J${f.append_start}:X${f.append_start+3}`,'import'],['Kiem tra SL 0','H1:R5','status']]:f.kind==='adjusted'?[['Import dieu chinh',when==='before'?'J9:X12':`J${f.append_start}:X${f.append_start+3}`,'import'],['Cho KT xac nhan','G1:I5','status'],['Dieu chinh','AC698:BC699','plan']]:[['Ngoai le',when==='before'?'A237:X240':'A241:X244','exceptions']];
  for(const [sheetName,range,label] of selection){const b=await w.render({sheetName,range,scale:1,format:'png'});await fs.writeFile(path.join(here,'previews',`${f.kind}-${when}-${label}.png`),new Uint8Array(await b.arrayBuffer()));}
}
for(const initial of d.files){
 let f=initial;
 if(filter&&filter!==f.kind)continue;
 console.log('LOAD',f.kind);
 const w=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(here,repairFields?'staged':'before',f.name)));
 if(inspectOnly){const a=await w.inspect({kind:'workbook,sheet,table',maxChars:1800,tableMaxRows:2,tableMaxCols:3});await fs.writeFile(path.join(here,`${f.kind}-before.ndjson`),a.ndjson);await render(w,f,'before');continue;}
 if(reviewFirst){
  const a=await w.inspect({kind:'workbook,sheet,table',maxChars:1800,tableMaxRows:2,tableMaxCols:3});await fs.writeFile(path.join(here,`${f.kind}-before.ndjson`),a.ndjson);
  if(repairFields){const b=await w.render({sheetName:'Import dieu chinh',range:`A${f.append_start}:I${f.append_start+2}`,scale:1,format:'png'});await fs.writeFile(path.join(here,'previews',`${f.kind}-fields-before.png`),new Uint8Array(await b.arrayBuffer()));}else await render(w,f,'before');
  console.log('READY_FOR_VISUAL_REVIEW',f.kind);
  const flag=path.join(here,`${f.kind}${repairFields?'-fields':''}-reviewed.flag`);
  while(true){try{await fs.access(flag);break;}catch{await new Promise(resolve=>setTimeout(resolve,1000));}}
 }
 // Resolve the final prepared input only at the edit boundary, after source review.
 const latest=JSON.parse(await fs.readFile(path.join(here,'edit-plan.json'),'utf8'));Object.assign(d,latest);f=latest.files.find(x=>x.kind===f.kind);
 if(repairFields){
  assert(f.kind==='adjusted');const s=w.worksheets.getItem('Import dieu chinh');
  s.getRange(`G${f.append_start}:H${9+f.rows}`).values=f.new_rows.map(r=>[r[6]||null,r[7]||null]);
  w.recalculate();
  const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:50}});await fs.writeFile(path.join(here,`${f.kind}-errors.ndjson`),errors.ndjson);
  for(const [sheetName,range,label] of [['Import dieu chinh',`A${f.append_start}:I${f.append_start+2}`,'fields-after'],['Huong dan','A11:A16','guide'],['42 - Can check 11','V3:AA7','review-status']]){const b=await w.render({sheetName,range,scale:1,format:'png'});await fs.writeFile(path.join(here,'previews',`${f.kind}-${label}.png`),new Uint8Array(await b.arrayBuffer()));}
  await(await SpreadsheetFile.exportXlsx(w)).save(path.join(here,'staged',f.name));console.log('DONE_FIELDS_REPAIR',f.kind);continue;
 }
 console.log('EDIT',f.kind);
 if(f.kind!=='exceptions'){
  const s=w.worksheets.getItem('Import dieu chinh');
  const rows=f.new_rows.map(r=>r.map((v,i)=>i===1?day(v):v));
  for(let i=0;i<rows.length;i+=500){const block=rows.slice(i,i+500);s.getRange(`A${f.append_start+i}:X${f.append_start+i+block.length-1}`).values=block;}
  const end=9+f.rows;
  s.getRange(`A${f.append_start}:X${end}`).format={font:{name:'Arial',size:10},rowHeight:62,verticalAlignment:'center'};
  s.getRange(`C${f.append_start}:H${end}`).format.wrapText=true;
  s.getRange(`N${f.append_start}:O${end}`).format.wrapText=true;
  s.getRange(`B${f.append_start}:B${end}`).setNumberFormat('dd/mm/yyyy');
  s.getRange(`E${f.append_start}:E${end}`).setNumberFormat('@');
  s.getRange(`K${f.append_start}:K${end}`).setNumberFormat('00000000');
  s.getRange(`Q${f.append_start}:W${end}`).setNumberFormat(money);
  for(const a of f.new_actions)if(a.difference!==0)s.getRange(`T${a.start}:W${a.start}`).format.fill='#FFF0C2';
  put(s,'A3','- Phần bổ sung theo chốt kế toán dùng ngày 12/10/2026; phần đã chuẩn bị trước giữ ngày 08/10 hoặc 09/10/2026. Kiểm tra ngày dự thảo trước khi nhập.');
  const g=w.worksheets.getItem('Huong dan');
  if(f.kind==='untreated'){
   put(s,'A7',`- Đã bổ sung 6 hóa đơn theo chốt kế toán; tổng ${f.invoices.toLocaleString('vi-VN')} hóa đơn / ${f.rows.toLocaleString('vi-VN')} dòng. Dòng chú thích giữ số lượng/đơn giá/thành tiền trống; bỏ đúng dòng 0 của 2796 và 2797.`);
   put(g,'A11','Đã cập nhật dữ liệu chuẩn bị import theo mẫu; chưa nhập, ký hoặc phát hành. Nguồn gốc giữ nguyên.');
   put(g,'A13',`Import dieu chinh: ${f.invoices.toLocaleString('vi-VN')} hóa đơn giảm / ${f.rows.toLocaleString('vi-VN')} dòng, gồm 6 hóa đơn vừa được kế toán chốt.`);
   put(g,'A14','Phần trước giữ ngày 08/10/2026; 6 hóa đơn bổ sung ngày 12/10/2026. Tổng tiền ghi dòng đầu mỗi hóa đơn; ngày có thể sửa đồng nhất trước khi nhập.');
   put(g,'A15','774/783/784/786: giữ dòng chú thích hàng biếu/tặng, để trống số lượng/đơn giá/thành tiền. Mẫu 24 cột chưa có trường tính chất dòng; kiểm tra/chọn tính chất diễn giải trên bản nháp nếu phần mềm không tự nhận.');
   put(g,'A16','Kiem tra SL 0 giữ dữ liệu lịch sử 27 dòng đã tách. Cột ghi chú nay ghi phương án và STT import đã bổ sung; không dùng số 0 lịch sử thay ô trống của bảng xuất lại.');
   const review=w.worksheets.getItem('Kiem tra SL 0');
   put(review,'A1','6 hóa đơn đã chốt và bổ sung vào import — giữ dữ liệu lịch sử đã tách để đối chiếu');
   for(let r=3;r<=29;r++){
    const k=review.getRange(`A${r}:B${r}`).values[0].join('/'),a=f.new_actions.find(a=>a.root.join('/')===k);
    assert(a,k);put(review,`Q${r}`,`Đã chốt; đã bổ sung STT ${a.seq}, dòng ${a.start}–${a.end} tại Import dieu chinh. ${d.decisions[a.case].decision} Dữ liệu số trong sheet này là bản lịch sử trước khi xuất lại.`);
    review.getRange(`A${r}:S${r}`).format.rowHeight=100;
   }
   const main=w.worksheets.getItem(f.main);
   for(const u of f.updates)put(main,`N${u.main_row}`,`Đã chốt kế toán; STT import ${u.actions[0].seq}. ${u.decision}`);
  }else{
   put(s,'A4',`- ${f.roots.toLocaleString('vi-VN')} chuỗi / ${f.invoices.toLocaleString('vi-VN')} hóa đơn / ${f.rows.toLocaleString('vi-VN')} dòng. Đã bổ sung 42 hóa đơn theo chốt kế toán ngày 11/10; dự thảo ngày 12/10/2026.`);
   put(s,'A7','- Không còn hồ sơ chờ trả lời trong nhóm đã chốt. Giữ 6 cặp ngoài đợt xử lý và 3 cặp cùng sản phẩm; 8 khoản đảo tô vàng giữ chênh tổng tiền theo XML, cần kiểm tra khả năng nhập trên phần mềm.');
   const main=w.worksheets.getItem(f.main);
   for(const u of f.updates){
    const r=u.main_row;
    main.getRange(`AC${r}:AZ${r}`).clear({applyTo:'contents'});
    let note=`Đã chốt kế toán 11/10/2026. ${u.decision}`;
    if(u.mode==='Giữ nguyên'){
     note='GIỮ NGUYÊN — '+note;
     for(let i=0;i<3;i++)main.getRange(`${col(52+i)}${r}`).formulas=[[`=IF(ISNUMBER(${col(24+i)}${r}),${col(24+i)}${r},"")`]];
    }else{
     assert(u.actions.length>0&&u.actions.length<=4);
     for(let i=0;i<u.actions.length;i++){
      const a=u.actions[i],c=28+i*6;
      main.getRange(`${col(c)}${r}:${col(c+5)}${r}`).values=[[`${a.action}: ${a.target.join('/')} — STT import ${a.seq}`,null,day('2026-10-12'),...a.money]];
      main.getRange(`${col(c+2)}${r}`).setNumberFormat('dd/mm/yyyy');
     }
     for(let i=0;i<3;i++){
      const terms=[31+i,37+i,43+i,49+i].map(c=>`${col(c)}${r}`).join(',');
      main.getRange(`${col(52+i)}${r}`).clear({applyTo:'contents'});
      // Preserve unknown cumulative source values rather than treating source blanks as zero.
      main.getRange(`${col(52+i)}${r}`).formulas=[[`=IF(AND(ISNUMBER(${col(24+i)}${r}),COUNT(${terms})=${u.actions.length}),${col(24+i)}${r}+SUM(${terms}),"")`]];
     }
     if(u.actions.some(a=>a.difference!==0))note+=' Số đảo giữ lệch tổng tiền đã phát hành; cần kiểm tra khả năng nhập trên phần mềm.';
     if(u.case==='CC07')note+=' Giữ 9186; thanh toán nguồn của9186 trống, không tự đổi thành0 trong bảng nguồn.';
    }
    put(main,`AB${r}`,note);
    main.getRange(`AB${r}:AB${r}`).format.rowHeight=Math.max(90,u.actions.length*45);
   }
   const pending=w.worksheets.getItem('Cho KT xac nhan');
   put(pending,'A1','19 chuỗi đã chốt — 13 chuỗi bổ sung import, 6 chuỗi giữ nguyên cả cặp');
   put(pending,'G2','Trạng thái sau chốt');put(pending,'H2','Phương án đã chốt / STT import');
   for(const x of f.pending_rows){const dec=d.decisions[x.case],acts=f.new_actions.filter(a=>a.case===x.case);put(pending,`G${x.row}`,dec.mode==='Giữ nguyên'?'Giữ nguyên cả chuỗi':'Đã bổ sung import');put(pending,`H${x.row}`,`${dec.decision}${acts.length?' STT import: '+acts.map(a=>a.seq).join(', '):''}`);pending.getRange(`A${x.row}:L${x.row}`).format.rowHeight=110;}
   const check=w.worksheets.getItem('42 - Can check 11');
   put(check,'A1','11 chuỗi đã chốt — 8 chuỗi làm lại, 3 chuỗi giữ nguyên');
   for(const x of f.check11_rows){const dec=d.decisions[x.case],acts=f.new_actions.filter(a=>a.case===x.case);put(check,`Z${x.row}`,`${dec.mode==='Giữ nguyên'?'Giữ nguyên':'Đã bổ sung import'}. ${dec.decision}${acts.length?' STT: '+acts.map(a=>a.seq).join(', '):''}`);}
   put(g,'A3','1.047 F0 lịch sử, 2.113 hóa đơn nguồn riêng biệt. Trong đợt này 6 cặp được chốt giữ nguyên toàn bộ, 3 cặp cùng sản phẩm giữ nguyên; các hồ sơ đã trao đổi không còn chờ trả lời.');
   put(g,'A5','Bảng giữ nguồn lịch sử; đã xóa phương án F4–F7 của 9 hồ sơ giữ nguyên. 6 cặp giữ nguyên cả chuỗi được bổ sung vào bộ Ngoai le, không có trong import.');
   put(g,'A12','Bộ dữ liệu chuẩn bị theo phương án kế toán đã chốt; chưa nhập, ký hoặc phát hành hóa đơn.');
   put(g,'A13',`Import dieu chinh: ${f.roots.toLocaleString('vi-VN')} chuỗi / ${f.invoices.toLocaleString('vi-VN')} hóa đơn / ${f.rows.toLocaleString('vi-VN')} dòng. Giữ nguyên 1.990 hóa đơn trước; bổ sung 42 hóa đơn mới.`);
   put(g,'A14','Cho KT xac nhan: 19 hồ sơ đã chốt, 13 chuỗi bổ sung import và 6 chuỗi giữ nguyên cả cặp. 42 - Can check 11: 8 chuỗi làm lại và 3 chuỗi cùng sản phẩm giữ nguyên.');
   put(g,'A15','Phần bổ sung ngày 12/10/2026. Đảo điều chỉnh cũ rồi giảm gốc theo từng phương án; 9072 chỉ giảm gốc và giữ9186. Không tự thêm số lượng/đơn giá cho dòng khoản tiền.');
   put(g,'A16','8 khoản đảo giữ đúng số XML đã phát hành dù lệch tiền hàng + thuế − thanh toán; tô vàng tổng tiền ở import, cần kiểm tra khả năng nhập trên phần mềm. Gốc6/16 chỉ giảm số lượng; các ô nguồn trống vẫn giữ trống.');
  }
  g.getRange('A11:A16').format={font:{name:'Arial',size:10},wrapText:true,rowHeight:55,columnWidth:145};
 }else{
  const s=w.worksheets.getItem('Ngoai le'),source=w.worksheets.getItem('Chi tiet nguon'),g=w.worksheets.getItem('Huong dan');
  const vals=f.exception_rows.map(r=>r.map((v,i)=>[4,10,16].includes(i)?day(v):v));
  s.getRange('A241:X246').values=vals;wrap(s,'A241:X246',110);s.getRange('A241:B246').format.wrapText=true;s.getRange('X241:X246').format.wrapText=true;
  for(const c of ['E','K','Q'])s.getRange(`${c}241:${c}246`).setNumberFormat('dd/mm/yyyy');
  for(const c of ['D','J','P'])s.getRange(`${c}241:${c}246`).setNumberFormat('00000000');
  for(const range of ['F241:H246','L241:N246','R241:W246'])s.getRange(range).setNumberFormat(money);
  replaceTable(s,'A2:X246');
  source.getRange('A328:M339').values=f.source_rows.map(r=>r.map((v,i)=>i===4?day(v):v));wrap(source,'A328:M339',90);source.getRange('E328:E339').setNumberFormat('dd/mm/yyyy');source.getRange('J328:L339').setNumberFormat(money);
  replaceTable(source,'A1:M339');
  put(g,'A2','F0 từ 01/01/2024 đến 30/06/2025. Quy tắc giữ nguyên được rà thêm theo các liên kết và thay đổi người mua trong dữ liệu xuất lại đến 09/10/2026.');
  put(g,'A3','244 chuỗi giữ nguyên / 338 hóa đơn liên quan: gồm 238 ngoại lệ trước và 6 cặp vừa được chốt giữ nguyên toàn bộ.');
  put(g,'A4','F0 năm 2024: 121; F0 tháng 1–6/2025: 123.');
  put(g,'A5','Giữ nguyên toàn chuỗi theo tên công ty/hộ kinh doanh, định danh thuộc quy tắc ngoại lệ trước, các lần điều chỉnh đổi người mua và 6 cặp được chỉ định giữ nguyên.');
  put(g,'A11','Theo chốt điểm6: giữ mọi hóa đơn trong chuỗi nếu bất kỳ hóa đơn nào xuất cho công ty/hộ kinh doanh hoặc đã đổi người mua thành công ty/hộ kinh doanh. Giữ thêm6cặp được chỉ định, gồm3/508.');
  put(g,'A12','6 chuỗi mới được thêm ở cuối Ngoai le và Chi tiet nguon; bảng đã điều chỉnh vẫn lưu nguồn lịch sử nhưng đã xóa phương án phát sinh cho các chuỗi này. Không có hóa đơn thuộc diện giữ nguyên trong hai bảng import.');
  g.getRange('A11:A12').format={font:{name:'Arial',size:10},wrapText:true,rowHeight:65,columnWidth:145};
 }
 for(const meta of f.sheets){const s=w.worksheets.getItem(meta.name);s.freezePanes.unfreeze();if(meta.pane){const m=meta.pane.match(/^([A-Z]+)(\d+)$/);let n=0;for(const c of m[1])n=n*26+c.charCodeAt(0)-64;if(Number(m[2])>1)s.freezePanes.freezeRows(Number(m[2])-1);if(n>1)s.freezePanes.freezeColumns(n-1);}}
 w.recalculate();
 const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:50},summary:'formula error scan'});
 await fs.writeFile(path.join(here,`${f.kind}-errors.ndjson`),errors.ndjson);
 await render(w,f,'after');
 console.log('EXPORT',f.kind);
 await(await SpreadsheetFile.exportXlsx(w)).save(path.join(here,'staged',f.name));
 console.log('DONE',f.kind);
}
