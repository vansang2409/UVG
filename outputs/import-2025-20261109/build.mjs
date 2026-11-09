import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const dir=process.argv[3] ?? 'D:/UVG/outputs/import-2025-20261109';
const p=JSON.parse(await fs.readFile(`${dir}/authoring.json`,'utf8'));
const reporting=p.summary.amis_reporting_0911;
const serviceInput=p.manifest.find(m=>m.file.endsWith('service-quantity-confirmations.json'));
const confirmedServices=serviceInput?JSON.parse(await fs.readFile(`D:/UVG/${serviceInput.file}`,'utf8')).records:[];
const roundingInput=p.manifest.find(m=>m.file.endsWith('rounding-confirmation.json'));
const roundingPolicy=roundingInput?JSON.parse(await fs.readFile(`D:/UVG/${roundingInput.file}`,'utf8')):null;
const buyerPolicy=p.manifest.some(m=>m.file.endsWith('buyer-field-confirmation.json'));
await fs.mkdir(`${dir}/deliverables`,{recursive:true});
await fs.mkdir(`${dir}/qa`,{recursive:true});
const col=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
function makeSheet(w,name,headers,rows,start=1,kind='report'){
 const s=w.worksheets.add(name),last=col(headers.length),end=start+rows.length;
 s.showGridLines=false;
 s.getRangeByIndexes(start-1,0,1,headers.length).values=[headers];
 for(let i=0;i<rows.length;i+=1000){let part=rows.slice(i,i+1000);s.getRangeByIndexes(start+i,0,part.length,headers.length).values=part;}
 const body=s.getRangeByIndexes(start-1,0,Math.max(1,rows.length+1),headers.length);
 body.format={font:{name:'Arial',size:10},rowHeight:['adjustment','replacement','return_adjustment'].includes(kind)?60:28,verticalAlignment:'center'};
 s.getRange(`A${start}:${last}${start}`).format={fill:'#24435C',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:56,wrapText:true,horizontalAlignment:'center'};
 headers.forEach((h,i)=>{
  let width=/Tên hàng|Ghi chú|Chờ vì|Lý do/.test(h)?70:/Tên|Khách|Địa chỉ|Chuỗi|Các HĐ/.test(h)?38:/Dòng nguồn|Chứng từ/.test(h)?27:23;
  s.getRangeByIndexes(0,i,1,1).format.columnWidth=width;
  if(rows.length){let r=s.getRangeByIndexes(start,i,rows.length,1);
   if(/Mã|Số hóa đơn|Ký hiệu|Số chứng từ|Chứng từ AMIS|Chuỗi HĐ|HĐ gốc/.test(h))r.setNumberFormat('@');
   const fractional=rows.some(row=>typeof row[i]==='number'&&!Number.isInteger(row[i]));
   if(/Tiền|tiền|VAT|Thanh toán|thanh toán|Thành tiền|Đơn giá|Số dư/.test(h))r.setNumberFormat(fractional?'#,##0.00####;[Red](#,##0.00####);0':'#,##0;[Red](#,##0);0');
   if(/Số lượng|SL trả/.test(h))r.setNumberFormat(fractional?'#,##0.00####':'#,##0');
   if(kind!=='report'&&(/Ngày hóa đơn|Ngày HĐ/.test(h)))r.setNumberFormat('dd/mm/yyyy');
   if(/Tên|Ghi chú|Chờ vì|Lý do|Các HĐ|Chuỗi|Kiểm chi tiết/.test(h))r.format.wrapText=true;
  }
 });
 if(rows.length){s.tables.add(`A${start}:${last}${end}`,true,'T'+name.replace(/[^A-Za-z0-9]/g,'')).style='TableStyleMedium2';}
 if(kind==='report'&&headers.some(h=>/Ghi chú|Chờ vì/.test(h))&&rows.length)s.getRangeByIndexes(start,0,rows.length,headers.length).format.autofitRows();
 s.freezePanes.freezeRows(start);s.freezePanes.freezeColumns(kind==='report'?1:2);
 return s;
}
const only=process.argv[2];
for(let index=0;index<p.outputs.length;index++){
 if(only!==undefined&&index!==Number(only))continue;
 const o=p.outputs[index],w=Workbook.create();
 if(o.kind==='report'){
  for(const sh of o.sheets){
   const detail=makeSheet(w,sh.name,sh.headers,sh.rows,sh.start||1,sh.kind||'report');
   if(sh.name==='Ra soat 12 don T6'){detail.getRange('F1:F13').format.columnWidth=38;detail.getRange('F2:F13').format.wrapText=true;}
   if(sh.notes){sh.notes.forEach((t,i)=>{detail.mergeCells(`A${i+1}:H${i+1}`);detail.getRange(`A${i+1}`).values=[[t]];detail.getRange(`A${i+1}:H${i+1}`).format={wrapText:true,rowHeight:30};});detail.getRange(`J10:M${sh.rows.length+9}`).format.fill='#FFF2CC';}
  }
  const s=w.worksheets.getItem('Tong hop');
  s.getRange('A14').values=[['Tổng']];s.getRange('B14:G14').formulas=[['=SUM(B2:B13)','=SUM(C2:C13)','=SUM(D2:D13)','=SUM(E2:E13)','=SUM(F2:F13)','=SUM(G2:G13)']];
  s.getRange('A14:G14').format={fill:'#E6EEF4',font:{bold:true},rowHeight:30};s.getRange('G2:G14').setNumberFormat('#,##0');
  const returnStatus=p.summary.return_status_sync?.documents.map(item=>item.code).join(', ');
  const returnDetail=returnStatus
   ?`${returnStatus} dùng số tiền trên phiếu và chờ số HĐ bán mới.`
   :'2 chứng từ có ngoại lệ tiền trả.';
  const notes=[['Phạm vi',`Báo cáo lấy AMIS 2025 xuất ${reporting?.source_date ?? '08/11/2026'}. Danh sách công ty và AMIS bản sao đọc ngày 09/11/2026.`],['Trạng thái','DỰ THẢO. Chưa nhập, ký hoặc phát hành.'],['Trả hàng',`${p.summary.source_returns_docs} phiếu/${p.summary.source_returns_lines} dòng, tổng thanh toán ${p.summary.source_returns_pay}đ. Đã nối số CT bán gốc từ AMIS live. Bán gốc và điều chỉnh trả hàng tách riêng.`],['Ngày hóa đơn','Dự thảo 09/11/2026; ngày AMIS 2025 giữ ở cột kiểm tra, không dùng làm ngày phát hành mới.'],['Khớp tiền','Giữ nguyên các đơn khớp tiền hàng/VAT/thanh toán. Khác tên hàng/ĐVT/SL nguyên văn vẫn ghi ở CSV toàn bộ; chưa kết luận sai hàng vì có combo/tên khác.'],['Lịch sử meInvoice','Đến 29/10/2026 và bổ sung 01/11/2026. Cần kiểm lại phát sinh mới hơn trước khi nhập.'],['Danh sách nguồn công ty','https://docs.google.com/spreadsheets/d/1bhOnhnBcHaHiy06-REAdVVEk160FG94a6XgkfH0AQsw/edit#gid=622685968'],['AMIS bản sao','https://docs.google.com/spreadsheets/d/1C5hoPDvZKB1vGFSUG9rzChcORt2L9NnsqETVHxPui9I/edit#gid=1264954379'],['Dự thảo trả lại',`${p.summary.returns.draft_docs} phiếu/${p.summary.returns.draft_lines} dòng trong tab Tra lai cho so HD. Điền cột J:M theo HĐ thực tế trước khi nhập. ${returnDetail}`]];
  if(confirmedServices.length)notes.push(['SL dịch vụ đã xác nhận',confirmedServices.map(c=>`${c.doc}: SL nguồn ${c.source_quantity}, SL import ${c.import_quantity} ${c.unit}`).join('. ')+'. Áp dụng xác nhận người dùng cho hai dòng dịch vụ này; giữ đơn giá/tiền hàng/VAT/thanh toán và file AMIS gốc.']);
  if(roundingPolicy)notes.push(['Chênh làm tròn đã chấp nhận',`Chấp nhận tối đa ${roundingPolicy.max_abs_delta_vnd}đ giữa SL × đơn giá và thành tiền AMIS theo xác nhận người dùng. Giữ nguyên SL/đơn giá/tiền hàng/VAT/thanh toán nguồn; không nới tiêu chí khớp AMIS–meInvoice. Bổ sung ${roundingPolicy.newly_eligible_documents.length} chứng từ/${roundingPolicy.newly_eligible_rows} dòng.`]);
  for(let i=0;i<notes.length;i++){let row=17+i;s.getRange(`A${row}`).values=[[notes[i][0]]];s.mergeCells(`B${row}:G${row}`);s.getRange(`B${row}`).values=[[notes[i][1]]];s.getRange(`A${row}:G${row}`).format={rowHeight:43,wrapText:true};}
  if(o.sheets.some(sh=>sh.name==='Ra soat 12 don T6')){
   const r=await w.render({sheetName:'Ra soat 12 don T6',range:'I1:K3',scale:1,format:'png'});await fs.writeFile(`${dir}/preview-h1-12-links.png`,new Uint8Array(await r.arrayBuffer()));
  }
 }else{
  const s=makeSheet(w,'Hoa don GTGT',o.headers,o.rows,9,o.kind);
  const title=o.kind==='return_adjustment'?'DỰ THẢO IMPORT ĐIỀU CHỈNH TRẢ LẠI HÀNG BÁN 2025':o.kind==='new'?'DỰ THẢO IMPORT HÓA ĐƠN MỚI 2025':o.kind==='adjustment'?'DỰ THẢO IMPORT ĐIỀU CHỈNH VỀ 0':'DỰ THẢO IMPORT THAY THẾ VỀ 0';
  const notes=[title,...o.notes,'Chưa nhập, ký hoặc phát hành. Kiểm tra báo cáo và phát sinh mới trước khi dùng.'];
  notes.forEach((text,i)=>{s.mergeCells(`A${i+1}:H${i+1}`);s.getRange(`A${i+1}`).values=[[text]];s.getRange(`A${i+1}:H${i+1}`).format={font:{name:'Arial',size:10,bold:i===0},wrapText:true,rowHeight:30};});
  if(o.rows.length)s.getRange(`B10:B${9+o.rows.length}`).setNumberFormat('dd/mm/yyyy');
  if(o.kind==='return_adjustment'&&o.rows.length){
   s.getRange(`J10:M${9+o.rows.length}`).format.fill='#FFF2CC';
   s.getRange(`AE10:AE${9+o.rows.length}`).conditionalFormats.addCustom('=$AE10="Chờ kiểm tra số lượng/tiền trả"',{fill:'#FDE9E7',font:{bold:true,color:'#9C0006'}});
  }
 }
 if(o.kind==='report'&&buyerPolicy){const s=w.worksheets.getItem('Tong hop');s.getRange('A28:B28').values=[['Tên đơn vị / người mua','Khách lẻ: trống tên đơn vị, dùng Người mua hàng. Giữ tên người mua đã có và thông tin pháp nhân/MST nguồn; không đổi số tiền.']];s.mergeCells('B28:G28');s.getRange('A28:G28').format={font:{name:'Arial',size:10},rowHeight:44,wrapText:true,verticalAlignment:'center'};}
 if(o.kind==='report'&&p.summary.unmapped_meinvoice){
  const s=w.worksheets.getItem('Tong hop'),u=p.summary.unmapped_meinvoice;
  const keeps=p.summary.gift_keeps_and_amis_deletion_20261109;
  const gifts=u.gift_kt_kept?`KT chốt giữ ${u.gift_kt_kept} HĐ biếu tặng, theo dõi riêng.`:`${u.gift_kt_check} HĐ biếu tặng tách KT kiểm.`;
  const confirmed=keeps?` Giữ thêm ${keeps.records.length} HĐ KOL/KOC. ${keeps.deleted_amis.join(', ')} đã xóa trên AMIS theo xác nhận.`:'';
  s.getRange('A29:B29').values=[['Bổ sung meInvoice2025 chưa nối AMIS',`Đã kiểm toàn bộ chuỗi liên quan. Bổ sung ${u.added.adjustment||0} điều chỉnh và ${u.added.replacement||0} thay thế về 0. ${gifts} ${u.remaining_source_exceptions} ngoại lệ nguồn chưa nhập.${confirmed}`]];
  s.mergeCells('B29:G29');s.getRange('A29:G29').format={font:{name:'Arial',size:10},rowHeight:70,wrapText:true,verticalAlignment:'center'};
 }
 if(o.kind==='report'&&p.adjustment_rule_2024){
  const s=w.worksheets.getItem('Tong hop'),plan=p.adjustment_rule_2024;
  const originals=plan.chains.length, reversals=plan.invoices-originals;
  s.getRange('B28').values=[['Riêng điều chỉnh về0: Tên khách hàng và Người mua hàng chép riêng nguyên trạng HĐ F0; các bộ bán mới/thay thế/BTL giữ quy tắc tên trước đó.']];
  s.getRange('A30:B30').values=[['Điều chỉnh theo quy trình2024',`${originals} chuỗi / ${plan.invoices} HĐ / ${plan.line_count} dòng: đảo ${reversals} khoản cũ còn tiền, sau đó giảm chi tiết ${originals} F0. ${plan.skipped_zero_adjustments.length} khoản cũ 0đ không tạo thêm. Số dư cuối bằng 0 đủ ba khoản.`]];
  s.mergeCells('B30:G30');s.getRange('A30:G30').format={font:{name:'Arial',size:10},rowHeight:70,wrapText:true,verticalAlignment:'center'};
 }
 const exchanges=p.summary.psc2_exchanges;
 if(o.kind==='report'&&exchanges){
  const s=w.worksheets.getItem('Tong hop');
  const details=reporting
   ?reporting.vat_policy_delta===0
    ?'Sáu chứng từ đã bỏ PSC2. Tiền hàng, VAT và thanh toán trên dự thảo lấy theo AMIS 09/11.'
    :`Bản AMIS 09/11 đã bỏ 6 dòng PSC2. VAT dự thảo chênh ${reporting.vat_policy_delta.toLocaleString('vi-VN')}đ so với AMIS.`
   :Object.entries(exchanges).map(([code,e])=>`${code} trừ PSC2 ${e.excluded_goods}đ, ${e.rate_policy==='recalculate'?'tính lại VAT 8% thành':'giữ VAT'} ${e.vat}đ.`).join(' ');
  s.mergeCells('B31:G31');
  s.getRange('A31:B31').values=[['Đổi sản phẩm PSC2',details]];
  s.getRange('A31:G31').format={font:{name:'Arial',size:10},rowHeight:44,wrapText:true,verticalAlignment:'center'};
 }
 if(o.kind==='report'&&reporting){
  const s=w.worksheets.getItem('Tong hop');
  const labels=['Đối chiếu cùng phạm vi','AMIS theo bản 09/11/2026','Doanh thu theo phương án dự kiến','Chênh lệch dự kiến so với AMIS'];
  labels.forEach((label,i)=>{const row=33+i;s.mergeCells(`A${row}:E${row}`);s.getRange(`A${row}`).values=[[label]];});
  s.getRange('F33:G35').values=[['Tiền hàng (đ)','VAT (đ)'],reporting.scope_current,reporting.projection];
  s.getRange('F36:G36').formulas=[['=F35-F34','=G35-G34']];
  s.getRange('A33:G36').format={font:{name:'Arial',size:10},rowHeight:28,verticalAlignment:'center'};
  s.getRange('A33:G33').format={fill:'#24435C',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'}};
  s.getRange('A36:G36').format={fill:'#E6EEF4',font:{name:'Arial',size:10,bold:true}};
  s.getRange('F34:G36').setNumberFormat('#,##0;[Red](#,##0);0');
 }
 w.recalculate();
 if(o.kind==='new'&&exchanges){
  const rows=o.rows.flatMap((r,i)=>r[18] in exchanges?[i+10]:[]);
  const q=await w.render({sheetName:'Hoa don GTGT',range:`K${rows.at(-1)-1}:Q${rows.at(-1)}`,scale:1,format:'png'});
  await fs.writeFile(`${dir}/qa/preview-psc2.png`,new Uint8Array(await q.arrayBuffer()));
 }
 if(o.kind==='new'&&buyerPolicy){const q=await w.render({sheetName:'Hoa don GTGT',range:'C9:F13',scale:1,format:'png'});await fs.writeFile(`${dir}/preview-buyer-fields.png`,new Uint8Array(await q.arrayBuffer()));}
 if(o.kind==='new'&&confirmedServices.length){
  const codes=new Set(confirmedServices.map(c=>c.doc)),rs=o.rows.flatMap((r,i)=>codes.has(r[18])?[i+10]:[]);
  if(rs.length){const q=await w.render({sheetName:'Hoa don GTGT',range:`M${Math.min(...rs)}:S${Math.max(...rs)}`,scale:1,format:'png'});await fs.writeFile(`${dir}/preview-service-quantities.png`,new Uint8Array(await q.arrayBuffer()));}
 }
 if(o.kind==='new'&&o.headers.includes('Hàng khuyến mại')){
  const pr=o.rows.flatMap((r,i)=>r[22]===1?[i+10]:[]);
  if(pr.length){const q=await w.render({sheetName:'Hoa don GTGT',range:`S${Math.min(...pr)}:W${Math.max(...pr)}`,scale:1,format:'png'});await fs.writeFile(`${dir}/preview-promotions.png`,new Uint8Array(await q.arrayBuffer()));}
 }
 const sh=o.kind==='report'?'Tong hop':'Hoa don GTGT',range=o.kind==='report'?'A1:G14':'J9:Q12';
 const preview=await w.render({sheetName:sh,range,scale:1,format:'png'});await fs.writeFile(`${dir}/preview-${index}.png`,new Uint8Array(await preview.arrayBuffer()));
 if(o.kind==='report'){
  if(p.summary.unmapped_meinvoice?.gift_kt_kept){
   const q=await w.render({sheetName:'Bieu tang - KT check',range:'I1:K5',scale:1,format:'png'});
   await fs.writeFile(`${dir}/qa/gift-keep-after.png`,new Uint8Array(await q.arrayBuffer()));
  }
  for(const name of ['MI2025 chua map','MI2025 chuoi','Bieu tang - KT check'])if(o.sheets.some(sh=>sh.name===name)){const q=await w.render({sheetName:name,range:'A1:G5',scale:1,format:'png'});await fs.writeFile(`${dir}/qa/preview-${name.replaceAll(' ','-')}.png`,new Uint8Array(await q.arrayBuffer()));}
  for(const [sn,rg,key]of [['Tra lai cho so HD','I9:P12','return-reference'],['Tra lai cho so HD','O9:W12','return-money'],['Cho xac minh','I1:K4','pending-reasons']]){const q=await w.render({sheetName:sn,range:rg,scale:1,format:'png'});await fs.writeFile(`${dir}/preview-${key}.png`,new Uint8Array(await q.arrayBuffer()));}
 }
 console.log((await w.inspect({kind:'region',sheetId:sh,range:o.kind==='report'?'A14:G14':'A9:D11',maxChars:700,tableMaxRows:3,tableMaxCols:4})).ndjson);
 await(await SpreadsheetFile.exportXlsx(w)).save(`${dir}/deliverables/${o.file}`);
 // Artifact Tool emits a technical inspection sidecar; keep it out of the delivery folder.
 const sidecar=`${dir}/deliverables/${o.file}.inspect.ndjson`;
 try{await fs.rename(sidecar,`${dir}/qa/${o.file}.inspect.ndjson`);}catch(e){if(e.code!=='ENOENT')throw e;}
 console.log(JSON.stringify({saved:o.file,rows:o.rows?.length,invoices:o.invoices}));
}
