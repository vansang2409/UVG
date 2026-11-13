import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';

const out=path.dirname(fileURLToPath(import.meta.url));
const d=JSON.parse(await fs.readFile(path.join(out,'issued-2026-report-data.json'),'utf8'));
const verification=JSON.parse(await fs.readFile(path.join(out,'issued-2026-source-verification.json'),'utf8'));
if(!verification.passed)throw new Error('Source verification required');
if(process.argv.includes('--fit-reference')){
 const filePath=path.join(out,'TONG_HOP_HD_PHAT_HANH_2026_CHO_DON_2024_2025.xlsx');
 console.log('IMPORTING for reference layout');
 const fitted=await SpreadsheetFile.importXlsx(await FileBlob.load(filePath));
 const sh=fitted.worksheets.getItem('BHH và CT liên quan');
 sh.getRange(`L6:M${d.bhh_rows.length+5}`).format.wrapText=true;
 sh.getRange(`A6:M${d.bhh_rows.length+5}`).format.autofitRows();
 fitted.recalculate();
 const totals=fitted.worksheets.getItem('Tổng hợp').getRange('B11:G11').values[0];
 if(JSON.stringify(totals)!==JSON.stringify([40161,42046,82207,...d.summary.money_sums]))throw new Error('Totals changed during formatting');
 const idx=d.bhh_rows.findIndex(r=>String(r[12]).includes('\n'))+6;
 for(const [suffix,range] of [['left','A1:H12'],['references',`J${Math.max(5,idx-1)}:M${idx+3}`]]){
  const p=await fitted.render({sheetName:'BHH và CT liên quan',range,scale:1,format:'png'});
  await fs.writeFile(path.join(out,'issued-preview-BHH-'+suffix+'.png'),new Uint8Array(await p.arrayBuffer()));
 }
 await (await SpreadsheetFile.exportXlsx(fitted)).save(filePath);
 console.log('REFORMATTED references');
 process.exit(0);
}
const wb=Workbook.create();
const summary=wb.worksheets.add('Tổng hợp');
const detail=wb.worksheets.add('HĐ 2026');
const amis=wb.worksheets.add('BHH và CT liên quan');
const missing=wb.worksheets.add('Chuỗi thiếu BHH');
const ink='#233348',head='#253B53',muted='#5A6878',money='#,##0;[Red](#,##0);0';
const cc=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const dt=x=>x?new Date(x.slice(0,10)+'T00:00:00Z'):null;
function styleSheet(sh,cols,end,header,widths){
 sh.showGridLines=false;
 sh.getRange(`A1:${cc(cols)}${end}`).format.font={name:'Arial',size:10,color:ink};
 sh.getRange(`A1:${cc(cols)}${end}`).format.verticalAlignment='center';
 sh.getRange(`A1:A${end}`).format.rowHeight=23;
 sh.getRange('A2').format.font={name:'Arial',size:15,bold:true,color:ink};
 if(header){
  sh.getRange(`A${header}:${cc(cols)}${header}`).format={fill:head,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center',wrapText:true,rowHeight:46};
  sh.freezePanes.freezeRows(header);
 }
 widths.forEach((x,i)=>sh.getRange(`${cc(i+1)}1`).format.columnWidthPx=x);
}
function writeBlocks(sh,headers,rows,tableName){
 sh.getRange(`A5:${cc(headers.length)}5`).values=[headers];
 for(let start=0;start<rows.length;start+=4000){
  const chunk=rows.slice(start,start+4000);
  sh.getRangeByIndexes(start+5,0,chunk.length,headers.length).values=chunk;
  if(start%20000===0)console.log('WRITE',sh.name,start+chunk.length,'/',rows.length);
 }
 const t=sh.tables.add(`A5:${cc(headers.length)}${rows.length+5}`,true,tableName);
 t.showFilterButton=true;
}

// Đưa ký hiệu và số HĐ sang trái, giữ nguyên cặp khóa để lọc và đối chiếu.
const order=[0,1,6,7,8,2,3,5,9,10,11,4,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29];
const rows=d.rows.map(r=>order.map(i=>[8,14].includes(i)?dt(r[i]):r[i]));
detail.getRange('A2').values=[['HĐ phát hành năm 2026 cho đơn 2024–2025']];
detail.getRange('A3').values=[['Mỗi dòng là một HĐ duy nhất. Nguồn đến 13/11/2026. Cột F: BHH; cột W: HĐ hiện map AMIS; cột AD: HĐ cũ cùng BHH.']];
writeBlocks(detail,order.map(i=>d.headers[i]),rows,'IssuedInvoices2026');
styleSheet(detail,30,rows.length+5,5,[62,90,115,120,115,210,180,235,155,145,165,225,120,125,115,130,130,105,270,165,155,180,250,370,215,160,420,630,660,290]);
detail.freezePanes.freezeColumns(4);detail.tabColor=head;
detail.getRange(`E6:E${rows.length+5}`).setNumberFormat('dd/mm/yyyy');
detail.getRange(`O6:O${rows.length+5}`).setNumberFormat('dd/mm/yyyy');
detail.getRange(`I6:K${rows.length+5}`).setNumberFormat(money);
detail.getRange(`T6:V${rows.length+5}`).setNumberFormat(money);
detail.getRange(`A6:B${rows.length+5}`).format.horizontalAlignment='center';
detail.getRange(`C6:D${rows.length+5}`).setNumberFormat('@');
detail.getRange(`M6:N${rows.length+5}`).setNumberFormat('@');
detail.getRange(`P6:Q${rows.length+5}`).setNumberFormat('@');
detail.getRange(`F6:F${rows.length+5}`).conditionalFormats.addCustom('AND($F6="",$G6="")',{fill:'#FFF1D6',font:{color:'#875D0F'}});
detail.getRange(`Z6:Z${rows.length+5}`).conditionalFormats.add('containsText',{text:'không hợp lệ',format:{fill:'#FCE7E7',font:{color:'#9C2626'}}});

const ar=d.bhh_rows.map(r=>r.map((v,i)=>i===2?dt(v):i===9?v.replace(/^.*? \| /,'AMIS | '):v));
amis.getRange('A2').values=[['BHH và chứng từ AMIS liên quan']];
amis.getRange('A3').values=[['Một dòng mỗi chứng từ. Cột E: HĐ hiện tại; cột L: HĐ cũ; cột M: toàn bộ HĐ phát hành 2026 liên quan.']];
writeBlocks(amis,d.bhh_headers,ar,'AmisDocumentReferences');
styleSheet(amis,13,ar.length+5,5,[95,205,115,410,250,165,155,175,255,620,650,300,460]);
amis.freezePanes.freezeColumns(2);
amis.getRange(`C6:C${ar.length+5}`).setNumberFormat('dd/mm/yyyy');
amis.getRange(`F6:H${ar.length+5}`).setNumberFormat(money);

const ur=d.unmapped_rows.map(r=>r.map((v,i)=>i===3?dt(v):v));
missing.getRange('A2').values=[['Chuỗi HĐ cũ chưa xác định BHH']];
missing.getRange('A3').values=[['Giữ đầy đủ các HĐ 2026 của chuỗi đã về 0. Những chuỗi này thiếu mã đơn/liên kết BHH trong nguồn.']];
writeBlocks(missing,d.unmapped_headers,ur,'OldChainsWithoutDocuments');
styleSheet(missing,12,ur.length+5,5,[65,95,250,115,460,105,290,160,145,175,420,670]);
missing.freezePanes.freezeColumns(3);
missing.getRange(`D6:D${ur.length+5}`).setNumberFormat('dd/mm/yyyy');
missing.getRange(`H6:J${ur.length+5}`).setNumberFormat(money);
missing.getRange(`E6:E${ur.length+5}`).format.wrapText=true;
for(let i=0;i<ur.length;i++){
 const lines=String(ur[i][4]||'').split('\n').length;
 if(lines>1)missing.getRange(`A${i+6}`).format.rowHeight=Math.min(400,lines*15+8);
}
missing.getRange(`G6:G${ur.length+5}`).conditionalFormats.add('containsText',{text:'trống',format:{fill:'#FFF1D6',font:{color:'#875D0F'}}});

summary.getRange('A2').values=[['Tổng hợp HĐ phát hành 2026 cho đơn 2024–2025']];
summary.getRange('A3').values=[['Ngày chốt nguồn: 13/11/2026. Bao gồm HĐ xuất mới/xuất lại, điều chỉnh và thay thế trong các chuỗi liên quan.']];
summary.getRange('A5:G5').values=[['Loại HĐ phát hành 2026','Đơn 2024','Đơn 2025','Tổng số HĐ','Tiền hàng trên HĐ (đ)','VAT trên HĐ (đ)','Thanh toán có số liệu (đ)']];
const categories=['Xuất mới / xuất lại','Điều chỉnh trong chuỗi về 0','Thay thế bằng HĐ 0','Điều chỉnh khác','Thay thế khác'];
summary.getRange('A6:A10').values=categories.map(x=>[x]);
summary.getRange('A11').values=[['Tổng HĐ phát hành 2026']];
summary.getRange('A14:B20').values=[
 ['BHH / BHDV đã nối được',d.summary.bhh_count],
 ['CT trả/giảm liên quan',d.summary.other_doc_count],
 ['Chuỗi HĐ tổng cũ chưa xác định BHH',d.summary.unmapped_chains],
 ['HĐ 2026 thuộc các chuỗi thiếu BHH',d.summary.unmapped_2026_invoices],
 ['HĐ 2026 chưa có số thanh toán',null],
 ['HĐ 2026 nối theo mã đơn',d.summary.mapping_basis['Mã đơn/chứng từ khớp nguyên vẹn bảng kê 13/11']||0],
 ['Bảng kê nguồn đã kiểm',verification.header_rows_read],
 ];
summary.getRange('A23').values=[['Cách đọc số liệu']];
const notes=[
 'Tiền trên HĐ cộng theo dấu của từng HĐ đã phát hành, bao gồm HĐ đã bị thay thế. Tổng này không đại diện doanh thu còn lại.',
 'Kết quả chuỗi ở cột S và số dư T–V lặp trên các HĐ cùng chuỗi. Tra theo F0, tránh cộng lặp số dư chuỗi.',
 'F0, F1, F2 được xác định theo liên kết điều chỉnh/thay thế. HĐ xuất lại được nối với HĐ cũ qua BHH ở cột AD khi có bằng chứng.',
 'Các ô thanh toán trống trên HĐ được giữ nguyên. Số dư chuỗi sử dụng số 0 ở các trường hợp đã có bằng chứng xác nhận riêng.',
 'Nguồn không có BHH/mã đơn được giữ ở tab Chuỗi thiếu BHH. Đây là nhóm HĐ đã về 0, khác phạm vi HĐ còn tiền đã rà trước đó.',
 'BHH có liên kết lịch sử được đối chiếu với mapping đã lưu. Căn cứ nối theo mã đơn được ghi rõ tại cột X để kiểm tra riêng.',
 'Các trường hợp mã đơn chỉ sang BHH khác mapping đã chốt được giữ theo mapping hiện tại, không tự ghi đè.',
 ];
summary.getRange('A24:A30').values=notes.map(x=>[x]);
summary.getRange('A33').values=[['Nguồn đối chiếu']];
summary.getRange('A34:A39').values=[
 ['AMIS: Sổ chi tiết bán hàng 2024 và 2025, bộ dữ liệu 13/11/2026. Vị trí chứng từ tại tab BHH và CT liên quan.'],
 ['meInvoice: bảng kê HĐ đã sử dụng, chi tiết, điều chỉnh và thay thế xuất ngày 13/11/2026. Dòng nguồn tại cột AC.'],
 ['Liên kết lịch sử: bảng đối chiếu BHH 09/11, Import HD moi đã bổ sung HĐ phát hành 11/11, mapping nghiệm thu 13/11.'],
 ['Bảng xử lý 2025: https://docs.google.com/spreadsheets/d/1jACpDPSlOCjq-25hF_g5-06eZKD3e2Fq88j472wjTiA'],
 ['Bảng kế toán: https://docs.google.com/spreadsheets/d/1C5hoPDvZKB1vGFSUG9rzChcORt2L9NnsqETVHxPui9I'],
 ['Quan hệ thiếu trong bảng kê mới được bổ sung từ 31 liên kết lịch sử đã kiểm, không có xung đột đồ thị.'],
 ];
styleSheet(summary,7,39,5,[335,130,130,135,195,175,225]);
summary.freezePanes.unfreeze();summary.tabColor=head;
summary.getRange('A5:G5').format.rowHeight=48;
summary.getRange('B6:D11').setNumberFormat('#,##0');summary.getRange('E6:G11').setNumberFormat(money);
summary.getRange('A11:G11').format={fill:'#E8EEF4',font:{name:'Arial',size:10,bold:true,color:ink},rowHeight:28};
summary.getRange('B14:B20').setNumberFormat('#,##0');summary.getRange('A16:B17').format.fill='#FFF1D6';
summary.getRange('A3').format.font={name:'Arial',size:10,italic:true,color:muted};
summary.getRange('A23').format.font.bold=true;summary.getRange('A33').format.font.bold=true;

// Công thức tổng hợp đọc bảng HĐ duy nhất để không nhân tiền theo số BHH.
const end=rows.length+5;
for(let r=6;r<=10;r++){
 summary.getRange(`B${r}:G${r}`).formulas=[[
  `=COUNTIFS('HĐ 2026'!$B$6:$B$${end},"2024",'HĐ 2026'!$H$6:$H$${end},$A${r})`,
  `=COUNTIFS('HĐ 2026'!$B$6:$B$${end},"2025",'HĐ 2026'!$H$6:$H$${end},$A${r})`,
  `=SUM(B${r}:C${r})`,
  `=SUMIFS('HĐ 2026'!$I$6:$I$${end},'HĐ 2026'!$H$6:$H$${end},$A${r})`,
  `=SUMIFS('HĐ 2026'!$J$6:$J$${end},'HĐ 2026'!$H$6:$H$${end},$A${r})`,
  `=SUMIFS('HĐ 2026'!$K$6:$K$${end},'HĐ 2026'!$H$6:$H$${end},$A${r})`,
 ]];
}
summary.getRange('B11:G11').formulas=[['=SUM(B6:B10)','=SUM(C6:C10)','=SUM(D6:D10)','=SUM(E6:E10)','=SUM(F6:F10)','=SUM(G6:G10)']];
summary.getRange('B18').formulas=[[`=COUNTBLANK('HĐ 2026'!K6:K${end})`]];
wb.recalculate();
const totals=summary.getRange('B11:G11').values[0];
if(JSON.stringify(totals)!==JSON.stringify([40161,42046,82207,...d.summary.money_sums]))throw new Error('Totals mismatch '+JSON.stringify(totals));
if(summary.getRange('B18').values[0][0]!==d.summary.missing_money_counts[2])throw new Error('Blank count mismatch');
console.log('SUMMARY VERIFIED',JSON.stringify(totals));
console.log((await wb.inspect({kind:'region',sheetId:'Tổng hợp',range:'A5:G11',tableMaxRows:7,tableMaxCols:7,maxChars:2400})).ndjson);
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30},maxChars:1800});
console.log('ERROR_SCAN',errors.ndjson);
for(const [sheetName,range] of [['Tổng hợp','A1:G21'],['HĐ 2026','A1:K13'],['BHH và CT liên quan','A1:H12'],['Chuỗi thiếu BHH','A1:G12']]){
 const p=await wb.render({sheetName,range,scale:1,format:'png'});
 await fs.writeFile(path.join(out,'issued-preview-'+sheetName+'.png'),new Uint8Array(await p.arrayBuffer()));
 console.log('RENDERED',sheetName);
 if(global.gc)global.gc();
}
const file=await SpreadsheetFile.exportXlsx(wb);
await file.save(path.join(out,'TONG_HOP_HD_PHAT_HANH_2026_CHO_DON_2024_2025.xlsx'));
await fs.writeFile(path.join(out,'issued-2026-authoring-verification.json'),JSON.stringify({passed:true,totals,blankPayment:summary.getRange('B18').values[0][0],sheets:[['HĐ 2026',rows.length],['BHH và CT liên quan',ar.length],['Chuỗi thiếu BHH',ur.length]]},null,2));
console.log('EXPORTED issued 2026');
