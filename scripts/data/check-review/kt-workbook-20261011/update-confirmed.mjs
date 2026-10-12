import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';

const dir=path.dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/,'$1'));
const data=JSON.parse(await fs.readFile(path.join(dir,'confirmed-plan-data.json'),'utf8'));
const outputPath=path.resolve(dir,'../../../../outputs/kt-review-20261011-01a12992/UVG_PHUONG_AN_DA_CHOT_20261011.xlsx');
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(data.source));
const key=k=>k.join('/');
const day=d=>d?new Date(`${d.slice(0,10)}T00:00:00Z`):null;
const fmtDay=d=>d?d.slice(0,10).split('-').reverse().join('/'):'ngày chưa có trong phần nguồn';
const money='#,##0;[Red](#,##0);0';
const col=n=>{let s='';for(n++;n>0;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const navy='#203E59', soft='#F2F5F8', teal='#E3F1EA', red='#FCE8E6';
await fs.mkdir(path.join(dir,'confirmed-previews'),{recursive:true});
if(process.argv.includes('--inspect')){
  const inspection=await wb.inspect({kind:'workbook,sheet,table',maxChars:2500,tableMaxRows:2,tableMaxCols:4});
  await fs.writeFile(path.join(dir,'confirmed-before.ndjson'),inspection.ndjson);
  for(const [i,name,range] of [[1,'Tong hop','A1:G25'],[2,'Cach ghi dieu chinh','A6:N9']]){
    const blob=await wb.render({sheetName:name,range,scale:1,format:'png'});
    await fs.writeFile(path.join(dir,'confirmed-previews',`before-${i}.png`),new Uint8Array(await blob.arrayBuffer()));
  }
  console.log('Rendered source summary and representative response rows.');
  process.exit(0);
}

const actionsSheet=wb.worksheets.add('Phuong an phat sinh');
const holdSheet=wb.worksheets.add('Giu nguyen ca chuoi');
function extendTable(sheet,endCol,endRow){
  const old=sheet.tables.items[0], name=old.name;
  old.delete();
  const t=sheet.tables.add(`A7:${endCol}${endRow}`,true,name);
  t.style='TableStyleLight1';t.showFilterButton=true;
}
function newTable(sheet,title,headers,rows,widths,name,height){
  const last=col(headers.length-1),end=7+rows.length;
  sheet.showGridLines=false;
  sheet.getRange(`A1:${last}${end}`).format.font={name:'Arial',size:10,color:'#172B3A'};
  sheet.getRange('A2').values=[[title]];sheet.getRange('A2').format.font={size:14,bold:true,color:navy};
  sheet.getRange('A2').format.rowHeight=27;
  sheet.getRange('A3:B3').values=[['Cập nhật',day('2026-10-11')]];sheet.getRange('B3').setNumberFormat('dd/mm/yyyy');
  sheet.getRange(`A7:${last}${end}`).values=[headers,...rows];
  const table=sheet.tables.add(`A7:${last}${end}`,true,name);table.style='TableStyleLight1';table.showFilterButton=true;
  sheet.getRange(`A7:${last}7`).format={fill:navy,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:44,wrapText:true,horizontalAlignment:'center',verticalAlignment:'center'};
  sheet.getRange(`A8:${last}${end}`).format={wrapText:true,rowHeight:height,verticalAlignment:'top'};
  for(let row=8;row<=end;row++)if(row%2===1)sheet.getRange(`A${row}:${last}${row}`).format.fill=soft;
  widths.forEach((w,i)=>sheet.getRange(`${col(i)}7:${col(i)}${end}`).format.columnWidth=w);
  sheet.freezePanes.freezeRows(7);sheet.freezePanes.freezeColumns(2);
  return end;
}

for(const name of [...new Set(data.case_plans.map(c=>c.sheet))]){
  const sheet=wb.worksheets.getItem(name),cases=data.case_plans.filter(c=>c.sheet===name);
  const isMoney=name==='Lech tien',status=isMoney?'R':'N',decision=isMoney?'S':'O',mode=isMoney?'T':'P';
  const end=7+cases.length;
  sheet.getRange(`${decision}7:${mode}7`).values=[['Phương án đã chốt 11/10','Xử lý đợt này']];
  sheet.getRange(`${decision}8:${mode}${end}`).values=cases.map(c=>[c.decision,c.mode]);
  sheet.getRange(`${decision}7:${mode}7`).format={fill:'#375F4D',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:46,horizontalAlignment:'center',verticalAlignment:'center'};
  sheet.getRange(`${decision}8:${mode}${end}`).format={fill:teal,font:{name:'Arial',size:10,color:'#172B3A'},wrapText:true,verticalAlignment:'top'};
  sheet.getRange(`${decision}7:${decision}${end}`).format.columnWidth=76;
  sheet.getRange(`${mode}7:${mode}${end}`).format.columnWidth=18;
  sheet.getRange(`${status}7`).values=[['Trạng thái chốt']];
  sheet.getRange(`${status}8:${status}${end}`).conditionalFormats.clear();
  sheet.getRange(`${status}8:${status}${end}`).clear({applyTo:'contents'});
  sheet.getRange(`${status}8:${status}${end}`).formulas=cases.map((c,i)=>{
    const r=i+8;
    return [`=IF(${mode}${r}="Giữ nguyên","Giữ nguyên",IF(AND(${mode}${r}="Điều chỉnh",${decision}${r}<>""),"Đã chốt xử lý","Chưa chốt"))`];
  });
  sheet.getRange('A4').values=[['Phương án cuối tại hai cột xanh, theo file kế toán và xác nhận trực tiếp ngày 11/10/2026.']];
  sheet.getRange('A5').values=[['Cột trạng thái theo quyết định đã chốt; phản hồi/căn cứ gốc vẫn được giữ để đối chiếu.']];
  if(name==='Thanh toan trong')sheet.getRange(`A8:${mode}${end}`).format.rowHeight=178;
  if(name==='Cach ghi dieu chinh')sheet.getRange(`A8:${mode}${end}`).format.rowHeight=140;
  if(name==='So luong can kiem')sheet.getRange(`A8:${mode}${end}`).format.rowHeight=126;
  if(name==='Hang va mo ta')sheet.getRange(`A8:${mode}${end}`).format.rowHeight=126;
  if(isMoney)sheet.getRange(`A8:${mode}${end}`).format.rowHeight=112;
  sheet.getRange(`C8:C${end}`).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(`E8:E${end}`).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(`${isMoney?'Q':'M'}8:${isMoney?'Q':'M'}${end}`).setNumberFormat('dd/mm/yyyy');
  extendTable(sheet,mode,end);
  sheet.freezePanes.freezeRows(7);sheet.freezePanes.freezeColumns(2);
}

const known=wb.worksheets.getItem('Da co can cu');
known.getRange('A2').values=[['Hồ sơ tham khảo – phương án đã chốt']];
known.getRange('A4').values=[['Phương án cuối tại cột I–J: 6 cặp giữ cả chuỗi; 9072 giảm số lượng, giữ nguyên 9186.']];
known.getRange('F7:G7').values=[['Căn cứ/ghi nhận trước đó','Hướng chuẩn bị trước khi chốt']];
known.getRange('I7:J7').values=[['Phương án đã chốt 11/10','Xử lý đợt này']];
known.getRange('I8:J14').values=data.known_plans.map(c=>[c.decision,c.mode]);
known.getRange('H7').format={fill:navy,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:46};
known.getRange('H8:J14').format={wrapText:true,verticalAlignment:'top',rowHeight:90,font:{name:'Arial',size:10}};
known.getRange('I7:J7').format={fill:'#375F4D',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,horizontalAlignment:'center'};
known.getRange('I8:J14').format.fill=teal;
known.getRange('H7:H14').format.columnWidth=53;known.getRange('I7:I14').format.columnWidth=68;known.getRange('J7:J14').format.columnWidth=18;
extendTable(known,'J',14);
known.freezePanes.freezeRows(7);known.freezePanes.freezeColumns(2);
known.getRange('C8:C14').setNumberFormat('dd/mm/yyyy');known.getRange('E8:E14').setNumberFormat('dd/mm/yyyy');

const source=wb.worksheets.getItem('Hoa don lien quan'),targets=new Set(data.actions.map(a=>key(a.target)));
for(let row=8;row<=74;row++){
  const sig=source.getRange(`D${row}`).values[0][0],num=source.getRange(`E${row}`).values[0][0];
  source.getRange(`B${row}`).values=[[targets.has(`${sig}/${String(num).padStart(8,'0')}`)?'Có phương án điều chỉnh':'Giữ nguyên trong đợt này']];
}
source.getRange('A5').values=[['Tiền/ngày nguồn giữ nguyên. Phạm vi xử lý cập nhật theo phương án đã chốt.']];
source.getRange('E8:E74').values=source.getRange('E8:E74').values.map(row=>[String(row[0]).padStart(8,'0')]);
source.getRange('E8:E74').setNumberFormat('@');source.getRange('F8:F74').setNumberFormat('dd/mm/yyyy');
source.freezePanes.freezeRows(7);source.freezePanes.freezeColumns(2);

const actionHeaders=['Mã phương án','Mã hồ sơ','Gốc của chuỗi','Hóa đơn được điều chỉnh','Ngày hóa đơn tham chiếu','Nghiệp vụ đề xuất','Cách thể hiện đã chốt','Tiền hàng đề xuất (đ)','Thuế đề xuất (đ)','Thanh toán đề xuất (đ)','Hàng + thuế − TT (đ)','Tiền hàng nguồn (đ)','Thuế nguồn (đ)','Thanh toán nguồn (đ)','Kiểm phần đảo','Căn cứ tiền'];
const actionRows=data.actions.map(a=>[a.id,a.case,key(a.root),key(a.target),day(a.date),a.action,a.method,...a.money,null,...a.source_money,null,`${a.source_kind}${a.source_xml?' – '+a.source_xml:''}; xem Hoa don lien quan.`]);
const actionEnd=newTable(actionsSheet,'Phương án phát sinh theo xác nhận',actionHeaders,actionRows,[13,12,25,25,16,23,75,20,20,22,21,20,20,22,29,55],'ConfirmedActions',78);
actionsSheet.getRange('A4').values=[['48 phương án trên 27 hóa đơn gốc. Đây là bảng chuẩn bị; chưa nhập, ký hoặc phát hành.']];
actionsSheet.getRange('A5').values=[['8 bộ đảo giữ độ lệch ngược dấu của XML đã phát hành. Không tự tính lại số nguồn.']];
actionsSheet.getRange(`E8:E${actionEnd}`).setNumberFormat('dd/mm/yyyy');
actionsSheet.getRange(`H8:N${actionEnd}`).setNumberFormat(money);
actionsSheet.getRange(`H8:N${actionEnd}`).format.horizontalAlignment='right';
actionsSheet.getRange('K8').formulas=[['=H8+I8-J8']];actionsSheet.getRange(`K8:K${actionEnd}`).fillDown();
actionsSheet.getRange('O8').formulas=[['=IF(F8="Đảo điều chỉnh cũ",IF(AND(H8+L8=0,I8+M8=0,J8+N8=0),"Về 0 từng chỉ tiêu","Chưa về 0"),IF(F8="Giảm số lượng gốc","Chỉ số lượng; nguồn TT trống","Giảm theo gốc"))']];
actionsSheet.getRange(`O8:O${actionEnd}`).fillDown();
actionsSheet.getRange(`K8:K${actionEnd}`).conditionalFormats.add('cellIs',{operator:'notEqual',formula:0,format:{fill:red,font:{bold:true,color:'#9C2424'}}});

const heldRows=data.held_chains.map((h,i)=>{
  const evidence=h.evidence.find(e=>e.reason.includes('Người dùng'))??h.evidence.find(e=>e.reason.startsWith('Đổi'))??h.evidence.find(e=>e.name)??h.evidence[0];
  const related=h.invoices.map(k=>`${key(k)} – ${fmtDay(data.invoice_dates[key(k)])}`).join('\n');
  const rawSource=evidence.source??'';
  const sourceLabel=rawSource.startsWith('data/')?rawSource.split('/').at(-1):rawSource;
  return [i+1,key(h.root),day(h.date),h.already_in_baseline_exceptions?'Ngoại lệ đã tách':'Giữ nguyên theo chốt',related,evidence.name??evidence.reason,key(evidence.invoice),evidence.tax??null,sourceLabel,'Giữ nguyên tất cả hóa đơn liên quan.'];
});
const holdEnd=newTable(holdSheet,'Giữ nguyên toàn bộ chuỗi liên quan',['STT','Hóa đơn gốc','Ngày gốc','Nhóm giữ nguyên','Các hóa đơn liên quan và ngày','Căn cứ nhận diện/giữ nguyên','Hóa đơn có căn cứ','MST/định danh nguồn','Nguồn căn cứ','Hướng xử lý'],heldRows,[8,25,15,25,61,85,25,22,65,45],'ProtectedChains',84);
holdSheet.getRange('A4').values=[['Phạm vi gốc 01/01/2024–30/06/2025; giữ cả các hóa đơn liên quan được thấy trong nguồn mới đến 09/10/2026.']];
holdSheet.getRange('A5').values=[['238 chuỗi ngoại lệ cũ và6 cặp chốt giữ nguyên. 3 cặp cùng sản phẩm giữ riêng tại Hang va mo ta.']];
holdSheet.getRange(`C8:C${holdEnd}`).setNumberFormat('dd/mm/yyyy');
holdSheet.getRange(`H8:H${holdEnd}`).setNumberFormat('@');
for(let i=0;i<heldRows.length;i++)if(heldRows[i][4].split('\n').length>4)holdSheet.getRange(`A${i+8}:J${i+8}`).format.rowHeight=Math.min(260,heldRows[i][4].split('\n').length*15);

const summary=wb.worksheets.getItem('Tong hop');
summary.getRange('A2').values=[['UVG – Phương án đã chốt với kế toán']];
summary.getRange('A4').values=[['Đã chốt 29 hồ sơ và 7 hồ sơ tham khảo theo phản hồi và xác nhận ngày 11/10/2026.']];
summary.getRange('A5').values=[['Hai cột xanh tại các sheet chi tiết thể hiện phương án cuối. Không còn điểm 1–6 chờ xác nhận.']];
summary.getRange('D7:F7').values=[['Chưa chốt','Đã chốt xử lý','Giữ nguyên']];
summary.getRange('G7').values=[['Phạm vi xử lý đã chốt']];
summary.getRange('G8:G12').values=[['Đảo số đã phát hành; chuẩn bị giảm gốc.'],['Giữ chú thích, bỏ dòng 0 đã chỉ rõ; dòng 555 chỉ đảo tiền.'],['Gốc 6/16 chỉ giảm số lượng; đảo riêng 655/657.'],['Hai cặp khác hàng có phương án; ba cặp cùng sản phẩm giữ nguyên.'],['Giảm gốc, đảo điều chỉnh cũ đúng dấu theo từng cặp số.']];
const groups=[...new Set(data.case_plans.map(c=>c.sheet))];
groups.forEach((name,i)=>{
  const count=data.case_plans.filter(c=>c.sheet===name).length,status=name==='Lech tien'?'R':'N';
  ['Chưa chốt','Đã chốt xử lý','Giữ nguyên'].forEach((v,j)=>summary.getRange(`${col(3+j)}${i+8}`).formulas=[[`=COUNTIFS('${name}'!${status}8:${status}${7+count},"${v}")`]]);
});
summary.getRange('A14').values=[['Tổng 29 hồ sơ']];
summary.getRange('A17:G25').clear({applyTo:'contents'});
summary.getRange('A17').values=[['7 hồ sơ tham khảo']];
summary.getRange('C17:F17').formulas=[['=COUNTA(\'Da co can cu\'!A8:A14)','=0','=COUNTIFS(\'Da co can cu\'!J8:J14,"Điều chỉnh")','=COUNTIFS(\'Da co can cu\'!J8:J14,"Giữ nguyên")']];
summary.getRange('A19').values=[['Tổng tất cả hồ sơ']];summary.getRange('C19:F19').formulas=[['=SUM(C14,C17)','=SUM(D14,D17)','=SUM(E14,E17)','=SUM(F14,F17)']];
summary.getRange('A19:G19').format={fill:'#E5EDF4',font:{bold:true},rowHeight:29};
summary.getRange('A21').values=[['Phương án phát sinh']];summary.getRange('C21').formulas=[[`=COUNTA('Phuong an phat sinh'!A8:A${actionEnd})`]];
summary.getRange('G21').values=[['27 gốc: 21 đảo điều chỉnh cũ và 27 giảm gốc.']];
summary.getRange('A23').values=[['Chuỗi giữ nguyên theo quy tắc']];summary.getRange('C23').formulas=[[`=COUNTA('Giu nguyen ca chuoi'!A8:A${holdEnd})`]];
summary.getRange('G23').values=[['238 ngoại lệ cũ + 6 cặp đã chốt giữ nguyên.']];
summary.getRange('A25').values=[['Rà nguồn hiện có: chưa thấy hóa đơn thuộc chuỗi giữ nguyên trong 10.394 hóa đơn đang có ở sheet import.']];
summary.getRange('A27').values=[['Đã chuẩn bị phương án; ba workbook chính và dữ liệu nguồn chưa thay đổi.']];
summary.getRange('C8:F23').setNumberFormat('0');
for(const [name,view] of Object.entries(data.source_views)){
  const sheet=wb.worksheets.getItem(name);sheet.freezePanes.unfreeze();
  if(view.freeze_panes){
    const [,letters,row]=view.freeze_panes.match(/^([A-Z]+)(\d+)$/);
    const column=[...letters].reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0);
    if(Number(row)>1)sheet.freezePanes.freezeRows(Number(row)-1);
    if(column>1)sheet.freezePanes.freezeColumns(column-1);
  }
  if(typeof view.show_grid_lines==='boolean')sheet.showGridLines=view.show_grid_lines;
}

wb.recalculate();
for(const sheet of wb.worksheets.items)sheet.getRange('B3').setNumberFormat('dd/mm/yyyy');
const testSheet=wb.worksheets.getItem('Lech tien');
testSheet.getRange('T8').values=[['Giữ nguyên']];
wb.recalculate();
assert.deepEqual(summary.getRange('C14:F14').values[0],[29,0,25,4]);
testSheet.getRange('T8').values=[['Điều chỉnh']];
wb.recalculate();
assert.deepEqual(summary.getRange('C14:F14').values[0],[29,0,26,3]);
assert.deepEqual(summary.getRange('C19:F19').values[0],[36,0,27,9]);
assert.equal(summary.getRange('C21').values[0][0],48);
assert.equal(summary.getRange('C23').values[0][0],244);
for(let i=0;i<data.actions.length;i++){
  assert.equal(actionsSheet.getRange(`K${i+8}`).values[0][0],data.actions[i].difference);
  if(data.actions[i].action==='Đảo điều chỉnh cũ')assert.equal(actionsSheet.getRange(`O${i+8}`).values[0][0],'Về 0 từng chỉ tiêu');
}
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:80},maxChars:1800});
await fs.writeFile(path.join(dir,'confirmed-errors.ndjson'),errors.ndjson);
for(const [i,name] of wb.worksheets.items.map(s=>s.name).entries()){
  let range=name==='Tong hop'?'A1:G27':name==='Lech tien'?'J6:T9':name==='Da co can cu'?'F6:J10':name==='Hoa don lien quan'?'A6:P10':name==='Phuong an phat sinh'?'D6:P10':name==='Giu nguyen ca chuoi'?'A6:J10':'F6:P9';
  const blob=await wb.render({sheetName:name,range,scale:1,format:'png'});
  await fs.writeFile(path.join(dir,'confirmed-previews',`after-${i+1}.png`),new Uint8Array(await blob.arrayBuffer()));
}
await fs.mkdir(path.dirname(outputPath),{recursive:true});
const file=await SpreadsheetFile.exportXlsx(wb);await file.save(outputPath);
await fs.writeFile(path.join(dir,'confirmed-output.json'),JSON.stringify({output:outputPath,summary:[36,0,27,9],actions:48,heldChains:244},null,2));
console.log(JSON.stringify({output:outputPath,sheets:wb.worksheets.items.length,summary:summary.getRange('C19:F19').values,errors:errors.ndjson}));
