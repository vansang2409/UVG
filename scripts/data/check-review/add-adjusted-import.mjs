import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const h='D:/UVG/scripts/data/check-review';
const d=JSON.parse(await fs.readFile(`${h}/adjusted-import-data.json`,'utf8'));
console.log('READ BASE');const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${h}/before-adjusted-import.xlsx`));
const s=w.worksheets.add('Import dieu chinh'),p=w.worksheets.add('Cho KT xac nhan');
const headers=['Số thứ tự hóa đơn (*)','Ngày hóa đơn','Tên khách hàng','Địa chỉ','Mã số thuế','Người mua hàng','Email','Hình thức thanh toán (*)','HĐ bị điều chỉnh thuộc hệ thống khác','Ký hiệu HĐ bị điều chỉnh (*)','Số hóa đơn bị điều chỉnh (*)','Ngày hóa đơn bị điều chỉnh','Mã của CQT trên HĐ bị điều chỉnh','Lý do điều chỉnh','Tên hàng hóa/dịch vụ (*)','ĐVT','Số lượng','Đơn giá','Thành tiền','Tổng tiền hàng','Thuế suất GTGT (%)','Tổng tiền thuế GTGT','Tổng tiền thanh toán','Hàng khuyến mãi'];
const text=v=>v===null||v===''?null:String(v).trim(),neg=v=>v===null?null:-v,rate=v=>Number(String(v).trim().replace('%',''));
const date=v=>{const [y,m,dd]=v.slice(0,10).split('-');return `${dd}/${m}/${y}`;};
const rows=[],meta=[];let seq=0;
for(const x of d.ready){const root=x.key.join('/'),rd=d.headers[root],lines=x.lines,common=lines[0];let slot=4;
 for(const op of x.old.slice(1)){
  seq++;const key=op.key.join('/'),hd=d.headers[key],sl=d.detail[key],tr=sl.find(l=>l.rate!==null&&l.rate!=='').rate;
  const desc=`Điều chỉnh ${hd.money[0]<=0?'tăng':'giảm'} để đảo khoản đã điều chỉnh trên HĐ ${key}; tham chiếu HĐ gốc ${root}`;
  rows.push([seq,new Date(Date.UTC(2026,9,8)),text(common.buyer),text(common.address),text(common.taxid),text(common.buyer),text(common.email),text(common.method),null,...x.key,date(rd.date),null,desc,desc,null,null,null,-hd.money[0],-hd.money[0],rate(tr),-hd.money[1],-hd.money[2],null]);
  meta.push({row:rows.length+9,root:x.key,source_key:op.key,source:hd.source,kind:'Đảo khoản điều chỉnh cũ',slot:slot++,seq,first:true});
 }
 seq++;const reason=`Điều chỉnh giảm hàng bán HĐ số ${x.key[1]} ngày ${date(rd.date)} do xuất sai thông tin`;
 for(let i=0;i<lines.length;i++){const l=lines[i],first=i===0;
  rows.push([seq,new Date(Date.UTC(2026,9,8)),text(l.buyer),text(l.address),text(l.taxid),text(l.buyer),text(l.email),text(l.method),null,...x.key,date(rd.date),null,reason,text(l.name),text(l.unit),-l.qty,l.price,-l.goods,first?-rd.money[0]:null,first?rate(lines.find(l=>l.rate!==null&&l.rate!=='').rate):null,first?-rd.money[1]:null,first?-rd.money[2]:null,l.goods===0?1:null]);
  meta.push({row:rows.length+9,root:x.key,source_key:x.key,source:l.source,kind:'Giảm toàn bộ chi tiết gốc',slot,seq,first});
 }
}
if(seq!==1600||rows.length!==40777)throw new Error('Unexpected import count');
const instructions=['File mẫu danh sách hóa đơn điều chỉnh để nhập vào phần mềm','Hướng dẫn:','- Ngày hóa đơn mới 08/10/2026; sửa đồng nhất các dòng cùng số thứ tự nếu cần trước khi import.','- 791 chuỗi đủ dữ liệu: 1.600 hóa đơn mới, 40.777 dòng. Mỗi khoản cũ được đảo riêng, sau đó giảm toàn bộ chi tiết gốc.','- Tổng tiền hàng, thuế suất, tổng thuế, tổng thanh toán chỉ ghi dòng đầu mỗi hóa đơn; lấy số tổng từ bảng kê tổng, không thay bằng phép nhân SL × đơn giá.','- Tất cả hóa đơn mới đều tham chiếu F0. Dòng đảo khoản cũ ghi rõ hóa đơn cũ trong lý do và tên dịch vụ; SL, đơn giá để trống theo mẫu.','- 19 chuỗi chờ kế toán ở sheet Cho KT xac nhan; 237 chuỗi đã bằng 0 chưa đưa vào import. Thành tiền dòng gốc 0 thì Hàng khuyến mãi = 1.'];
for(let i=0;i<7;i++){s.mergeCells(`A${i+1}:X${i+1}`);s.getRange(`A${i+1}`).values=[[instructions[i]]];}
s.getRange('A9:X9').values=[headers];console.log('WRITE',rows.length);
for(let i=0;i<rows.length;i+=1000){const part=rows.slice(i,i+1000);s.getRange(`A${i+10}:X${i+9+part.length}`).values=part;}
console.log('FORMAT');const last=rows.length+9;s.showGridLines=false;s.getRange('A1:X7').format={font:{name:'Arial',size:10},wrapText:true,rowHeight:27};s.getRange('A1').format.font.bold=true;
s.getRange('A9:X9').format={fill:'#243746',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:54};s.getRange(`A10:X${last}`).format.rowHeight=62;
const widths=[14,16,42,52,23,42,30,24,18,21,24,20,25,85,100,12,18,22,24,24,19,24,26,18];widths.forEach((v,i)=>s.getRangeByIndexes(0,i,1,1).format.columnWidth=v);
s.getRange(`C10:C${last}`).format.wrapText=true;s.getRange(`N10:O${last}`).format.wrapText=true;s.getRange(`B10:B${last}`).setNumberFormat('dd/mm/yyyy');s.getRange(`E10:E${last}`).setNumberFormat('@');s.getRange(`K10:K${last}`).setNumberFormat('00000000');s.getRange(`Q10:W${last}`).setNumberFormat('#,##0;[Red](#,##0);0');s.getRange(`A10:A${last}`).setNumberFormat('0');s.getRange(`X10:X${last}`).setNumberFormat('0');s.freezePanes.freezeRows(9);s.freezePanes.freezeColumns(1);
const pendingHeaders=['Ký hiệu F0','Số hóa đơn F0','Ngày F0','Cộng dồn tiền hàng','Cộng dồn thuế','Cộng dồn thanh toán','Tình trạng','Cần kế toán xác nhận','Các hóa đơn điều chỉnh cũ','Số dòng gốc','Dòng SL 0/trống','Nguồn chi tiết gốc'];
function refs(lines){const files=new Map();for(const l of lines){const [name,n]=l.source.split(' | dòng ');if(!files.has(name))files.set(name,[]);files.get(name).push(Number(n));}return [...files].map(([name,ns])=>`${name} | dòng ${Math.min(...ns)}–${Math.max(...ns)} (${ns.length} dòng)`).join('; ');}
p.mergeCells('A1:L1');p.getRange('A1').values=[['19 chuỗi chưa đưa vào import — giữ riêng để kế toán xác nhận']];p.getRange('A2:L2').values=[pendingHeaders];
p.getRange('A3:L21').values=d.pending.map(x=>[...x.key,date(d.headers[x.key.join('/')].date),...x.prior,x.state==='unknown'?'Còn ô tiền trống':'Chưa về 0',x.reasons.join('; '),x.old.slice(1).map(op=>op.key.join('/')).join('; '),x.lines.length,x.lines.filter(l=>l.qty===null||l.qty===0).length,refs(x.lines)]);
p.showGridLines=false;p.getRange('A1:L21').format={font:{name:'Arial',size:10},verticalAlignment:'center',rowHeight:88};p.getRange('A1:L2').format={fill:'#243746',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true};p.getRange('A1').format.rowHeight=30;p.getRange('A2:L2').format.rowHeight=48;
[18,19,19,26,24,28,28,120,60,18,20,100].forEach((v,i)=>p.getRangeByIndexes(0,i,1,1).format.columnWidth=v);p.getRange('H3:L21').format.wrapText=true;p.getRange('B3:B21').setNumberFormat('00000000');p.getRange('D3:F21').setNumberFormat('#,##0;[Red](#,##0);0');p.freezePanes.freezeRows(2);p.freezePanes.freezeColumns(2);const t=p.tables.add('A2:L21',true,'ChoKeToanXacNhan');t.style='TableStyleMedium2';t.showFilterButton=true;
for(let i=0;i<d.pending.length;i++)p.getRange(`A${i+3}:L${i+3}`).format.rowHeight=Math.max(88,Math.ceil(d.pending[i].reasons.join('; ').length/140)*14+12);
const g=w.worksheets.getItem('Huong dan');g.getRange('A13:A16').values=[['Import dieu chinh: 791 chuỗi / 1.600 hóa đơn mới / 40.777 dòng; đảo từng khoản điều chỉnh cũ theo bảng tổng rồi giảm từng dòng gốc.'],['Cho KT xac nhan: 19 chuỗi giữ riêng vì ghi chú đang chờ kiểm tra, tiền trống/lệch, số lượng gốc 0 hoặc thuế suất chưa xác định. 237 chuỗi đã bằng 0 chưa đưa vào đợt này.'],['Các F4–F7 ở bảng chính vẫn là kế hoạch dự kiến; import mới chỉ lấy 791 chuỗi đã đủ dữ liệu. Ngày hóa đơn mới 08/10/2026, Tên khách hàng bằng Người mua hàng.'],['Tổng tiền mới lấy số đảo dấu P/Q/R của bảng tổng; thành tiền hàng gốc lấy số đảo dấu đúng nguồn. Chưa import, ký hoặc phát hành.']];g.getRange('A13:A16').format={font:{name:'Arial',size:10},wrapText:true,rowHeight:48,columnWidth:145};
w.recalculate();console.log('PREVIEW');for(const [sheetName,range,file] of [['Import dieu chinh','J9:X15','adjusted-import-money.png'],['Import dieu chinh','A9:I12','adjusted-import-buyer.png'],['Cho KT xac nhan','G1:L5','adjusted-import-pending.png']]){const b=await w.render({sheetName,range,scale:1,format:'png'});await fs.writeFile(`${h}/${file}`,new Uint8Array(await b.arrayBuffer()));}
console.log('EXPORT');await(await SpreadsheetFile.exportXlsx(w)).save(`${h}/with-adjusted-import.xlsx`);await fs.writeFile(`${h}/adjusted-import-meta.json`,JSON.stringify({rows,meta}));console.log(JSON.stringify({chains:d.ready.length,invoices:seq,lines:rows.length,pending:d.pending.length,staged:true}));
