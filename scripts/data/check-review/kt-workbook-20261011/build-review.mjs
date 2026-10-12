import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const dir = path.dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, '$1'));
const data = JSON.parse(await fs.readFile(path.join(dir, 'review-data.json'), 'utf8'));
const rootDir = path.resolve(dir, '../../../..');
const outputDir = path.join(rootDir, 'outputs', 'kt-review-20261011-01a12992');
const outputPath = path.join(outputDir, 'UVG_CAC_VAN_DE_CAN_KE_TOAN_XAC_NHAN_20261011.xlsx');
const key = x => x.join('/');
const num = x => x == null ? null : Number(x);
const day = x => x == null ? null : new Date(`${String(x).slice(0,10)}T00:00:00Z`);
const fmtDay = x => String(x).slice(0,10).split('-').reverse().join('/');
const money = '#,##0;[Red](#,##0);0';
const xmls = new Map(data.xmls.map(x => [key(x.key), x]));
const pending = data.audit.pending;
const cases = [];
const groups = [
  {name:'Lech tien', title:'Lệch tiền hàng, thuế và thanh toán', count:8, prefix:'LT', question:'Chốt ba khoản tiền đúng, nguyên nhân chênh và hướng xử lý.'},
  {name:'So luong can kiem', title:'Số lượng bằng 0 hoặc để trống', count:8, prefix:'SL', question:'Phân biệt dòng hàng và dòng diễn giải. Chốt số lượng, tính chất dòng.'},
  {name:'Thanh toan trong', title:'Thanh toán gốc để trống', count:2, prefix:'TT', question:'Xác nhận thanh toán gốc bằng 0 hay đang thiếu số liệu.'},
  {name:'Hang va mo ta', title:'Khác loại hàng hoặc mô tả', count:5, prefix:'HH', question:'Xác định hàng thực tế, mã hàng và cách xử lý nội dung điều chỉnh.'},
  {name:'Cach ghi dieu chinh', title:'Cách ghi số lượng và giá trị điều chỉnh', count:6, prefix:'DC', question:'Chốt giảm giá, giảm số lượng hay giảm một khoản tổng.'},
];
function add(group, chain, focus, obs, question) {
  const idx = cases.filter(x=>x.group===group).length+1;
  const g = groups.find(x=>x.name===group);
  const root = chain.root ?? chain.key;
  const h = chain.invoices?.find(x=>key(x.key)===key(root))?.header ?? chain.header ?? chain.original_header;
  const c = {id:`${g.prefix}${String(idx).padStart(2,'0')}`, group, root, date:h.date, focus:focus?.key??null,
    focusDate:focus?.header?.date??null, obs, question, chain, focusRecord:focus};
  cases.push(c); return c;
}
const arith = pending.flatMap(c=>c.invoices.filter(i=>i.header_arithmetic_difference!=null && num(i.header_arithmetic_difference)!==0).map(i=>({c,i})));
arith.sort((a,b)=>Math.abs(num(b.i.header_arithmetic_difference))-Math.abs(num(a.i.header_arithmetic_difference)));
for (const {c,i} of arith) {
  const x=xmls.get(key(i.key));
  assert(x && x.money.every((v,j)=>num(v)===num(i.header.money[j])));
  let obs='Tổng tiền trong XML khớp bảng kê nhưng tiền hàng + thuế khác thanh toán. ';
  if (i.key[1]==='00000649') obs+='Tiền hàng và thuế âm, thanh toán dương. ';
  if (x.lines.some(l=>num(l.goods)===0) && num(x.money[0])!==0) obs+='Thành tiền dòng XML bằng 0 trong khi tổng tiền hàng khác 0. ';
  if(c.invoices.length>2) obs+='Chuỗi còn điều chỉnh 1C25TUV/00002646 ngày 31/05/2025; xem Hoa don lien quan.';
  add('Lech tien',c,i,obs,'Xác định tiền hàng, thuế và thanh toán đúng, nguyên nhân chênh, căn cứ và hướng xử lý. Kiểm cả phần dòng và phần tổng. Nếu đã có chứng từ xử lý tiếp, ghi số/ngày.');
}
for (const c of pending.filter(c=>c.zero_or_blank_qty_lines.length)) {
  const l=c.zero_or_blank_qty_lines[0];
  const obs=c.root[1]==='00000555'
    ? 'Dòng màng lọc HEPA H13 KA350 FT (BS HĐ 547): số lượng và đơn giá trống. Tiền hàng 46.296, thuế 3.705, thanh toán dòng 50.001. Hóa đơn gốc có 6 dòng.'
    : 'Dòng màng/lõi lọc HEPA H13 KA600 FT: số lượng, đơn giá và tiền dòng trống trong bảng xuất mới. Hóa đơn gốc có 48 dòng. Không tự coi các ô trống là 0.';
  add('So luong can kiem',c,c.invoices.find(i=>i.role==='adjustment'),obs,
    'Đối chiếu đúng dòng hàng trên XML gốc/bản hiển thị hoặc chứng từ bán. Chốt số lượng, đơn giá, tiền dòng và đây là hàng bán, khuyến mãi hay diễn giải. Ghi nguyên nhân thiếu số liệu và hướng chuẩn bị phần giảm chi tiết.');
}
for (const r of data.extra) {
  const l=r.zero_lines[0];
  const isNote=l.name.includes('không thu tiền');
  const current=r.current_details.find(x=>String(x.name).trim()===l.name.trim());
  assert(current);
  const currentText=`Bảng xuất mới: SL ${current.qty??'trống'}, đơn giá ${current.price??'trống'}, thành tiền ${current.goods??'trống'}.`;
  const obs=`Nhóm chưa điều chỉnh. Toàn bộ ${r.lines.length} dòng đã được giữ ngoài import vì dòng “${l.name.trim()}” lưu SL/đơn giá/thành tiền = 0 ở bảng kiểm. ${currentText}`;
  add('So luong can kiem',r,null,obs,isNote
    ? 'Xác nhận dòng này là chú thích hàng tặng hay một mặt hàng có số lượng thực tế. Chốt tính chất và số lượng cần thể hiện khi chuẩn bị điều chỉnh toàn bộ hóa đơn. Không tự bỏ dòng hoặc đưa lại vào import.'
    : 'Đối chiếu mặt hàng tặng KM002/KM003 với chứng từ bán hoặc XML gốc. Chốt số lượng thực tế và tính chất dòng, giải thích vì sao SL/đơn giá/thành tiền bằng 0. Xác định cách thể hiện khi chuẩn bị điều chỉnh.');
}
for (const c of pending.filter(c=>c.root[0]==='1C25TUV' && ['00000006','00000016'].includes(c.root[1]))) {
  add('Thanh toan trong',c,c.invoices.find(i=>i.role==='adjustment'),
    'Bảng tổng của hóa đơn gốc ghi tiền hàng = 0, thuế = 0 nhưng thanh toán để trống. Chưa có XML gốc trong phần nguồn đã lưu hiện tại để xác nhận giá trị thanh toán. Điều chỉnh liên quan đã có XML.',
    'Thanh toán gốc đúng là 0 hay bảng xuất đang thiếu? Nếu thiếu, chốt giá trị đúng và căn cứ. Nếu đúng 0, xác nhận rõ cho riêng hóa đơn này. Không áp dụng quy ước của gốc 1C24TUV/00000574 sang đây.');
}
for (const r of data.audit.review_11.filter(r=>['Khác loại hàng','Khác mô tả'].includes(r.category))) {
  const a=r.original_detail[0], b=r.xml[0].lines[0];
  add('Hang va mo ta',r,{key:r.adjustment,header:r.adjustment_header},
    `${r.category}. Gốc: ${a.name.trim()}. Điều chỉnh: ${b.name.trim()}. Cộng dồn tiền hàng, thuế, thanh toán đã bằng 0.`,
    r.category==='Khác loại hàng'
      ? 'Xác định hàng thực tế bán/trả, tên và mã hàng đúng. Bên nào ghi sai giữa gốc và điều chỉnh? Chốt hướng xử lý lần điều chỉnh cũ và dữ liệu cần chuẩn bị tiếp.'
      : 'Hai mô tả có cùng mã hàng/sản phẩm thực tế không? Xác nhận rõ cùng hàng hoặc chỉ ra khác biệt. Chốt mô tả đúng và hướng xử lý nội dung điều chỉnh.');
}
for (const r of data.audit.review_11.filter(r=>!['Khác loại hàng','Khác mô tả'].includes(r.category))) {
  const l=r.xml[0].lines[0];
  add('Cach ghi dieu chinh',r,{key:r.adjustment,header:r.adjustment_header},
    r.category==='Giảm bằng đơn giá âm'
      ? `${r.original_detail[0].name.trim()}. Hóa đơn gốc SL +1; điều chỉnh vẫn SL +1, đơn giá âm. Tổng tiền đã về 0, số lượng hai phía chưa triệt tiêu.`
      : 'Gốc có 1 đèn bắt muỗi KM300. Điều chỉnh giảm một khoản tổng: SL/đơn giá = 0, thành tiền dòng -499.074, thuế -39.926, thanh toán -539.000. Tiền dòng khớp phần tổng và chuỗi đã về 0.',
    r.category==='Giảm bằng đơn giá âm'
      ? 'Nghiệp vụ thực tế là giảm giá hay giảm/trả toàn bộ số lượng hàng? Chốt mục tiêu xử lý số lượng và giá trị, căn cứ và hướng xử lý điều chỉnh cũ. Đơn giá âm chưa tự chứng minh nội dung sai.'
      : 'Nghiệp vụ thực tế là giảm một khoản tiền hay giảm toàn bộ 1 sản phẩm KM300? Chốt cách thể hiện nội dung/số lượng đúng với mục tiêu xử lý tiếp và căn cứ.');
}
assert.equal(cases.length,29);
assert.equal(new Set(cases.map(c=>key(c.root))).size,29);
for(const g of groups) assert.equal(cases.filter(c=>c.group===g.name).length,g.count);

const resolved = pending.filter(c=>!cases.some(i=>key(i.root)===key(c.root))).map((c,i)=>{
  const focus=c.invoices.find(i=>i.role==='adjustment');
  let found, next;
  if(c.root[1]==='00000003') {
    found='Phương án đảo khoản giảm 00000508 và giảm đầy đủ chi tiết gốc đã được người dùng chốt từ 09/10/2026. XML đã có; phần dòng khác tổng đã được ghi nhận.';
    next='Tiếp tục chuẩn bị/cập nhật dữ liệu theo phương án đã chốt. Không yêu cầu chốt lại cùng phương án.';
  } else if(c.root[0]==='1C24TUV' && c.root[1]==='00000574') {
    found='Người dùng đã chốt tạm tính thanh toán gốc bằng 0. Ô bảng tổng vẫn trống. Quy ước chỉ áp dụng cho hóa đơn này.';
    next='Tiếp tục chuẩn bị dữ liệu theo quy ước riêng đã có. Không tự đổi ô nguồn hoặc mở rộng quy ước.';
  } else if(focus.key[1]==='00002140') {
    found='XML 00002140 có thuế suất tổng hợp 8%. Tiền hàng 17.830.586, thuế 1.426.447, thanh toán 19.257.033. Không có danh sách dòng hàng.';
    next='Dùng bằng chứng XML để chuẩn bị dữ liệu tiếp. Không yêu cầu tìm lại thuế suất chỉ vì bảng chi tiết không thể hiện.';
  } else {
    const x=xmls.get(key(focus.key));
    found=focus.key[1]==='00009186'
      ? 'XML điều chỉnh tên hàng có 2 dòng TChat=4; ba tổng tiền đều 0. Bảng kê để trống thanh toán. Nội dung đổi tên hàng đã có trong XML.'
      : 'XML điều chỉnh thông tin người mua; ba tổng tiền đều 0. Bảng kê để trống thanh toán. Nội dung thông tin người mua nằm trong ghi chú tham chiếu XML.';
    next='Giữ nội dung điều chỉnh thông tin và phân biệt 0 trong XML với ô trống bảng kê khi chuẩn bị dữ liệu tiếp.';
  }
  return {id:`CC${String(i+1).padStart(2,'0')}`,root:c.root,date:c.invoices[0].header.date,focus:focus.key,focusDate:focus.header.date,found,next,chain:c};
});
assert.equal(resolved.length,7);

const records=[];
function sourceLabel(h){return `${h.evidence.file.split(/[\\/]/).at(-1).trim()}, dòng ${h.evidence.row}`;}
for(const c of [...cases,...resolved]) {
  let invs;
  if(c.chain.invoices) invs=c.chain.invoices;
  else if(c.chain.adjustment) invs=[{key:c.root,role:'original',header:c.chain.original_header},{key:c.chain.adjustment,role:'adjustment',header:c.chain.adjustment_header}];
  else invs=[{key:c.root,role:'original',header:c.chain.header}];
  for(const i of invs) {
    const x=xmls.get(key(i.key));
    let text=x?.related_note??'';
    if(x?.lines?.length) text += (text?'\n':'') + x.lines.map(l=>`${l.name.trim()} (TChat=${l.nature}, SL=${l.qty??'trống'}, đơn giá=${l.price??'trống'}, thành tiền=${l.goods??'trống'})`).join('\n');
    if(x && !x.lines.length) text += (text?'\n':'')+'XML không có danh sách dòng hàng.';
    if(x?.tax_breakdown?.length) text += '\nThuế suất tổng hợp: '+x.tax_breakdown.map(t=>t.rate).join(', ');
    if(c.chain.zero_lines) text += (text?'\n':'')+c.chain.zero_lines.map(l=>`${l.name.trim()}: SL/đơn giá/thành tiền lưu = 0 ở bảng kiểm.`).join('\n');
    if(c.chain.zero_or_blank_qty_lines && i.role==='original') text += (text?'\n':'')+c.chain.zero_or_blank_qty_lines.map(l=>`${l.name.trim()}: SL=${l.qty??'trống'}, đơn giá=${l.price??'trống'}, tiền hàng=${l.goods??'trống'}, thuế=${l.tax??'trống'}, thanh toán=${l.payment??'trống'}`).join('\n');
    let scope=x?'Có XML trong bộ nguồn hiện tại.':'Chưa có XML trong phần nguồn đã lưu hiện tại.';
    if(x && i.header.money.some((v,j)=>v==null && num(x.money[j])===0)) scope+=' XML ghi 0 tại ô bảng kê để trống.';
    records.push({case:c.id,classification:c.group?'Cần kế toán phản hồi':'Đã có căn cứ/quy ước',role:i.role==='original'?'Hóa đơn gốc':'Điều chỉnh liên quan',key:i.key,date:i.header.date,money:i.header.money.map(num),xmlMoney:x?.money?.map(num)??[null,null,null],text,source:sourceLabel(i.header),xml:x?.file??null,scope});
  }
}
assert.equal(records.length,67);
assert.equal(new Set(records.map(r=>key(r.key))).size,67);
assert.equal(records.filter(r=>r.classification==='Cần kế toán phản hồi').length,53);

const wb=Workbook.create();
const summary=wb.worksheets.add('Tong hop');
for(const g of groups) wb.worksheets.add(g.name);
const known=wb.worksheets.add('Da co can cu');
const src=wb.worksheets.add('Hoa don lien quan');
const palette={navy:'#23374D',ink:'#172B4D',muted:'#52606D',line:'#D9E2EC',input:'#FFF4CC',soft:'#F3F6FA',red:'#FDE7E7'};
function col(n){let s='';for(++n;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;}
function setup(s,title,lastCol,lastRow){
  s.showGridLines=false;
  s.getRange(`A1:${lastCol}${lastRow}`).format.font={name:'Arial',size:10,color:palette.ink};
  s.getRange(`A1:${lastCol}${lastRow}`).format.verticalAlignment='center';
  s.getRange('A2').values=[[title]]; s.getRange('A2').format.font={size:14,bold:true,color:palette.navy};
  s.getRange('A2').format.rowHeight=27;
  s.getRange('A3').values=[['Cập nhật']];s.getRange('B3').values=[[day('2026-10-11')]];s.getRange('B3').setNumberFormat('dd/mm/yyyy');
  s.getRange(`A6:${lastCol}6`).format.borders={bottom:{style:'thin',color:palette.navy}};
}
function table(s,headers,rows,widths,name,height=112){
  const end=7+rows.length,last=col(headers.length-1);
  s.getRange(`A7:${last}${end}`).values=[headers,...rows];
  const t=s.tables.add(`A7:${last}${end}`,true,name);t.style='TableStyleLight1';t.showFilterButton=true;
  s.getRange(`A7:${last}7`).format={fill:palette.navy,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,horizontalAlignment:'center',verticalAlignment:'center',rowHeight:46,borders:{insideVertical:{style:'thin',color:'#FFFFFF'}}};
  s.getRange(`A8:${last}${end}`).format.wrapText=true;
  s.getRange(`A8:${last}${end}`).format.rowHeight=height;
  s.getRange(`A8:${last}${end}`).format.verticalAlignment='top';
  for(let r=8;r<=end;r++) if(r%2===1)s.getRange(`A${r}:${last}${r}`).format.fill=palette.soft;
  widths.forEach((w,i)=>s.getRange(`${col(i)}7:${col(i)}${end}`).format.columnWidth=w);
  s.freezePanes.freezeRows(7);s.freezePanes.freezeColumns(2);
  return end;
}
const controls={};
const responseHeaders=['Nguyên nhân (KT điền)','Hướng xử lý (KT điền)','Số liệu/nội dung đúng (KT điền)','Căn cứ/chứng từ (KT điền)','Người phản hồi','Ngày phản hồi','Theo dõi phản hồi'];
for(const g of groups){
  const s=wb.worksheets.getItem(g.name), list=cases.filter(c=>c.group===g.name), isMoney=g.name==='Lech tien';
  const headers=['Mã hồ sơ','Hóa đơn gốc','Ngày gốc','HĐ điều chỉnh liên quan','Ngày điều chỉnh',...(isMoney?['Tiền hàng XML (đ)','Thuế XML (đ)','Thanh toán XML (đ)','Hàng + thuế − thanh toán (đ)']:[]),'Ghi nhận cần kiểm tra','Câu hỏi cần kế toán kết luận',...responseHeaders];
  const rows=list.map(c=>[c.id,key(c.root),day(c.date),c.focus?key(c.focus):null,day(c.focusDate),...(isMoney?[...xmls.get(key(c.focus)).money.map(num),null]:[]),c.obs,c.question,null,null,null,null,null,null,null]);
  const end=7+rows.length,last=col(headers.length-1),start=isMoney?11:7,status=col(start+6);
  setup(s,g.title,last,end);s.tabColor=isMoney?'#A43838':'#496B8A';
  s.getRange('A4').values=[['Ô vàng: kế toán điền nguyên nhân, hướng xử lý, số liệu/nội dung đúng và căn cứ.']];
  s.getRange('A5').values=[['Số/ngày của mọi hóa đơn trong chuỗi và số tiền nguồn nằm tại sheet Hoa don lien quan.']];
  s.getRange('A4:A5').format.font={color:palette.muted,size:10};
  const widths=[11,25,14,25,14,...(isMoney?[18,18,19,21]:[]),57,65,40,40,40,40,21,15,25];
  table(s,headers,rows,widths,`Review${g.prefix}`,72);
  s.getRange(`C8:C${end}`).setNumberFormat('dd/mm/yyyy');s.getRange(`E8:E${end}`).setNumberFormat('dd/mm/yyyy');
  s.getRange(`B8:B${end}`).setNumberFormat('@');s.getRange(`D8:D${end}`).setNumberFormat('@');
  if(isMoney){
    s.getRange(`F8:I${end}`).setNumberFormat(money);s.getRange(`F8:I${end}`).format.horizontalAlignment='right';
    s.getRange('I8').formulas=[['=F8+G8-H8']];s.getRange(`I8:I${end}`).fillDown();
    s.getRange(`I8:I${end}`).conditionalFormats.add('cellIs',{operator:'notEqual',formula:0,format:{fill:palette.red,font:{bold:true,color:'#9C2424'}}});
  }
  const first=col(start),basis=col(start+3),who=col(start+4),datecol=col(start+5);
  s.getRange(`${first}7:${datecol}7`).format.fill='#806525';
  s.getRange(`${first}8:${datecol}${end}`).format.fill=palette.input;
  s.getRange(`${first}8:${datecol}${end}`).format.font={color:'#1D4ED8',size:10};
  s.getRange(`${datecol}8:${datecol}${end}`).setNumberFormat('dd/mm/yyyy');
  s.getRange(`${status}8`).formulas=[[`=IF(COUNTA(${first}8:${datecol}8)=0,"Chưa phản hồi",IF(AND(${first}8<>"",${col(start+1)}8<>"",${col(start+2)}8<>"",${basis}8<>""),"Đủ trường phản hồi","Cần bổ sung"))`]];
  s.getRange(`${status}8:${status}${end}`).fillDown();
  s.getRange(`${status}8:${status}${end}`).conditionalFormats.add('containsText',{text:'Cần bổ sung',format:{fill:palette.red,font:{color:'#9C2424'}}});
  controls[g.name]={start,end,status,first,basis,datecol};
}

setup(summary,'UVG – Hồ sơ cần kế toán làm rõ','G',25);summary.tabColor=palette.navy;
summary.getRange('A4').values=[['Mỗi dòng là một hồ sơ/chuỗi. 29 hồ sơ cần phản hồi, liên quan 53 hóa đơn.']];
summary.getRange('A5').values=[['Đủ trường phản hồi nghĩa là đã điền 4 mục chính; chưa đồng nghĩa đã xử lý hoặc nhập hóa đơn.']];
const summaryRows=groups.map(g=>[g.title,g.name,null,null,null,null,g.question]);
const sumEnd=table(summary,['Nhóm vấn đề','Sheet chi tiết','Hồ sơ','Chưa phản hồi','Cần bổ sung','Đủ trường phản hồi','Việc cần kế toán chốt'],summaryRows,[43,26,12,19,19,24,70],'ReviewSummary',36);
summary.freezePanes.unfreeze();
groups.forEach((g,i)=>{
  const r=i+8,{end,status}=controls[g.name];
  summary.getRange(`C${r}`).formulas=[[`=COUNTA('${g.name}'!A8:A${end})`]];
  ['Chưa phản hồi','Cần bổ sung','Đủ trường phản hồi'].forEach((v,j)=>summary.getRange(`${col(3+j)}${r}`).formulas=[[`=COUNTIFS('${g.name}'!${status}8:${status}${end},"${v}")`]]);
});
summary.getRange('A14').values=[['Tổng cần phản hồi']];
summary.getRange('C14:F14').formulas=[['=SUM(C8:C12)','=SUM(D8:D12)','=SUM(E8:E12)','=SUM(F8:F12)']];
summary.getRange('A14:G14').format={fill:'#E5EDF4',font:{bold:true},rowHeight:29};
summary.getRange('C8:F14').setNumberFormat('0');
summary.getRange('A17').values=[['Phạm vi bổ sung']];summary.getRange('A17').format.font={bold:true};
summary.getRange('A18').values=[['Đã đưa thêm 6 hóa đơn chưa điều chỉnh có dòng SL 0 đang được giữ ngoài import vào nhóm số lượng.']];
summary.getRange('A20').values=[['Đã có căn cứ/quy ước']];summary.getRange('A20').format.font={bold:true};
summary.getRange('A21').values=[['7 hồ sơ ở sheet Da co can cu được giữ riêng để không hỏi lại thông tin hoặc quyết định đã có.']];
summary.getRange('A23').values=[['Phạm vi đợt này']];summary.getRange('A23').format.font={bold:true};
summary.getRange('A24').values=[['Các chuỗi thay thế chưa được mở xử lý trong đợt điều chỉnh này. File này tổng hợp các vấn đề đang chờ đã xác định.']];

setup(known,'Các hồ sơ đã có căn cứ hoặc quyết định','G',14);known.tabColor='#788594';
known.getRange('A4').values=[['Giữ riêng để tham khảo khi xử lý tiếp; không yêu cầu kế toán xác nhận lại cùng thông tin.']];
table(known,['Mã hồ sơ','Hóa đơn gốc','Ngày gốc','HĐ điều chỉnh liên quan','Ngày điều chỉnh','Nội dung đã làm rõ','Việc chuẩn bị tiếp'],resolved.map(c=>[c.id,key(c.root),day(c.date),key(c.focus),day(c.focusDate),c.found,c.next]),[11,25,14,25,14,70,65],'KnownCases',72);
known.getRange('C8:C14').setNumberFormat('dd/mm/yyyy');known.getRange('E8:E14').setNumberFormat('dd/mm/yyyy');

setup(src,'Hóa đơn liên quan và căn cứ đối chiếu','P',74);
src.getRange('A4').values=[['Đơn vị tiền: đồng. Ô trống được giữ trống; số 0 là giá trị có trong nguồn.']];
src.getRange('A5').values=[['Nguồn: bảng kê xuất lại 09/10/2026, XML đã nhận, hai workbook hiện tại và quyết định đã ghi nhận.']];
table(src,['Mã hồ sơ','Phạm vi công việc','Vai trò','Ký hiệu','Số hóa đơn','Ngày hóa đơn','Tiền hàng bảng kê (đ)','Thuế bảng kê (đ)','Thanh toán bảng kê (đ)','Tiền hàng XML (đ)','Thuế XML (đ)','Thanh toán XML (đ)','Nội dung dòng/ghi chú cần đối chiếu','Nguồn bảng kê','File XML','Ghi chú độ phủ nguồn'],records.map(r=>[r.case,r.classification,r.role,r.key[0],r.key[1],day(r.date),...r.money,...r.xmlMoney,r.text,r.source,r.xml,r.scope]),[11,25,25,15,17,15,20,19,22,20,19,22,85,65,37,55],'InvoiceEvidence',72);
src.getRange('E8:E74').setNumberFormat('@');src.getRange('F8:F74').setNumberFormat('dd/mm/yyyy');src.getRange('G8:L74').setNumberFormat(money);
src.getRange('G8:L74').format.horizontalAlignment='right';
for(let i=0;i<records.length;i++){
  const lines=records[i].text.split('\n').reduce((s,l)=>s+Math.ceil(l.length/100),0);
  if(lines>5)src.getRange(`A${i+8}:P${i+8}`).format.rowHeight=Math.min(390,Math.max(72,lines*13));
}

// Check the editable response workflow with temporary inputs, then restore blanks.
const firstSheet=wb.worksheets.getItem('Lech tien');
assert.equal(firstSheet.getRange('R8').values[0][0],'Chưa phản hồi');
firstSheet.getRange('L8').values=[['Kiểm tra thử']];
assert.equal(firstSheet.getRange('R8').values[0][0],'Cần bổ sung');
firstSheet.getRange('M8:O8').values=[['Hướng xử lý thử','Nội dung thử','Căn cứ thử']];
assert.equal(firstSheet.getRange('R8').values[0][0],'Đủ trường phản hồi');
firstSheet.getRange('L8:Q8').clear({applyTo:'contents'});
wb.recalculate();
assert.equal(summary.getRange('C14:F14').values[0].join(','),'29,29,0,0');
for(const c of cases.filter(c=>c.group==='Lech tien')){
  const rn=8+cases.filter(x=>x.group==='Lech tien').findIndex(x=>x.id===c.id);
  assert.equal(firstSheet.getRange(`I${rn}`).values[0][0],num(c.focusRecord.header_arithmetic_difference));
}
const errorScan=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:100},maxChars:2500});
await fs.mkdir(outputDir,{recursive:true});
await fs.mkdir(path.join(dir,'previews'),{recursive:true});
await fs.writeFile(path.join(dir,'formula-errors.ndjson'),errorScan.ndjson);
for(const [i,s] of [summary,...groups.map(g=>wb.worksheets.getItem(g.name)),known,src].entries()){
  const range=s===summary?'A1:G25':s===src?'A2:P11':s===known?'A2:G10':s.name==='Lech tien'?'A2:R10':'A2:N10';
  const png=await wb.render({sheetName:s.name,range,scale:1,format:'png'});
  await fs.writeFile(path.join(dir,'previews',`${i+1}.png`),new Uint8Array(await png.arrayBuffer()));
}
const file=await SpreadsheetFile.exportXlsx(wb);await file.save(outputPath);
await fs.writeFile(path.join(dir,'review-manifest.json'),JSON.stringify({output:outputPath,cases:cases.map(c=>({id:c.id,group:c.group,root:c.root,focus:c.focus,date:c.date,focusDate:c.focusDate})),resolved:resolved.map(c=>({id:c.id,root:c.root,focus:c.focus})),related:records.map(r=>({id:r.case,key:r.key,date:r.date,money:r.money,xmlMoney:r.xmlMoney})),controls,workbookHashes:data.workbook_hashes},null,2));
console.log(JSON.stringify({output:outputPath,sheets:8,activeCases:29,resolvedReferences:7,relatedInvoices:67,pendingRelatedInvoices:53,summary:summary.getRange('C14:F14').values,formulaErrors:errorScan.ndjson},null,2));
