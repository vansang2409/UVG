import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const out=path.dirname(fileURLToPath(import.meta.url));const d=JSON.parse(await fs.readFile(path.join(out,'five-deep-findings.json'),'utf8'));
const wb=Workbook.create();const s=wb.worksheets.add('Kết luận 5 cặp');const it=wb.worksheets.add('Chi tiết 10 HĐ');const am=wb.worksheets.add('AMIS đối chiếu');
const fmt='#,##0;[Red](#,##0);0';const cc=n=>{let z='';for(;n;n=Math.floor((n-1)/26))z=String.fromCharCode(65+(n-1)%26)+z;return z;};
const iso=x=>new Date(x+'T00:00:00Z');const vn=x=>{const [a,b,c]=x.split('/');return new Date(`${c}-${b.padStart(2,'0')}-${a.padStart(2,'0')}T00:00:00Z`)};
const short=x=>x.startsWith('KT đã chốt')?'KT chốt xuất trùng':x.startsWith('Khả năng cao')?'Nghi xuất lại cho CTY':'Chưa đủ căn cứ';
function base(sh,last,end,hr,widths){
 sh.showGridLines=false;sh.getRange(`A1:${last}${end}`).format.font={name:'Arial',size:10,color:'#202124'};sh.getRange(`A1:${last}${end}`).format.verticalAlignment='center';sh.getRange(`A1:A${end}`).format.rowHeight=25;
 sh.getRange('A2').format.font={name:'Arial',size:15,bold:true,color:'#202124'};sh.getRange('A3').format.font.italic=true;
 sh.getRange(`A${hr}:${last}${hr}`).format={fill:'#253B53',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:42,horizontalAlignment:'center',verticalAlignment:'center'};
 widths.forEach((v,i)=>sh.getRange(cc(i+1)+'1').format.columnWidthPx=v);sh.freezePanes.freezeRows(hr);sh.freezePanes.freezeColumns(1);
}
s.getRange('A2').values=[['Kiểm tra kỹ 5 cặp HĐ khách lẻ / công ty']];s.getRange('A3').values=[['Bằng chứng từ nguồn 13/11/2026 và bảng kế toán đọc live.']];
s.getRange('A5:C5').values=[['Mức chứng cứ','Số cặp','Thanh toán HĐ khách lẻ (đ)']];s.getRange('A6:A9').values=[['KT chốt xuất trùng'],['Nghi xuất lại cho CTY'],['Chưa đủ căn cứ'],['Tổng 5 cặp']];
const rows=d.summary_rows.map(r=>r.map((v,i)=>i===3?short(v):[5,6].includes(i)?iso(v):v));
const heads=['HĐ khách lẻ còn tiền','HĐ công ty đối chiếu','Thanh toán (đ)','Kết luận','BHH liên quan','Ngày HĐ khách lẻ','Ngày HĐ công ty','Đối chiếu dòng hàng','Công ty nhận HĐ','Mã đơn liên quan','Ý kiến KT hiện có','Lưu ý / phần chưa đủ căn cứ','Vị trí bằng chứng'];
s.getRange('A12:M12').values=[heads];s.getRange('A13:M17').values=rows;s.tables.add('A12:M17',true,'FiveCases');
s.getRange('A20').values=[['Điểm cần lưu ý']];s.getRange('A22').values=[['900/987: hai HĐ khớp 5 dòng, nhưng AMIS phân bổ tiền KA600/KA100 khác 57.771đ mỗi bên.']];s.getRange('A23').values=[['6823/7014: hai HĐ khớp 2 dòng; AMIS chia tiền máy/lõi, trong khi HĐ chỉ ghi tiền ở máy.']];s.getRange('A24').values=[['20896/11733: chưa xác nhận quan hệ đổi sang HĐ công ty; ngày lập cách 124 ngày.']];s.getRange('A25').values=[['Chưa sửa dữ liệu AMIS hoặc thực hiện xử lý hóa đơn trong lượt kiểm tra này.']];
s.getRange('A28').values=[['Nguồn meInvoice']];s.getRange('B28').values=[['Bộ ZIP tổng/chi tiết/điều chỉnh/thay thế trong data/1311/meinvoice; chi tiết nằm ở tab Chi tiết 10 HĐ.']];
s.getRange('A29').values=[['Nguồn đối chiếu']];s.getRange('B29').values=[['AMIS 2025 trong data/1311/amis; MISA_HD_GOC; Đối soát NB; MEINV F0, F1, F2 (Ktoan).']];
s.getRange('A30').values=[['Bảng kế toán']];s.getRange('B30').values=[['https://docs.google.com/spreadsheets/d/1C5hoPDvZKB1vGFSUG9rzChcORt2L9NnsqETVHxPui9I']];
base(s,'M',30,12,[235,235,170,230,195,110,110,350,560,225,520,760,650]);s.tabColor='#253B53';
s.getRange('A5:C5').format={fill:'#253B53',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:40,horizontalAlignment:'center'};
s.getRange('C6:C9').setNumberFormat(fmt);s.getRange('C13:C17').setNumberFormat(fmt);s.getRange('F13:G17').setNumberFormat('dd/mm/yyyy');s.getRange('A9:C9').format.font.bold=true;
s.getRange('D13:D17').conditionalFormats.addCustom('$D13="KT chốt xuất trùng"',{fill:'#FCE7E7',font:{color:'#9C2626'}});s.getRange('D13:D17').conditionalFormats.addCustom('$D13="Chưa đủ căn cứ"',{fill:'#FFF1D6',font:{color:'#875D0F'}});
it.getRange('A2').values=[['Chi tiết gốc của 10 HĐ']];it.getRange('A3').values=[['Giữ riêng ô tiền trống; tên hàng, số lượng và tiền lấy từ bảng kê chi tiết 13/11.']];
it.getRange('A6:M6').values=[['BHH liên quan','Bên','Ký hiệu / số HĐ','Ngày HĐ','Tên hàng nguyên văn','Số lượng','ĐVT','Đơn giá','Tiền hàng','VAT','Thanh toán','Thuế suất','File / dòng nguồn']];
const ir=d.item_rows.map(r=>r.map((v,i)=>i===3?vn(v):v));it.getRange('A7:M26').values=ir;it.tables.add('A6:M26',true,'InvoiceItems');base(it,'M',26,6,[190,110,235,105,620,95,90,150,155,140,165,100,650]);it.getRange('D7:D26').setNumberFormat('dd/mm/yyyy');it.getRange('F7:F26').setNumberFormat('#,##0.##');it.getRange('H7:K26').setNumberFormat(fmt);it.getRange('E7:E26').format.wrapText=true;it.getRange('A7:A26').format.rowHeight=42;
am.getRange('A2').values=[['Các dòng AMIS đang dùng để đối chiếu']];am.getRange('A3').values=[['Cặp 900 và 6823 có cách phân bổ tiền máy/lõi khác với dòng HĐ; tổng tiền và VAT vẫn khớp.']];am.getRange('A6:L6').values=[['BHH','SKU AMIS','Tên hàng AMIS','Số lượng','ĐVT','Đơn giá AMIS','Tiền hàng AMIS','VAT AMIS','Thanh toán AMIS','HĐ đang map AMIS','Sheet / dòng AMIS','Lưu ý phân bổ tiền']];am.getRange('A7:L13').values=d.amis_rows;am.tables.add('A6:L13',true,'AmisItems');base(am,'L',13,6,[190,240,620,95,90,155,160,140,170,235,400,800]);am.getRange('D7:D13').setNumberFormat('#,##0.##');am.getRange('F7:I13').setNumberFormat(fmt);am.getRange('C7:C13').format.wrapText=true;am.getRange('A7:A13').format.rowHeight=42;
// Viết công thức tổng hợp sau cùng để tránh tính lại khi nhập các phần bằng chứng.
for(let r=6;r<=8;r++)s.getRange(`B${r}:C${r}`).formulas=[[`=COUNTIFS(D13:D17,A${r})`,`=SUMIFS(C13:C17,D13:D17,A${r})`]];
s.getRange('B9:C9').formulas=[['=SUM(B6:B8)','=SUM(C6:C8)']];wb.recalculate();
const expected=[[1,1239000],[3,7876581],[1,606100],[5,9721681]];if(JSON.stringify(s.getRange('B6:C9').values)!==JSON.stringify(expected))throw new Error('Summary mismatch');
console.log((await wb.inspect({kind:'region',sheetId:'Kết luận 5 cặp',range:'A5:C9',maxChars:1800,tableMaxRows:5,tableMaxCols:3})).ndjson);
for(const [name,range] of [['Kết luận 5 cặp','A1:E18'],['Chi tiết 10 HĐ','A1:I10'],['AMIS đối chiếu','A1:I10']]){const p=await wb.render({sheetName:name,range,scale:1,format:'png'});await fs.writeFile(path.join(out,'deep-preview-'+name+'.png'),new Uint8Array(await p.arrayBuffer()));}
const file=await SpreadsheetFile.exportXlsx(wb);await file.save(path.join(out,'KIEM_TRA_KY_5_CAP_HD_1311.xlsx'));console.log('EXPORTED deep five');
