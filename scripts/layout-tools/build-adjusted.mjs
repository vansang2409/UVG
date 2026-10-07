import fs from 'node:fs/promises';
import { Workbook,SpreadsheetFile } from '@oai/artifact-tool';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'../..');
const out=path.join(root,'outputs/hoa-don-dieu-chinh-2024');
const d=JSON.parse(await fs.readFile(path.join(root,'scripts/data/adjusted_2024_review.json'),'utf8'));
const w=Workbook.create();
const s=w.worksheets.add('Dieu chinh 2024');
const h=[];
for(let i=0;i<d.summary.max_blocks;i++)h.push(...(i===0?['Ký hiệu','Số hóa đơn','Ngày','Tiền hàng','Thuế','Thanh toán']:[`Ký hiệu sửa ${i}`,`Số hóa đơn sửa ${i}`,`Ngày sửa ${i}`,`Tiền hàng sửa ${i}`,`Thuế sửa ${i}`,`Thanh toán sửa ${i}`]));
h.push('Cộng dồn tiền hàng','Cộng dồn thuế','Cộng dồn thanh toán','Cần kiểm tra');
function col(n){let r='';for(n++;n;n=Math.floor((n-1)/26))r=String.fromCharCode(65+(n-1)%26)+r;return r;}
function date(t){const[a,b,c]=t.split('/').map(Number);return new Date(Date.UTC(c,b-1,a));}
const rows=d.groups.map(g=>{const r=[];for(let i=0;i<d.summary.max_blocks;i++){const b=g.invoices[i];r.push(...(b?[b[0],b[1],date(b[2]),...b.slice(4)]:Array(6).fill(null)));}r.push(...g.sums,g.notes||'');return r;});
const last=rows.length+2,end=col(h.length-1);
s.showGridLines=false;s.getRange(`A2:${end}2`).values=[h];s.getRange(`A3:${end}${last}`).values=rows;
s.tables.add(`A2:${end}${last}`,true,'DieuChinh2024');
s.getRange(`A1:${end}${last}`).format.font.name='Arial';s.getRange(`A1:${end}${last}`).format.font.size=10;
s.getRange(`A1:${end}${last}`).format.verticalAlignment='center';
s.getRange(`A1:${end}2`).format.fill='#243746';s.getRange(`A1:${end}2`).format.font.color='#FFFFFF';s.getRange(`A1:${end}2`).format.font.bold=true;
s.getRange(`A2:${end}2`).format.wrapText=true;s.getRange(`A1:${end}1`).format.rowHeight=26;s.getRange(`A2:${end}2`).format.rowHeight=42;
s.getRange(`A1:${end}2`).format.horizontalAlignment='center';s.getRange(`A3:${end}${last}`).format.rowHeight=28;
for(let i=0;i<d.summary.max_blocks;i++){
 const start=i*6,color=['#243746','#315F77','#526579'][i]||'#526579';
 for(let j=0;j<6;j++)s.getRangeByIndexes(0,start+j,last,1).format.columnWidth=[16,14,16,20,18,21][j];
 const band=`${col(start)}1:${col(start+5)}1`;s.mergeCells(band);s.getRange(band).values=[[`F${i} — ${i===0?'Hóa đơn gốc':'Điều chỉnh'}`]];
 s.getRangeByIndexes(0,start,2,6).format.fill=color;
 s.getRangeByIndexes(2,start+1,last-2,1).setNumberFormat('00000000');
 s.getRangeByIndexes(2,start+2,last-2,1).setNumberFormat('dd/mm/yyyy');
 s.getRangeByIndexes(2,start+3,last-2,3).setNumberFormat('#,##0;[Red](#,##0);0');
}
const sumStart=6*d.summary.max_blocks;
s.mergeCells(`${col(sumStart)}1:${end}1`);s.getRange(`${col(sumStart)}1`).values=[['Giá trị cộng dồn và kiểm tra']];
s.getRangeByIndexes(0,sumStart,last,3).format.columnWidth=25;s.getRangeByIndexes(2,sumStart,last-2,3).setNumberFormat('#,##0;[Red](#,##0);0');
s.getRangeByIndexes(0,sumStart+3,last,1).format.columnWidth=105;
s.freezePanes.freezeRows(2);s.freezePanes.freezeColumns(2);
const g=w.worksheets.add('Huong dan');g.showGridLines=false;
const notes=[['UVG – rà soát hóa đơn đã bị điều chỉnh, gốc năm 2024'],
 [`${d.summary.roots} hóa đơn gốc, ${d.summary.invoices} hóa đơn trong chuỗi, ${d.summary.roots_with_notes} chuỗi cần kiểm tra thêm.`],
 ['Đã loại chuỗi thuộc nhóm ngoại lệ công ty/MST và chuỗi có hóa đơn thay thế.'],
 ['F0 = gốc. F1/F2 = hóa đơn điều chỉnh liên quan; xem bảng Chi tiet nguon để biết tham chiếu, không mặc định F2 sửa F1.'],
 ['Cộng dồn theo bảng tổng. Thiếu tiền giữ trống; riêng gốc 574 tạm tính thanh toán trống = 0 theo xác nhận người dùng.'],
 ['Cộng dồn là số theo file hiện có; chưa xác minh XML/lịch sử đầy đủ. Không phải số tiền đã chốt để phát hành.'],
 [`Lần điều chỉnh muộn nhất trong nhóm này: ${d.summary.max_date}. Bộ nguồn chưa cập nhật đến hiện tại.`],
 ['Mẫu số 3: cộng dồn theo nguồn là 11.856.288 đồng; hai hóa đơn mẫu cộng -11.856.288 đồng, khớp về số học để đưa chuỗi về 0.'],
 ['File mẫu chỉ là ví dụ nhập liệu, không được tính là hóa đơn đã phát hành. Chưa tạo file import, ký hay phát hành.'],
 ['Đơn vị: đồng. Dữ liệu nguồn giữ nguyên, 11 file đã xác minh SHA-256.']];
g.getRange('A1:A10').values=notes;g.getRange('A1:A10').format.columnWidth=145;g.getRange('A1:A10').format.rowHeight=28;g.getRange('A1').format.font.bold=true;
const t=w.worksheets.add('Chi tiet nguon');t.showGridLines=false;
const th=['Ký hiệu gốc','Số gốc','Ký hiệu','Số hóa đơn','Ngày','Loại / tham chiếu','Trạng thái nguồn','Tiền hàng','Thuế','Thanh toán','Vị trí nguồn'];
t.getRange('A1:K1').values=[th];t.getRange(`A2:K${d.detail.length+1}`).values=d.detail.map(r=>[...r.slice(0,4),date(r[4]),...r.slice(5)]);
t.tables.add(`A1:K${d.detail.length+1}`,true,'ChiTietNguon');
for(let i=0;i<11;i++)t.getRangeByIndexes(0,i,d.detail.length+1,1).format.columnWidth=[16,14,16,14,16,54,35,20,18,20,100][i];
t.getRange(`B2:B${d.detail.length+1}`).setNumberFormat('00000000');t.getRange(`D2:D${d.detail.length+1}`).setNumberFormat('00000000');t.getRange(`E2:E${d.detail.length+1}`).setNumberFormat('dd/mm/yyyy');t.getRange(`H2:J${d.detail.length+1}`).setNumberFormat('#,##0;[Red](#,##0);0');t.freezePanes.freezeRows(1);
w.recalculate();
const destination=process.argv[2]||path.join(out,'HOA_DON_DA_DIEU_CHINH_2024_BO_SUNG_TEN.xlsx');
await fs.mkdir(path.dirname(destination),{recursive:true});
await(await SpreadsheetFile.exportXlsx(w)).save(destination);
console.log(JSON.stringify(d.summary));
