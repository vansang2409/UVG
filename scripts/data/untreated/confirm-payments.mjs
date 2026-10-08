import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const h='D:/UVG/scripts/data/untreated';
const keys=['1C24TUV/00000671','1C24TUV/00000744','1C25TUV/00000391','1C25TUV/00000761','1C25TUV/00000835','1C25TUV/00004748'];
const d=JSON.parse(await fs.readFile(`${h}/untreated.json`,'utf8'));
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${h}/before-payment-confirmation.xlsx`));
const s=w.worksheets.getItem('Chua dieu chinh'),g=w.worksheets.getItem('Huong dan');
const before=await w.render({sheetName:'Chua dieu chinh',range:'M37:P39',scale:1,format:'png'});await fs.writeFile(`${h}/payment-before.png`,new Uint8Array(await before.arrayBuffer()));
const changes=[];
for(let i=0;i<d.groups.length;i++){
 const x=d.groups[i],key=x.root.join('/');if(!keys.includes(key))continue;
 const row=i+3,counts=x.line_counts;
 if(x.invoices[0][6]!==null||counts.zero_payment!==counts.lines)throw new Error(`Confirmation evidence changed: ${key}`);
 if(s.getRange(`A${row}:B${row}`).values[0].join('/')!==key)throw new Error('Wrong invoice row');
 const note='Đã chốt thanh toán tổng = 0 theo xác nhận người dùng ngày 08/10/2026. Bảng kê tổng nguyên bản để trống thanh toán; lấy số 0 ghi rõ trong dòng chi tiết. Tiền hàng = 0, thuế = 0; giữ nguyên dữ liệu nguồn. Quy tắc khuyến mãi: thành tiền dòng gốc = 0 thì nhập 1. Căn cứ: '+x.detail_source+'.';
 s.getRange(`F${row}`).values=[[0]];s.getRange(`M${row}`).values=[['Cả 3 khoản bằng 0']];s.getRange(`N${row}`).values=[[note]];
 s.getRange(`A${row}:P${row}`).format.rowHeight=Math.max(s.getRange(`N${row}`).format.rowHeight||38,Math.ceil(note.length/125)*14+10);
 changes.push({key,row,note,confirmed_on:'08/10/2026',payment:0,source_payment:null,evidence:x.detail_source});
}
if(changes.length!==6)throw new Error('Expected exactly 6');
g.getRange('A7').values=[['6 F0 có thanh toán bảng tổng nguyên bản trống đã được chốt thanh toán = 0 theo số 0 trong chi tiết và xác nhận người dùng ngày 08/10/2026; ghi căn cứ ở cột N, không sửa nguồn.']];
g.getRange('A9').values=[['Lọc M = Cả 3 khoản bằng 0 hoặc N chứa Đã chốt thanh toán tổng để xem 6 trường hợp đã xác nhận; lọc L > 0 để xem hóa đơn có dòng thành tiền bằng 0.']];
w.recalculate();
const b=await w.render({sheetName:'Chua dieu chinh',range:'M37:P39',scale:1,format:'png'});await fs.writeFile(`${h}/payment-after.png`,new Uint8Array(await b.arrayBuffer()));
await (await SpreadsheetFile.exportXlsx(w)).save(`${h}/confirmed-payment.xlsx`);
await fs.writeFile(`${h}/confirmed-payments.json`,JSON.stringify(changes,null,2));
console.log(JSON.stringify({confirmed:changes.map(x=>x.key),staged:true}));
