import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';

const root='D:/UVG',work=`${root}/scripts/data/btl-bhh-live-20261026`,out=process.env.UVG_BTL_OUTPUT_DIR||`${root}/outputs/btl-bhh-20261026`;
const p=JSON.parse(await fs.readFile(`${work}/prepared.json`,'utf8'));
const file=`${out}/LIEN_KET_BTL_BHH_2024.xlsx`;
await fs.mkdir(out,{recursive:true});
if(!process.argv.includes('--replace'))try{await fs.access(file);throw new Error('Không ghi đè file đã có');}catch(e){if(e.code!=='ENOENT')throw e;}
const w=Workbook.create();
let tableId=0;
function sheet(name,headers,rows,widths){
  const s=w.worksheets.add(name),n=rows.length+1,last=String.fromCharCode(64+headers.length);
  s.showGridLines=false;s.getRange(`A1:${last}1`).values=[headers];
  if(rows.length)s.getRangeByIndexes(1,0,rows.length,headers.length).values=rows;
  s.getRange(`A1:${last}${n}`).format={font:{name:'Arial',size:10},rowHeight:26,verticalAlignment:'center'};
  s.getRange(`A1:${last}1`).format={fill:'#24435C',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:44,wrapText:true,verticalAlignment:'center'};
  widths.forEach((width,i)=>s.getRangeByIndexes(0,i,1,1).format.columnWidth=width);
  if(rows.length){const t=s.tables.add(`A1:${last}${n}`,true,`UVG_${++tableId}_BTL`);t.style='TableStyleMedium2';t.showFilterButton=true;}
  s.freezePanes.freezeRows(1);s.freezePanes.freezeColumns(3);
  return s;
}
const rows=p.details.map(r=>r.map((v,i)=>i===1?new Date(v):v));
const s=sheet('Chi tiet BTL-BHH',['Số phiếu BTL','Ngày hạch toán','Số CT bán hàng gốc BHH','Dòng BTL','Mã hàng','Tên hàng','ĐVT','Số lượng trả','Đơn giá (VND)','Tiền hàng trả (VND)','Thuế suất GTGT (%)','Tiền thuế GTGT (VND)','Tổng tiền trả (VND)','Khách hàng','Liên kết trên AMIS','Đối chiếu sổ AMIS 2024','Dòng sổ AMIS'],rows,[20,18,29,13,25,85,12,18,25,25,21,25,27,34,27,32,24]);
const end=rows.length+1;
s.getRange(`B2:B${end}`).setNumberFormat('dd/mm/yyyy');
for(const c of ['A','C','E'])s.getRange(`${c}2:${c}${end}`).setNumberFormat('@');
s.getRange(`H2:I${end}`).setNumberFormat('#,##0.00########');
for(const c of ['J','L','M'])s.getRange(`${c}2:${c}${end}`).setNumberFormat('#,##0;[Red](#,##0);0');
s.getRange(`F2:F${end}`).format.wrapText=true;s.getRange(`A2:Q${end}`).format.rowHeight=42;
s.getRange(`M2:M${end}`).formulas=rows.map((r,i)=>[`=J${i+2}+L${i+2}`]);
s.getRange(`C2:C${end}`).conditionalFormats.add('containsBlanks',{format:{fill:'#FFF2CC'}});
const docs=p.docs.map(d=>[d.btl,new Date(d.date),d.customer,d.lines,d.bhhs.join(', '),null,null,null,null,d.payment,null,d.source_ok?'Khớp sổ AMIS 2024':'Khác sổ AMIS 2024']);
const d=sheet('Phieu BTL',['Số phiếu BTL','Ngày hạch toán','Khách hàng','Số dòng hàng','BHH gốc đã xác định','Số lượng trả','Tiền hàng trả (VND)','VAT trả (VND)','Tổng chi tiết (VND)','Tổng phiếu AMIS (VND)','Chênh lệch (VND)','Đối chiếu sổ 2024'],docs,[20,18,34,17,48,20,27,25,28,28,25,32]);
const de=docs.length+1;
d.getRange(`B2:B${de}`).setNumberFormat('dd/mm/yyyy');d.getRange(`A2:A${de}`).setNumberFormat('@');d.getRange(`F2:F${de}`).setNumberFormat('#,##0.00');d.getRange(`G2:K${de}`).setNumberFormat('#,##0;[Red](#,##0);0');
d.getRange(`E2:E${de}`).format.wrapText=true;
d.getRange(`A2:L${de}`).format.rowHeight=38;
for(let i=0;i<p.docs.length;i++){
 const doc=p.docs[i],r=i+2;
 d.getRange(`F${r}:I${r}`).formulas=[[`=SUM('Chi tiet BTL-BHH'!H${doc.first}:H${doc.last})`,`=SUM('Chi tiet BTL-BHH'!J${doc.first}:J${doc.last})`,`=SUM('Chi tiet BTL-BHH'!L${doc.first}:L${doc.last})`,`=SUM('Chi tiet BTL-BHH'!M${doc.first}:M${doc.last})`]];
 d.getRange(`K${r}`).formulas=[[`=I${r}-J${r}`]];
}
d.getRange(`K2:K${de}`).conditionalFormats.add('cellIs',{operator:'notEqual',formula:0,format:{fill:'#FCE4D6',font:{color:'#9C0006'}}});
const info=sheet('Thong tin',['Nội dung','Giá trị'],[
 ['Đơn vị',p.company],['Dữ liệu kế toán',p.database],['Kỳ lấy phiếu BTL',p.filter],['Số phiếu BTL',null],['Số dòng hàng trả',null],['Số BHH gốc khác nhau',p.summary.unique_bhh],['Số dòng trống BHH',p.summary.missing_bhh_lines],['Tiền hàng trả (VND)',null],['VAT trả (VND)',null],['Tổng tiền trả (VND)',null],['Phiếu khác sổ AMIS 2024',p.summary.source_mismatch_documents],['Nguồn đọc','Bán hàng → Trả lại hàng bán → Chi tiết → Số CT bán hàng'],['Phương pháp','Đọc trực tiếp AMIS, kiểm đủ dòng/tổng tiền và đối chiếu sổ AMIS 2024. Bổ sung riêng BTL00145 → BHHT2024024972 theo xác nhận người dùng 27/10/2026; giữ nguyên snapshot đọc AMIS ban đầu.'],['Ý nghĩa BHH','BHH lấy từ AMIS; riêng BTL00145 do người dùng xác nhận. Nguồn liên kết ghi rõ ở cột Liên kết trên AMIS. Không suy đoán theo tên/giá.'],['Nguồn sổ đối chiếu','So_chi_tiet_ban_hang_2024.xlsx, sheet SỔ CHI TIẾT BÁN HÀNG; số dòng gốc ở cột cuối bảng chi tiết.'],['Ngày trích dữ liệu','26/10/2026'],['Thao tác trên AMIS','Chỉ xem, lọc và đọc dữ liệu. Không sửa, ghi sổ, bỏ ghi hay xóa chứng từ.'],['Liên kết bổ sung','BTL00145 / SP000042 → BHHT2024024972 do người dùng xác nhận 27/10/2026. Snapshot AMIS lúc trích vẫn trống. SL, đơn giá, tiền hàng/VAT/thanh toán khớp hoàn toàn, trả hết.']
],[34,125]);
info.freezePanes.unfreeze();info.getRange('B2:B19').format.wrapText=true;info.getRange('A2:B19').format.rowHeight=42;
for(const [row,formula] of [[5,`=COUNTA('Phieu BTL'!A2:A${de})`],[6,`=COUNTA('Chi tiet BTL-BHH'!A2:A${end})`],[9,`=SUM('Chi tiet BTL-BHH'!J2:J${end})`],[10,`=SUM('Chi tiet BTL-BHH'!L2:L${end})`],[11,`=SUM('Chi tiet BTL-BHH'!M2:M${end})`]])info.getRange(`B${row}`).formulas=[[formula]];
info.getRange('B9:B11').setNumberFormat('#,##0;[Red](#,##0);0');
const missing=rows.filter(r=>!r[2]);
if(missing.length){
 const m=sheet('Chua co BHH',['Số phiếu BTL','Ngày hạch toán','BHH gốc đang trống','Dòng BTL','Mã hàng','Tên hàng','ĐVT','Số lượng trả','Đơn giá (VND)','Tiền hàng trả (VND)','Thuế suất GTGT (%)','Tiền thuế GTGT (VND)','Tổng tiền trả (VND)','Khách hàng','Tình trạng','Đối chiếu sổ AMIS 2024','Dòng sổ AMIS'],missing,[20,18,29,13,25,85,12,18,25,25,21,25,27,34,27,32,24]);
 m.getRange('B2').setNumberFormat('dd/mm/yyyy');m.getRange('C2').format.fill='#FFF2CC';m.getRange('F2').format.wrapText=true;m.getRange('A2:Q2').format.rowHeight=42;m.getRange('H2:I2').setNumberFormat('#,##0.00########');m.getRange('J2:M2').setNumberFormat('#,##0');m.getRange('M2').formulas=[['=J2+L2']];
}
w.recalculate();
console.log((await w.inspect({kind:'region',sheetId:'Thong tin',range:'A5:B12',maxChars:1800,tableMaxRows:8,tableMaxCols:2})).ndjson);
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!',options:{useRegex:true,maxResults:20},maxChars:1000})).ndjson);
for(const [name,range,suffix] of [['Chi tiet BTL-BHH','A1:H5','detail'],['Chi tiet BTL-BHH','I1:Q5','money'],['Phieu BTL','A1:F5','docs'],['Thong tin','A1:B8','info'],...(missing.length?[['Chua co BHH','A1:H2','missing']]:[])]){
 const png=await w.render({sheetName:name,range,scale:1,format:'png'});await fs.writeFile(`${work}/${suffix}.png`,new Uint8Array(await png.arrayBuffer()));
}
await(await SpreadsheetFile.exportXlsx(w)).save(file);
console.log('SAVED',file);
