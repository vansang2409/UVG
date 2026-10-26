import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const root='D:/UVG',work=`${root}/scripts/data/amis-2024-linked-20261026`,out=process.env.UVG_OUTPUT_DIR||`${root}/outputs/amis-2024-theo-btl-bhh-20261026`;
const p=JSON.parse(await fs.readFile(`${work}/prepared.json`,'utf8'));
await fs.mkdir(out,{recursive:true});
const target=`${out}/KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx`;
if(!process.argv.includes('--replace'))try{await fs.access(target);throw Error('Không ghi đè file hiện có');}catch(e){if(e.code!=='ENOENT')throw e;}
const w=Workbook.create();let tid=0;
function col(n){let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;}
function sheet(name,headers,rows,widths,{dates=[],money=[],decimals=[],texts=[],wrap=[],inputs=[]}={}){
 const s=w.worksheets.add(name),end=rows.length+1,last=col(headers.length);s.showGridLines=false;
 s.getRange(`A1:${last}1`).values=[headers];
 if(rows.length)s.getRangeByIndexes(1,0,rows.length,headers.length).values=rows.map(r=>r.map((v,i)=>dates.includes(i)&&v?new Date(v):v));
 s.getRange(`A1:${last}${end}`).format={font:{name:'Arial',size:10},rowHeight:42,verticalAlignment:'center'};
 s.getRange(`A1:${last}1`).format={fill:'#24435C',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:48,wrapText:true};
 widths.forEach((x,i)=>s.getRangeByIndexes(0,i,1,1).format.columnWidth=x);
 if(rows.length){
  s.tables.add(`A1:${last}${end}`,true,`UVG_Linked_${++tid}`).style='TableStyleMedium2';
  for(const [indices,format] of [[dates,'dd/mm/yyyy'],[money,'#,##0;[Red](#,##0);0'],[decimals,'#,##0.00########'],[texts,'@']])for(const c of indices)s.getRange(`${col(c+1)}2:${col(c+1)}${end}`).setNumberFormat(format);
  for(const c of wrap)s.getRange(`${col(c+1)}2:${col(c+1)}${end}`).format.wrapText=true;
  for(const c of inputs)s.getRange(`${col(c+1)}2:${col(c+1)}${end}`).format.fill='#FFF2CC';
 }
 s.freezePanes.freezeRows(1);s.freezePanes.freezeColumns(1);return s;
}
const summary=sheet('Tong hop',['Phân nhóm nguồn','Số chứng từ','Số dòng nguồn','Tiền hàng gốc (VND)','Trả lại gốc (VND)','VAT gốc (VND)','Thanh toán gốc (VND)'],Object.entries(p.summary).map(([k,v])=>[k,v.documents,v.source_lines,v.goods,v.returns,v.vat,v.payment]),[36,18,20,28,28,28,29],{money:[3,4,5,6]});
summary.getRange('A7').values=[['Tổng sổ AMIS']];summary.getRange('B7:G7').formulas=[['=SUM(B2:B6)','=SUM(C2:C6)','=SUM(D2:D6)','=SUM(E2:E6)','=SUM(F2:F6)','=SUM(G2:G6)']];
summary.getRange('A7:G7').format={font:{bold:true},fill:'#E6EEF4',rowHeight:28};summary.getRange('D7:G7').setNumberFormat('#,##0');
const counts=p.summary['Đưa vào import'];
const info=[
 ['Hóa đơn trong file import',counts.documents],['Dòng hàng import',counts.import_lines],['Tiền hàng sau bù trừ',counts.net_goods],['VAT sau bù trừ',counts.net_vat],['Thanh toán sau bù trừ',counts.net_payment],
 ['Cách bù trừ','Chỉ dùng BHH gốc trong file liên kết AMIS. Trả hết: khớp SL, tiền hàng, VAT theo SKU/ĐVT. Trả một phần: trừ chính xác tiền nguồn, giữ đơn giá nguồn nếu chênh làm tròn tối đa 1đ.'],
 ['Hai đơn đã bù trừ một phần','BHHT2024006103 / BTL00070 và BHHT2024024887 / BTL00180. Chi tiết xem Bu tru tra lai.'],
 ['Hai đơn cần kiểm tra tiền trả','BHHS2024000629 / BTL00023 và BHHS2024005761 / BTL00024. Chưa bù trừ và chưa đưa vào import.'],
 ['Phiếu chưa có BHH','BTL00145 đang trống BHH trên AMIS. Giữ riêng, không suy đoán hoặc giữ lại đơn bán khác.'],
 ['Nguồn ngày và thanh toán','Ngày chứng từ AMIS; TM/CK cho toàn bộ theo xác nhận.'],
 ['Thuế suất','Suy ra duy nhất 5/8/10% từ tiền hàng và VAT, dung sai làm tròn 1đ/dòng. File import hiện 8%; giữ nguyên tiền VAT nguồn sau bù trừ.'],
 ['Loại trừ chưa rõ','7 mục ở Loai chua ro: gồm mã cũ đã gộp, mã năm 2023 và các hóa đơn thiếu BHH. Chưa tự loại thêm theo tên/MST.'],
 ['Trạng thái','Thay thế bộ import tại outputs/amis-2024-moi-20261026. Chưa nhập, ký hoặc phát hành hóa đơn.'],
 ['Nguồn và dấu kiểm','Các file nguồn giữ nguyên; SHA256 tại cuối sheet. Số dòng AMIS giúp đối chiếu từng dòng.'],
 ...p.confirmations.map(c=>[c.doc,c.basis]),...p.manifest.map(m=>[m.file,m.sha256])
];
summary.getRangeByIndexes(9,0,info.length,2).values=info;summary.getRange(`A10:A${9+info.length}`).format={columnWidth:36,wrapText:true};
for(let i=0;i<info.length;i++){summary.mergeCells(`B${10+i}:G${10+i}`);summary.getRange(`B${10+i}`).values=[[info[i][1]]];summary.getRange(`A${10+i}:G${10+i}`).format={rowHeight:44,wrapText:true};}
summary.getRange('B12:B14').setNumberFormat('#,##0');
const common=['Chứng từ','Ngày chứng từ','Mã khách hàng','Khách hàng','MST','Tiền hàng gốc','VAT gốc','Thanh toán gốc','BTL liên quan','Dòng AMIS','Lý do','Quyết định của bạn','BHH đúng/bổ sung'];
const pending=p.docs.filter(d=>d.category==='Cần kiểm tra');
sheet('Can kiem tra',common,pending.map(d=>[d.doc,d.date,d.customer_code,d.buyer,d.tax,d.source_sums.goods,d.source_sums.vat,d.source_sums.payment,d.btls.join(', '),d.rows.join(', '),d.reasons.join('; '),null,null]),[27,18,22,38,19,25,24,27,26,30,110,55,35],{dates:[1],money:[5,6,7],texts:[0,2,4,12],wrap:[3,8,9,10],inputs:[11,12]});
const raw=pending.flatMap(d=>d.original_lines.map(l=>[l.row,...l.values]));
sheet('Dong nguon can kiem tra',['Dòng AMIS',...p.source_headers],raw,[16,...p.source_headers.map((_,i)=>[6,7,9,10,13,14].includes(i)?70:23)],{dates:[1,2,5],money:[21,25,26,28,29,31,33,34,35,40],decimals:[18,19,20,27,37],texts:[3,4,6,9,12,13],wrap:[7,8,10,11,14,15]});
const omitted=p.docs.filter(d=>['Đã loại theo danh sách','Đã trả hết'].includes(d.category));
sheet('Da loai va tra het',['Chứng từ','Phân nhóm','Ngày chứng từ','Khách hàng','Tiền hàng gốc','VAT gốc','Thanh toán gốc','BTL liên quan','Dòng AMIS','Căn cứ'],omitted.map(d=>[d.doc,d.category,d.date,d.buyer,d.source_sums.goods,d.source_sums.vat,d.source_sums.payment,d.btls.join(', '),d.rows.join(', '),d.reasons.join('; ')]),[27,32,18,38,25,24,27,28,35,100],{dates:[2],money:[4,5,6],texts:[0],wrap:[1,3,7,8,9]});
const audit=sheet('Bu tru tra lai',['BHH gốc','BTL','Dòng AMIS','Mã hàng','Tên hàng','ĐVT','SL bán gốc','Tiền hàng bán gốc','VAT bán gốc','SL đã bù trừ','Tiền hàng đã trừ','VAT đã trừ','SL còn lại','Tiền hàng còn lại','VAT còn lại','Phân nhóm BHH','Xử lý bù trừ','Ghi chú'],p.net_audit,[27,27,17,25,80,12,19,26,24,20,26,24,20,26,24,32,30,105],{decimals:[6,9,12],money:[7,8,10,11,13,14],texts:[0,3],wrap:[1,4,15,16,17]});
const ae=p.net_audit.length+1;
for(const [c,a,b] of [['M','G','J'],['N','H','K'],['O','I','L']])audit.getRange(`${c}2:${c}${ae}`).formulas=p.net_audit.map((_,i)=>[`=${a}${i+2}-${b}${i+2}`]);
const links=p.links.map(r=>[...r,p.return_status[r[2]]||'Chưa có BHH',p.link_issues[r[2]]||(!r[2]?'Trống BHH thực tế trên AMIS':null)]);
const ls=sheet('Lien ket BTL-BHH',['BTL','Ngày hạch toán','BHH gốc','Dòng BTL','Mã hàng','Tên hàng','ĐVT','SL trả','Đơn giá','Tiền hàng trả','Thuế suất','VAT trả','Tổng trả','Khách hàng','Liên kết AMIS','Khớp sổ nguồn','Dòng AMIS','Xử lý','Nội dung cần kiểm tra'],links,[22,18,27,14,26,80,12,19,25,25,18,25,25,37,30,30,20,30,110],{dates:[1],money:[9,11,12],decimals:[7,8],texts:[0,2,4,16],wrap:[5,13,14,15,17,18]});
ls.getRange(`M2:M${links.length+1}`).formulas=links.map((_,i)=>[`=J${i+2}+L${i+2}`]);
sheet('Loai chua ro',['Hóa đơn','Vị trí nguồn','BHH','Tên trong nguồn','Nội dung cần kiểm tra','BHH đúng do bạn bổ sung'],p.unresolved_exclusions,[29,49,28,65,100,38],{texts:[0,2,5],wrap:[1,3,4],inputs:[5]});
w.recalculate();
console.log((await w.inspect({kind:'region',sheetId:'Tong hop',range:'A1:G7',maxChars:2500,tableMaxRows:7,tableMaxCols:7})).ndjson);
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!',options:{useRegex:true,maxResults:10},maxChars:1000})).ndjson);
for(const [name,range,key] of [['Tong hop','A1:G7','summary'],['Can kiem tra','A1:F4','pending'],['Can kiem tra','I1:M4','reasons'],['Bu tru tra lai','G1:O4','net']]){
 const png=await w.render({sheetName:name,range,scale:1,format:'png'});await fs.writeFile(`${work}/${key}.png`,new Uint8Array(await png.arrayBuffer()));
}
await(await SpreadsheetFile.exportXlsx(w)).save(target);console.log('SAVED',target);
