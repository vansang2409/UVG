import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile, FileBlob } from '@oai/artifact-tool';

const root='D:/UVG';
const work=`${root}/scripts/data/new-amis-2024-20261026`;
const out=`${root}/outputs/amis-2024-moi-20261026`;
await fs.mkdir(out,{recursive:true});
const p=JSON.parse(await fs.readFile(`${work}/prepared.json`,'utf8'));
const mode=process.argv[2]||'import';
const template=`${root}/data/Du lieu Meinvoice/Mau_hoa_don_GTGT_MTT_1_thue_suat.xls`;
const wb=mode==='import'?await SpreadsheetFile.importXlsx(await FileBlob.load(template)):Workbook.create();
const summaryMeta=[];
const clean=v=>typeof v==='string'&&/^[=+@]/.test(v)?`'${v}`:v;
const date=v=>v?new Date(v):null;
const column=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};

function table(name,headers,rows,widths,opts={}){
  const start=opts.headerRow||1;
  const s=opts.template?wb.worksheets.getItem(name):wb.worksheets.add(name),last=column(headers.length-1),end=rows.length+start;
  s.showGridLines=false;
  s.getRange(`A${start}:${last}${start}`).values=[headers];
  for(let offset=0;offset<rows.length;offset+=1500){
    const chunk=rows.slice(offset,offset+1500).map(r=>r.map(clean));
    s.getRangeByIndexes(offset+start,0,chunk.length,headers.length).values=chunk;
  }
  s.getRange(`A1:${last}${end}`).format.font={name:'Arial',size:10};
  s.getRange(`A${start}:${last}${start}`).format={fill:'#24435C',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:44,verticalAlignment:'center'};
  widths.forEach((width,i)=>{s.getRangeByIndexes(0,i,1,1).format.columnWidth=width;});
  if(rows.length){
    const t=s.tables.add(`A${start}:${last}${end}`,true,`UVG_Table_${summaryMeta.length+1}`);t.style='TableStyleMedium2';t.showFilterButton=true;
    s.getRange(`A${start+1}:${last}${end}`).format.rowHeight=24;
    for(const c of opts.money||[])s.getRange(`${column(c)}${start+1}:${column(c)}${end}`).setNumberFormat('#,##0;[Red](#,##0);0');
    for(const c of opts.dates||[])s.getRange(`${column(c)}${start+1}:${column(c)}${end}`).setNumberFormat('dd/mm/yyyy');
    for(const c of opts.text||[])s.getRange(`${column(c)}${start+1}:${column(c)}${end}`).setNumberFormat('@');
  }
  s.freezePanes.freezeRows(start);s.freezePanes.freezeColumns(opts.freezeCols||2);
  summaryMeta.push({name,headers,rows:rows.length,headerRow:start});
  return s;
}

const candidates=p.docs.filter(d=>d.category==='Đưa vào import').slice(0,process.argv.includes('--preview-only')?3:undefined);
if(mode==='import'){
  const rows=[];
  candidates.forEach((d,index)=>d.lines.forEach((line,j)=>{
    const r=line.values,first=j===0;
    rows.push([index+1,first?date(d.date):null,first?d.buyer:null,first?d.tax:null,first?d.address:null,first?d.buyer:null,null,null,null,first?'TM/CK':null,first?d.rate:null,first?d.sums.vat:null,r[7],r[16],r[17],r[19],r[20],d.doc,first?d.sums.goods:null,first?d.sums.payment:null,line.row,r[12]]);
  }));
  wb.worksheets.getItem('Hóa đơn GTGT').getRange('A10:Q121').clear({applyTo:'contents'});
  wb.worksheets.getItem('Hóa đơn GTGT').getRange('A7').values=[['- Dữ liệu sổ AMIS 2024 sau loại trừ; ngày lập theo Ngày chứng từ, thanh toán TM/CK theo xác nhận.']];
  table('Hóa đơn GTGT',[
    'Số thứ tự hóa đơn (*)','Ngày hóa đơn','Tên đơn vị mua hàng','Mã số thuế','Địa chỉ','Người mua hàng','Email','Số điện thoại','Căn cước công dân','Hình thức thanh toán (*)','Thuế suất GTGT (%)','Tiền thuế GTGT','Tên hàng hóa/dịch vụ (*)','ĐVT','Số lượng','Đơn giá','Thành tiền','Số chứng từ AMIS','Tổng tiền hàng','Tổng tiền thanh toán','Dòng nguồn AMIS','Mã hàng hóa/dịch vụ'
  ],rows,[14,14,34,18,50,34,25,20,22,20,17,22,86,12,14,19,22,26,24,25,17,25],{dates:[1],money:[11,14,15,16,18,19],text:[3,17,21],template:true,headerRow:9});
  const guide=[
    ['Phạm vi','Hóa đơn mới từ sổ AMIS 2024 sau loại trừ và tách trường hợp cần kiểm tra.'],
    ['Sheet nhập khẩu','Hóa đơn GTGT; chọn dòng 9 làm tiêu đề. Mẫu GTGT máy tính tiền, một thuế suất do người dùng cung cấp.'],
    ['Gom hóa đơn','Mỗi số chứng từ AMIS là một hóa đơn; các dòng cùng số thứ tự thuộc cùng hóa đơn.'],
    ['Ngày lập','Lấy cột Ngày chứng từ AMIS theo yêu cầu người dùng; không dùng Ngày hóa đơn nguồn AMIS.'],
    ['Hình thức thanh toán','TM/CK theo xác nhận của người dùng.'],
    ['Thuế suất','8% suy ra duy nhất từ tiền hàng/VAT của từng dòng (sai số làm tròn tối đa 1đ). Giữ nguyên mọi số tiền nguồn.'],
    ['Tổng hóa đơn','Tổng tiền hàng, thuế suất, tổng VAT và thanh toán chỉ ghi tại dòng đầu của mỗi hóa đơn.'],
    ['Thông tin người mua','Giữ nguyên tên, địa chỉ và MST trong AMIS, kể cả tên khách hàng dùng chung theo sàn. Email trống vì nguồn không có.'],
    ['Cột đối chiếu','Giữ nguyên 17 cột A:Q của mẫu. Bổ sung R:V để đối chiếu BHH, tổng tiền, dòng nguồn và mã hàng.'],
    ['Mẫu nhập','Mau_hoa_don_GTGT_MTT_1_thue_suat.xls do người dùng tải từ meInvoice; đã bỏ toàn bộ dữ liệu ví dụ.'],
    ['Giới hạn phạm vi','Chỉ dùng danh sách loại HD cty 2024 và ghi chú BHH tách/gộp. Chưa kiểm chứng trạng thái mọi hóa đơn cũ trên meInvoice.'],
    ['Trả lại','Phiếu trả lại và mọi đơn bán có thể liên quan được giữ ngoài import, nằm trong file kiểm tra.'],
    ['Nguồn AMIS','So_chi_tiet_ban_hang_2024.xlsx, sheet SỔ CHI TIẾT BÁN HÀNG; cột Dòng nguồn AMIS giữ số dòng gốc.'],
    ['Nguồn loại trừ','HD Xuất CTY 2024.xlsx, tab HD cty 2024, cột BHH và Hướng xử lý.'],
    ['Hướng dẫn MISA','https://helpv4.meinvoice.vn/kb/nhap-khau-hoa-don-mot-thue-suat/'],
    ['Thực hiện','Chưa nhập, ký hay phát hành trên phần mềm.']
  ];
  table('Huong dan',['Mục','Nội dung'],guide,[30,125],{freezeCols:1});
  wb.worksheets.getItem('Huong dan').getRange('B2:B17').format.wrapText=true;
  wb.worksheets.getItem('Huong dan').getRange('A2:B17').format.rowHeight=38;
}else{
  const statHeaders=['Phân nhóm','Số chứng từ','Số dòng AMIS','Doanh số bán (VND)','Trả lại (VND)','VAT (VND)','Thanh toán (VND)'];
  const statRows=Object.entries(p.summary).map(([k,v])=>[k,v.documents,v.lines,v.goods,v.returns,v.vat,v.payment]);
  const sumSheet=table('Tong hop',statHeaders,[...statRows,['Tổng sổ AMIS',null,null,null,null,null,null]], [29,18,20,27,24,24,27],{money:[3,4,5,6]});
  sumSheet.getRange('B5:G5').formulas=[['=SUM(B2:B4)','=SUM(C2:C4)','=SUM(D2:D4)','=SUM(E2:E4)','=SUM(F2:F4)','=SUM(G2:G4)']];
  sumSheet.getRange('A5:G5').format.font={name:'Arial',size:10,bold:true};
  sumSheet.getRange('A8:B15').values=[
    ['Cách sử dụng','Lọc sheet Cần kiểm tra theo Lý do. Nhập quyết định và BHH đúng ở hai cột cuối.'],
    ['Trả lại','15.879 đơn bán chỉ là ứng viên của phiếu trả lại, chưa kết luận đã trả. Chưa đưa vào import.'],
    ['Đối chiếu trả lại','Cùng mã KH, mã hàng, đơn giá ±1đ, bán trước hoặc cùng ngày trả, đủ SL. Không chọn tự động.'],
    ['Ngày lập','Theo Ngày chứng từ AMIS. Hình thức thanh toán TM/CK theo xác nhận.'],
    ['Thuế suất','AMIS không có cột thuế suất. 8% ở file import được suy ra từ tiền hàng/VAT, không đổi tiền.'],
    ['Thiếu mã loại','Xem sheet Loai chua ro: HĐ thiếu BHH, mã cũ không còn và 3 HĐ chỉ có ở bảng ghi nhận.'],
    ['Nguồn','So_chi_tiet_ban_hang_2024.xlsx và HD Xuất CTY 2024.xlsx, lưu nguyên bản trong data.'],
    ['Phạm vi','Tab F0/F1/F2 trong file loại là lịch sử tham khảo, không loại mọi hóa đơn trong tab đó.']
  ];
  sumSheet.getRange('B8:G15').merge(true);sumSheet.getRange('B8:B15').format.wrapText=true;sumSheet.getRange('A8:G15').format.rowHeight=44;
  const commonHeaders=['Số chứng từ AMIS','Ngày chứng từ','Mã khách hàng','Tên khách hàng','MST','Doanh số bán (VND)','Trả lại (VND)','VAT (VND)','Thanh toán (VND)','Số dòng','Dòng AMIS','Lý do','Quyết định của bạn','BHH đúng/liên quan'];
  const commonRows=ds=>ds.map(d=>[d.doc,date(d.date),d.customer_code,d.buyer,d.tax,d.sums.goods,d.sums.returns,d.sums.vat,d.sums.payment,d.lines.length,d.rows.join(', '),d.reasons.join('; '),null,null]);
  const check=p.docs.filter(d=>d.category==='Cần kiểm tra');
  table('Can kiem tra',commonHeaders,commonRows(check),[26,17,20,38,18,24,23,23,25,12,30,120,55,35],{dates:[1],money:[5,6,7,8],text:[0,2,4,13]});
  const pending=wb.worksheets.getItem('Can kiem tra');pending.getRange(`M2:N${check.length+1}`).format.fill='#FFF2CC';
  const rawRows=[];
  for(const d of check)for(const l of d.lines){
    const r=l.values.map((v,i)=>[0,1,4].includes(i)?date(v):v);
    rawRows.push([l.row,...r]);
  }
  table('Chi tiet can kiem tra',['Dòng AMIS',...p.headers],rawRows,[14,...p.headers.map((h,i)=>[6,7,9,10,13,14].includes(i)?68:([0,1,4].includes(i)?17:22))],{dates:[1,2,5],money:[18,19,20,21,25,26,27,28,29,31,33,34,35,37,40],text:[3,4,6,9,12,13]});
  table('Da loai',commonHeaders,commonRows(p.docs.filter(d=>d.category==='Đã loại')),[26,17,20,38,18,24,23,23,25,12,30,120,55,35],{dates:[1],money:[5,6,7,8],text:[0,2,4,13]});
  const unresolved=[];
  for(const e of p.exclusions){
    if(!e.codes.length)unresolved.push([e.key,`HD cty 2024 dòng ${e.row}`,null,e.values[4],'Thiếu BHH: chưa tự loại đơn theo tên/MST',null]);
    for(const code of e.codes)if(p.exclusion_codes_not_found.includes(code))unresolved.push([e.key,`HD cty 2024 dòng ${e.row}`,code,e.values[4],code==='BHHN2024000079'?'Mã cũ đã gộp vào BHHN2024000078; đã loại mã mới':'Mã thuộc năm 2023, không có trong sổ AMIS 2024',null]);
  }
  const extraKeys=new Set();
  for(const [row,key,buyer,tax,order] of p.extra_invoices)if(!extraKeys.has(key)){
    extraKeys.add(key);unresolved.push([key,`HĐ ghi nhận UVG 2024 dòng ${row}`,null,buyer,`Không có trong tab HD cty 2024; chưa có BHH. Mã đơn ghi nhận: ${order||''}`,null]);
  }
  table('Loai chua ro',['Hóa đơn','Vị trí nguồn','BHH','Tên trong nguồn','Nội dung cần kiểm tra','BHH đúng do bạn bổ sung'],unresolved,[29,49,28,65,110,35]);
  const backRows=[];
  candidates.forEach((d,i)=>backRows.push([i+1,d.doc,date(d.date),d.buyer,d.lines.length,d.rows.join(', '),d.sums.goods,d.sums.vat,d.sums.payment,d.rate,date(d.accounting_date),date(d.invoice_date_source)]));
  table('Doi chieu import',['STT hóa đơn','Số chứng từ AMIS','Ngày chứng từ','Tên khách hàng','Số dòng','Dòng nguồn AMIS','Tiền hàng (VND)','VAT (VND)','Thanh toán (VND)','Thuế suất suy ra (%)','Ngày hạch toán AMIS','Ngày hóa đơn nguồn AMIS'],backRows,[16,27,18,45,13,35,25,24,26,22,20,24],{dates:[2,10,11],money:[6,7,8],text:[1]});
  const linkRows=[];
  for(const [doc,row,candidates] of p.return_evidence)linkRows.push([doc,row,candidates.length,`Lọc cột Lý do của sheet Can kiem tra theo ${doc}; các đơn trong Da loai vẫn được giữ ngoài import.`]);
  table('Phieu tra lai',['Phiếu trả lại','Dòng AMIS','Số dòng bán có thể liên quan','Cách tìm BHH bán cần kiểm tra'],linkRows,[25,17,25,130]);
}

console.log('RECALCULATE',mode,summaryMeta.map(x=>[x.name,x.rows]));
wb.recalculate();
console.log((await wb.inspect({kind:'region',sheetId:summaryMeta[0].name,range:mode==='import'?'A9:F12':'A1:G5',maxChars:1600,tableMaxRows:5,tableMaxCols:8})).ndjson);
// Xem các vùng vừa tạo trước khi xuất.
for(const info of summaryMeta){
  const endCol=column(Math.min(info.headers.length-1,info.name==='Hóa đơn GTGT'?7:5));
  const range=`A${info.headerRow}:${endCol}${info.headerRow+Math.min(info.rows,info.name==='Tong hop'?4:3)}`;
  const blob=await wb.render({sheetName:info.name,range,scale:1,format:'png'});
  await fs.writeFile(`${work}/${mode}-${summaryMeta.indexOf(info)}.png`,new Uint8Array(await blob.arrayBuffer()));
}
if(mode==='import'){
  const blob=await wb.render({sheetName:'Hóa đơn GTGT',range:'K9:Q12',scale:1,format:'png'});
  await fs.writeFile(`${work}/import-money.png`,new Uint8Array(await blob.arrayBuffer()));
}
console.log('EXPORT',mode);
if(process.argv.includes('--preview-only'))process.exit(0);
const file=path.join(out,mode==='import'?'IMPORT_HOA_DON_MOI_AMIS_2024.xlsx':'KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx');
try{await fs.access(file);throw new Error(`Không ghi đè file đã có: ${file}`);}catch(e){if(e.code!=='ENOENT')throw e;}
await(await SpreadsheetFile.exportXlsx(wb)).save(file);
await fs.writeFile(`${work}/${mode}-sheets.json`,JSON.stringify(summaryMeta,null,2));
console.log('SAVED',file);
