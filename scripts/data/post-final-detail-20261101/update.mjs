import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const root='D:/UVG',work=`${root}/scripts/data/post-final-detail-20261101`;
const file=`${root}/outputs/kiem-tra-sau-dieu-chinh-thay-the-20261101/KET_QUA_KIEM_TRA_SAU_PHAT_HANH_20261101.xlsx`;
const p=JSON.parse(await fs.readFile(`${work}/summary.json`,'utf8'));
const a=JSON.parse(await fs.readFile(`${work}/adjustments.json`,'utf8'));
const r=JSON.parse(await fs.readFile(`${work}/replacements.json`,'utf8'));
const rounding=JSON.parse(await fs.readFile(`${work}/price-rounding.json`,'utf8'));
if(p.issues.length)throw new Error(JSON.stringify(p.issues));
if(a.length!==163||r.length!==2000||rounding.length!==7)throw new Error('Unexpected source coverage');
const before=`${work}/before.xlsx`;
try {await fs.access(before);}catch{await fs.copyFile(file,before);}
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(before));
const s=w.worksheets.getItem('Tong hop'),ad=w.worksheets.getItem('Dieu chinh 163'),re=w.worksheets.getItem('Thay the 2000');
const beforeInspect=await w.inspect({kind:'workbook,sheet,table',maxChars:2500,tableMaxRows:2,tableMaxCols:3});
await fs.writeFile(`${work}/before-inspection.ndjson`,beforeInspect.ndjson);
for(let i=0;i<a.length;i++){
 const x=a[i],n=i+5;
 if(ad.getRange(`E${n}`).values[0][0]!==x.new)throw new Error(`Adjustment row mismatch ${n}`);
 ad.getRange(`P${n}:Q${n}`).values=[[x.detail_cqt_statuses.join(', '),x.detail_parent_statuses.join(', ')]];
 ad.getRange(`S${n}`).values=[[`Khớp lý do; ${x.detail_lines} dòng hàng${x.price_rounding_only?'; đơn giá làm tròn':''}`]];
 ad.getRange(`U${n}`).values=[[`${p.source_manifest[0].file.split('/').at(-1)} / dòng ${x.detail_source_rows.join(', ')}`]];
}
ad.getRange('A2').values=[['Đủ 163 HĐ hợp lệ / 219 dòng hàng. Có 7 dòng thuộc 6 HĐ làm tròn đơn giá; thành tiền không đổi.']];
ad.getRange('S4').values=[['Đối chiếu chi tiết']];
ad.getRange('S1').format.columnWidth=48;
ad.getRange('S5:S167').format.wrapText=true;
ad.getRange('U1').format.columnWidth=93;
ad.getRange('U5:U167').format.wrapText=true;
ad.getRange('S5:S167').conditionalFormats.add('containsText',{text:'đơn giá làm tròn',format:{fill:'#FFF2CC'}});
ad.getRange('A171').values=[['Đơn giá làm tròn: 7 dòng thuộc 6 hóa đơn']];
ad.getRange('A171').format={font:{name:'Arial',size:12,bold:true,color:'#24435C'},rowHeight:28};
const rh=['STT','HĐ điều chỉnh','HĐ gốc','Tên hàng','Số lượng','Đơn giá import','Đơn giá bảng kê','Thành tiền giữ nguyên','Dòng bảng kê chi tiết'];
ad.getRange('A173:I173').values=[rh];
ad.getRange('A174:I180').values=rounding.map((x,i)=>[i+1,x.invoice,x.old,x.name,x.qty,Number(x.planned_price),x.price,x.goods,x.source_row]);
const t=ad.tables.add('A173:I180',true,'RoundedReturnPrices');t.style='TableStyleMedium2';t.showFilterButton=true;
ad.getRange('A173:I173').format={fill:'#24435C',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:42,horizontalAlignment:'center',verticalAlignment:'center'};
ad.getRange('A174:I180').format={font:{name:'Arial',size:10},verticalAlignment:'center',wrapText:true,rowHeight:74};
ad.getRange('B174:C180').setNumberFormat('@');
ad.getRange('E174:E180').setNumberFormat('0');
ad.getRange('F174:G180').setNumberFormat('#,##0.00');
ad.getRange('H174:I180').setNumberFormat('#,##0;[Red](#,##0);0');
ad.getRange('A182').values=[['Bảng kê không xuất cột mã hàng hóa/dịch vụ, chưa xác nhận được SKU đã phát hành.']];
ad.getRange('A182').format={font:{name:'Arial',size:10,color:'#525C66'},rowHeight:24};

for(let i=0;i<r.length;i++){
 const x=r[i],n=i+5;
 if(re.getRange(`D${n}`).values[0][0]!==x.new)throw new Error(`Replacement row mismatch ${n}`);
 re.getRange(`L${n}`).values=[[x.detail_cqt_statuses.join(', ')]];
 re.getRange(`N${n}:O${n}`).values=[['Đúng 1 dòng ghi chú, đúng nội dung','Nguồn trống; tiền hàng + VAT = 0']];
 re.getRange(`Q${n}`).values=[[`${x.detail_fields.source_file.split('/').at(-1)} / dòng ${x.detail_fields.source_row}`]];
}
re.getRange('A2').values=[['2.000 HĐ hợp lệ; mỗi HĐ đúng 1 dòng ghi chú. Cột thanh toán nguồn trống, không ghi đè thành số 0.']];
re.getRange('N4:O4').values=[['Đối chiếu dòng ghi chú','Thanh toán nguồn / cách tính']];
re.getRange('N1:O1').format.columnWidth=45;
re.getRange('N5:O2004').format.wrapText=true;
re.getRange('A5:Q2004').format.rowHeight=34;
re.getRange('Q1').format.columnWidth=82;

s.getRange('A2').values=[['Đã đối chiếu thêm bảng kê chi tiết xuất ngày 01/11/2026, kỳ 03/10–01/11/2026.']];
s.getRange('A7:B7').format={fill:'#FFFFFF',font:{name:'Arial',size:10,bold:false}};
s.getRange('E16:E18').values=[
 ['Đủ 163 HĐ hợp lệ; khớp 219 dòng hàng và toàn bộ số tiền'],
 ['Dòng ghi chú; thanh toán tính từ tiền hàng + VAT, nguồn vẫn trống'],
 ['Thanh toán tính toán; ô thanh toán nguồn trống được giữ nguyên'],
];
s.getRange('D17').formulas=[["=SUM('Thay the 2000'!I5:J2004)"]];
s.getRange('D18').formulas=[['=SUM(D15:D17)']];
s.getRange('D17:D18').format.fill='#FFFFFF';
s.getRange('A19:E19').clear({applyTo:'contents'});
s.getRange('A22:A30').values=[
 ['163/163 HĐ điều chỉnh đều HĐ hợp lệ; không còn 63 HĐ chờ như bảng tổng lần trước.'],
 ['2.000 HĐ thay thế đúng dòng ghi chú; số lượng, đơn giá, thành tiền, VAT, thanh toán và thuế suất dòng đều trống.'],
 ['Thanh toán tính từ tiền hàng + VAT. Giữ nguyên các ô thanh toán nguồn trống trong tab thay thế.'],
 ['Nguồn: Bang_ke_hoa_don_da_su_dung_1793524594152.zip (hai phần bảng kê tổng).'],
 ['Nguồn: Danh_sach_hoa_don_dieu_chinh.xlsx; Danh_sach_hoa_don_thay_the.xlsx.'],
 ['Nguồn mới: Bang_ke_chi_tiet_hoa_don_da_su_dung_20261101_161720.zip.'],
 ['Bảng kê không xuất SKU. Dòng ghi chú không xuất thuế suất, chưa kiểm được yêu cầu 8% trên dòng thay thế.'],
 ['6 HĐ / 7 dòng làm tròn đơn giá đến đồng; tên hàng, số lượng, thành tiền, VAT và thanh toán không đổi.'],
 ['MISA: dòng ghi chú không cho nhập đơn giá, thành tiền, thuế suất, tiền thuế.'],
];
s.getRange('A31').values=[['https://helpv4.meinvoice.vn/kb/lamthenaodephathanhhoadoncodongghichudiengiai/']];
s.getRange('A22:A31').format={font:{name:'Arial',size:10,color:'#525C66'},rowHeight:24};
s.getRange('A22').format.font={name:'Arial',size:10,bold:true,color:'#24435C'};
w.recalculate();
const inspected=await w.inspect({kind:'region',sheetId:'Tong hop',range:'A4:E18',tableMaxRows:15,tableMaxCols:5,maxChars:4000});
const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:20},maxChars:1000});
await fs.writeFile(`${work}/inspection.ndjson`,`${inspected.ndjson}\n${errors.ndjson}`);
await(await SpreadsheetFile.exportXlsx(w)).save(file);
for(const [sheetName,range,name]of[['Tong hop','A1:E18','summary-preview'],['Tong hop','A21:E31','notes-preview'],['Dieu chinh 163','P102:S108','status-preview'],['Dieu chinh 163','A171:I180','rounding-preview'],['Thay the 2000','I4:O7','replacement-preview']]){
 const png=await w.render({sheetName,range,scale:1,format:'png'});
 await fs.writeFile(`${work}/${name}.png`,new Uint8Array(await png.arrayBuffer()));
}
console.log(inspected.ndjson);console.log(errors.ndjson);console.log(JSON.stringify({output:file,statuses:'163 + 2000 HĐ hợp lệ',rounded_invoices:6,rounded_lines:7}));
