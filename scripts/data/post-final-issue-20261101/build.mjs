import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const root='D:/UVG',work=`${root}/scripts/data/post-final-issue-20261101`;
const output=`${root}/outputs/kiem-tra-sau-dieu-chinh-thay-the-20261101`;
const p=JSON.parse(await fs.readFile(`${work}/summary.json`,'utf8'));
const adjustments=JSON.parse(await fs.readFile(`${work}/adjustments.json`,'utf8'));
const replacements=JSON.parse(await fs.readFile(`${work}/replacements.json`,'utf8'));
if(p.issues.length)throw new Error(JSON.stringify(p.issues));
const w=Workbook.create();
const overview=w.worksheets.add('Tong hop');
const ad=w.worksheets.add('Dieu chinh 163');
const re=w.worksheets.add('Thay the 2000');
const headerFormat={fill:'#24435C',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:42,horizontalAlignment:'center',verticalAlignment:'center'};
function setup(s,title,headers,rows,widths,name){
 const last=rows.length+4;
 s.showGridLines=false;
 s.getRangeByIndexes(0,0,last,headers.length).format={font:{name:'Arial',size:10},verticalAlignment:'center',rowHeight:24};
 widths.forEach((v,i)=>s.getRangeByIndexes(0,i,1,1).format.columnWidth=v);
 s.getRange('A1').values=[[title]];
 s.getRange('A1').format.font={name:'Arial',size:14,bold:true,color:'#24435C'};
 s.getRange('A1').format.rowHeight=30;
 s.getRangeByIndexes(3,0,1,headers.length).values=[headers];
 for(let i=0;i<rows.length;i+=500)s.getRangeByIndexes(i+4,0,Math.min(500,rows.length-i),headers.length).values=rows.slice(i,i+500);
 const table=s.tables.add(s.getRangeByIndexes(3,0,rows.length+1,headers.length),true,name);
 table.style='TableStyleMedium2';table.showFilterButton=true;
 s.getRangeByIndexes(3,0,1,headers.length).format=headerFormat;
 s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(4);
 return last;
}
const adHeaders=['STT','Số chứng từ AMIS','Mã đơn HĐ gốc','HĐ gốc','HĐ điều chỉnh','Ngày HĐ điều chỉnh','Tiền hàng dự kiến','Tiền hàng thực tế','Chênh tiền hàng','VAT dự kiến','VAT thực tế','Chênh VAT','Thanh toán dự kiến','Thanh toán thực tế','Chênh thanh toán','Trạng thái gửi CQT','Trạng thái HĐ gốc','Lý do điều chỉnh','Khớp lý do','Mã CQT HĐ mới','Nguồn bảng kê'];
const adRows=adjustments.map((a,i)=>[i+1,a.doc,a.order,a.old,a.new,new Date(`${a.new_date}T00:00:00Z`),a.expected[0],a.actual[0],null,a.expected[1],a.actual[1],null,a.expected[2],a.actual[2],null,a.cqt_status,a.parent_status,a.reason,a.checks.reason?'Khớp':'Khác',a.cqt,`${a.source_file.split('/').at(-1)} / dòng ${a.source_row}`]);
const adLast=setup(ad,'Điều chỉnh trả hàng đã phát hành',adHeaders,adRows,[8,24,27,25,25,20,24,24,22,22,22,20,25,25,23,23,28,105,16,34,65],'IssuedReturns163');
ad.getRange('A2').values=[['Đủ 163 liên kết và số tiền. Theo dõi riêng 63 HĐ còn trạng thái Đã gửi CQT.']];
for(const c of ['B','C','D','E','T'])ad.getRange(`${c}5:${c}${adLast}`).setNumberFormat('@');
ad.getRange(`F5:F${adLast}`).setNumberFormat('dd/mm/yyyy');
ad.getRange(`G5:O${adLast}`).setNumberFormat('#,##0;[Red](#,##0);0');
for(let r=5;r<=adLast;r++){
 ad.getRange(`I${r}`).formulas=[[`=H${r}-G${r}`]];
 ad.getRange(`L${r}`).formulas=[[`=K${r}-J${r}`]];
 ad.getRange(`O${r}`).formulas=[[`=N${r}-M${r}`]];
}
ad.getRange(`R5:R${adLast}`).format.wrapText=true;
ad.getRange(`A5:U${adLast}`).format.rowHeight=44;
ad.getRange(`P5:P${adLast}`).conditionalFormats.add('containsText',{text:'Đã gửi CQT',format:{fill:'#FFF2CC'}});
const reHeaders=['STT','HĐ giữ lại','HĐ trùng bị thay thế','HĐ thay thế mới','Ngày HĐ mới','Tiền hàng HĐ cũ','VAT HĐ cũ','Thanh toán HĐ cũ','Tiền hàng HĐ mới','VAT HĐ mới','Thanh toán HĐ mới','Trạng thái gửi CQT mới','Trạng thái HĐ cũ','Tiền hàng/VAT về 0','Thanh toán nguồn','Mã CQT HĐ mới','Nguồn bảng kê'];
const reRows=replacements.map((r,i)=>[i+1,r.retained,r.old,r.new,new Date(`${r.new_date}T00:00:00Z`),...r.before,...r.after,r.cqt_status,r.parent_status,r.checks.new_goods_vat_zero?'Đã về 0':'Cần kiểm',r.after[2]===null?'Trống; chờ bảng chi tiết':'Có giá trị',r.cqt,`${r.source_file.split('/').at(-1)} / dòng ${r.source_row}`]);
const reLast=setup(re,'Thay thế lô hóa đơn trùng phát hành sau',reHeaders,reRows,[8,25,26,25,18,24,22,25,24,22,25,27,28,24,34,34,65],'IssuedReplacements2000');
re.getRange('A2').values=[['Đủ 2.000 liên kết, tiền hàng/VAT bản mới bằng 0. Tổng thanh toán nguồn đang trống.']];
for(const c of ['B','C','D','P'])re.getRange(`${c}5:${c}${reLast}`).setNumberFormat('@');
re.getRange(`E5:E${reLast}`).setNumberFormat('dd/mm/yyyy');
re.getRange(`F5:K${reLast}`).setNumberFormat('#,##0;[Red](#,##0);0');

overview.showGridLines=false;
overview.getRange('A1:E31').format={font:{name:'Arial',size:10},verticalAlignment:'center',rowHeight:24};
[60,25,25,30,70].forEach((v,i)=>overview.getRangeByIndexes(0,i,1,1).format.columnWidth=v);
overview.getRange('A1').values=[['Kiểm tra sau phát hành điều chỉnh và thay thế']];
overview.getRange('A1').format.font={name:'Arial',size:14,bold:true,color:'#24435C'};
overview.getRange('A1').format.rowHeight=32;
overview.getRange('A2').values=[['Nguồn xuất kỳ 31/10–01/11/2026. Kết quả theo dữ liệu trong ba file nhận được.']];
const counts=[
 ['HĐ điều chỉnh đã có số',`=COUNTA('Dieu chinh 163'!E5:E${adLast})`],
 ['HĐ điều chỉnh khớp tiền và tham chiếu',p.adjustments.invoices],
 ['Điều chỉnh: HĐ hợp lệ',`=COUNTIF('Dieu chinh 163'!P5:P${adLast},"HĐ hợp lệ")`],
 ['Điều chỉnh: Đã gửi CQT',`=COUNTIF('Dieu chinh 163'!P5:P${adLast},"Đã gửi CQT")`],
 ['HĐ thay thế đủ liên kết',`=COUNTA('Thay the 2000'!D5:D${reLast})`],
 ['Thay thế: HĐ hợp lệ',`=COUNTIF('Thay the 2000'!L5:L${reLast},"HĐ hợp lệ")`],
 ['HĐ trùng cũ đã bị thay thế',`=COUNTIF('Thay the 2000'!M5:M${reLast},"Hóa đơn đã bị thay thế")`],
 ['HĐ thay thế trống tổng thanh toán',`=COUNTBLANK('Thay the 2000'!K5:K${reLast})`],
];
counts.forEach(([label,v],i)=>{
 overview.getRange(`A${i+4}`).values=[[label]];
 if(typeof v==='string')overview.getRange(`B${i+4}`).formulas=[[v]];else overview.getRange(`B${i+4}`).values=[[v]];
});
overview.getRange('B4:B11').setNumberFormat('#,##0');
overview.getRange('A7:B7').format={fill:'#FFF2CC',font:{name:'Arial',size:10,bold:true}};
overview.getRange('A11:B11').format.fill='#FFF2CC';
overview.getRange('A14:E14').values=[['Khoản','Tiền hàng chưa VAT','VAT','Thanh toán','Cơ sở / giới hạn']];
overview.getRange('A14:E14').format=headerFormat;
overview.getRange('A15:E19').values=[
 ['Bán mới 2024, tính một bản mỗi chứng từ',...p.base_sales_money,'34.739 chứng từ; đã loại bản trùng cũ khỏi phép cộng'],
 ['163 HĐ điều chỉnh trả hàng',null,null,null,'Gồm 100 HĐ hợp lệ và 63 HĐ đang Đã gửi CQT'],
 ['2.000 HĐ thay thế mới',null,null,null,'Tiền hàng/VAT 0; 2.000 ô thanh toán nguồn trống'],
 ['Bán mới 2024 sau trả hàng',null,null,null,'Chưa chốt tổng thanh toán; cần bảng kê chi tiết mới'],
 ['Thanh toán dự kiến nếu 2.000 HĐ thay thế = 0',null,null,null,'Giá trị dự kiến có điều kiện, không thay ô trống nguồn thành 0'],
];
overview.getRange('B16').formulas=[[`=SUM('Dieu chinh 163'!H5:H${adLast})`]];
overview.getRange('C16').formulas=[[`=SUM('Dieu chinh 163'!K5:K${adLast})`]];
overview.getRange('D16').formulas=[[`=SUM('Dieu chinh 163'!N5:N${adLast})`]];
overview.getRange('B17').formulas=[[`=SUM('Thay the 2000'!I5:I${reLast})`]];
overview.getRange('C17').formulas=[[`=SUM('Thay the 2000'!J5:J${reLast})`]];
overview.getRange('B18').formulas=[['=SUM(B15:B17)']];
overview.getRange('C18').formulas=[['=SUM(C15:C17)']];
overview.getRange('D19').formulas=[['=SUM(D15:D16)']];
overview.getRange('B15:D19').setNumberFormat('#,##0;[Red](#,##0);0');
overview.getRange('A15:E19').format={rowHeight:42,wrapText:true};
overview.getRange('A18:E18').format.font={name:'Arial',size:10,bold:true};
overview.getRange('D17:D18').format.fill='#FFF2CC';
overview.getRange('A21').values=[['Phạm vi: phần bán mới 2024 sau trả hàng; 273 HĐ ngoài bộ và hóa đơn cũ giữ riêng chưa cộng vào bảng này.']];
overview.getRange('A22').values=[['63 HĐ cần cập nhật trạng thái: 1C26MUV/00150113–00150175.']];
overview.getRange('A23').values=[['Chưa có bảng kê chi tiết mới để kiểm dòng hàng, cờ diễn giải và thanh toán của 2.000 HĐ thay thế.']];
overview.getRange('A25').values=[['Nguồn: Bang_ke_hoa_don_da_su_dung_1793524594152.zip (hai phần bảng kê tổng).']];
overview.getRange('A26').values=[['Nguồn: Danh_sach_hoa_don_dieu_chinh.xlsx; Danh_sach_hoa_don_thay_the.xlsx.']];
overview.getRange('A27').values=[['Kế hoạch đối chiếu: bộ bán mới 2024, file import 163 phiếu trả hàng và bộ thay thế 2.000 HĐ đã kiểm.']];
overview.getRange('A29').values=[['Hỗ trợ trạng thái MISA:']];
overview.getRange('A30').values=[[p.status_help]];
overview.getRange('A25:A30').format.font={name:'Arial',size:10,color:'#525C66'};
w.recalculate();
const inspected=await w.inspect({kind:'region',sheetId:'Tong hop',range:'A4:D19',tableMaxRows:16,tableMaxCols:4,maxChars:4000});
const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:20},maxChars:1000});
await fs.writeFile(`${work}/inspection.ndjson`,`${inspected.ndjson}\n${errors.ndjson}`);
await fs.mkdir(output,{recursive:true});
const filename=`${output}/KET_QUA_KIEM_TRA_SAU_PHAT_HANH_20261101.xlsx`;
await(await SpreadsheetFile.exportXlsx(w)).save(filename);
for(const [sheetName,range,name]of[['Tong hop','A1:E19','summary-preview'],['Dieu chinh 163','D102:F108','adjustment-numbers-preview'],['Dieu chinh 163','P102:S108','pending-status-preview'],['Thay the 2000','I4:O8','replacement-preview']]){
 const png=await w.render({sheetName,range,scale:1,format:'png'});
 await fs.writeFile(`${work}/${name}.png`,new Uint8Array(await png.arrayBuffer()));
}
console.log(inspected.ndjson);console.log(errors.ndjson);console.log(JSON.stringify({output:filename,adjustments:163,replacements:2000,pending_cqt:63}));
