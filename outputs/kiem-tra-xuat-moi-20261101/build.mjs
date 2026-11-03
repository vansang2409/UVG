import fs from 'node:fs/promises';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';

const root = 'D:/UVG';
const output = `${root}/outputs/kiem-tra-xuat-moi-20261101`;
const work = `${root}/scripts/data/new-sales-issued-20261101`;
const data = JSON.parse(await fs.readFile(`${work}/resolved-coverage.json`, 'utf8'));
const summary = JSON.parse(await fs.readFile(`${work}/summary.json`, 'utf8'));
const detail = JSON.parse(await fs.readFile(`${work}/detail-summary.json`, 'utf8'));
const detailPairs = JSON.parse(await fs.readFile(`${work}/detail-pairs.json`, 'utf8'));
const byFirstInvoice = new Map(detailPairs.map(p=>[p.first_invoice,p]));
const pairs = data.duplicate_pairs;
const w = Workbook.create();
const s = w.worksheets.add('Kiem tra trung');
s.showGridLines = false;
const end = 14 + pairs.length;
s.getRange(`A1:S${end}`).format = {font: {name:'Arial',size:10}, verticalAlignment:'center',rowHeight:22};
const widths = [8,24,25,25,25,16,23,21,24,23,19,18,18,18,23,26,26,25,23];
widths.forEach((v,i) => s.getRangeByIndexes(0,i,1,1).format.columnWidth = v);
s.getRange('A1').values = [['Kiểm tra bộ xuất mới 2024']];
s.getRange('A1').format.font = {name:'Arial',size:14,bold:true,color:'#24435C'};
s.getRange('A1:M1').format.rowHeight = 30;
s.getRange('A2').values = [['Bảng kê 31/10/2026–01/11/2026; so với bộ 34.739 chứng từ.']];
s.getRange('A2').format.font = {name:'Arial',size:10,italic:true};
s.getRange('A4:D8').values = [
  ['Chứng từ trong bộ',null,null,summary.plan.invoices],
  ['Cặp trùng chi tiết',null,null,null],
  ['Tiền hàng dư (đ)',null,null,null],
  ['VAT dư (đ)',null,null,null],
  ['Thanh toán dư (đ)',null,null,null],
];
s.getRange('D5').formulas = [[`=COUNTA(B15:B${end})`]];
s.getRange('D6').formulas = [[`=SUM(G15:G${end})`]];
s.getRange('D7').formulas = [[`=SUM(H15:H${end})`]];
s.getRange('D8').formulas = [[`=SUM(I15:I${end})`]];
s.getRange('D4:D8').setNumberFormat('#,##0');
s.getRange('A4:D8').format.rowHeight = 24;
s.getRange('A5:D5').format.font = {name:'Arial',size:10,bold:true,color:'#A32929'};
s.getRange('A8:D8').format.font = {name:'Arial',size:10,bold:true};
s.getRange('A10').values = [['HĐ 00129632 và 00129806: bảng chi tiết ghi hàng khuyến mại, thanh toán 0đ; bảng tổng vẫn trống ô thanh toán.']];
s.getRange('A11').values = [['Đã so 2.720 dòng mỗi lô: đủ 2.000 cặp trùng hàng và người mua. Giữ riêng 273 HĐ ngày 31/10 ngoài bộ 2024.']];
s.getRange('A12').values = [['Nguồn: bảng kê tổng 1793518718337 (2 phần), bảng chi tiết 20261101_143931, bộ import 2024 và ánh xạ mã đơn trong IMPORT MEINV 2024.']];
s.getRange('A13').values = [['33 chứng từ có đơn giá làm tròn đồng, thành tiền giữ nguyên. Xem tab Don gia va hang KM.']];
s.getRange('A13').format.font = {name:'Arial',size:10,italic:true};
s.getRange('A12').format.font = {name:'Arial',size:10,italic:true,color:'#525C66'};
const headers = ['STT','Số chứng từ AMIS','Mã đơn hàng','HĐ số nhỏ','HĐ số lớn','Ngày hóa đơn','Tiền hàng mỗi HĐ','VAT mỗi HĐ','Thanh toán mỗi HĐ','Trạng thái gửi CQT (cả 2)','Có mã CQT (cả 2)','Dòng HĐ số nhỏ','Dòng HĐ số lớn','Số dòng mỗi HĐ','Chi tiết hai HĐ','Phát hành HĐ số nhỏ','Phát hành HĐ số lớn','So bộ import','Người mua hai HĐ'];
s.getRange('A14:S14').values = [headers];
const values = pairs.map((p,i) => {
  const [a,b] = p.invoices;
  const d = byFirstInvoice.get(a.key);
  return [i+1,p.doc,p.order,a.key,b.key,new Date(`${a.date}T00:00:00Z`),...p.expected,a.cqt_status,a.cqt_code_present&&b.cqt_code_present?'Có':'Thiếu',a.source.row,b.source.row,d.first_lines,d.identical_product_lines?'Trùng hoàn toàn':'Khác',d.first_issue_times.join('; '),d.second_issue_times.join('; '),d.first_matches_plan&&d.second_matches_plan?'Khớp':'Đơn giá làm tròn',d.identical_buyer_information?'Trùng':'Khác'];
});
for(let i=0;i<values.length;i+=500) s.getRangeByIndexes(i+14,0,Math.min(500,values.length-i),19).values = values.slice(i,i+500);
for(const c of ['B','C','D','E']) s.getRange(`${c}15:${c}${end}`).setNumberFormat('@');
s.getRange(`F15:F${end}`).setNumberFormat('dd/mm/yyyy');
s.getRange(`G15:I${end}`).setNumberFormat('#,##0');
const table = s.tables.add(`A14:S${end}`,true,'DuplicatePairs2024');
table.style = 'TableStyleMedium2';
table.showFilterButton = true;
s.getRange('A14:S14').format = {fill:'#24435C',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:42,horizontalAlignment:'center',verticalAlignment:'center'};
s.freezePanes.freezeRows(14);
s.freezePanes.freezeColumns(3);
const other = w.worksheets.add('Don gia va hang KM');
other.showGridLines=false;
const adjustments = [
 ...detail.price_rounding.items.map(p=>[p.doc,p.invoice,'Đơn giá làm tròn',p.item,p.qty,Number(p.plan_price),Number(p.issued_price),p.goods,'Giữ nguyên thành tiền. Đơn giá làm tròn đến đồng trên bảng kê.']),
 ...detail.promotion_representation_changes.map(p=>[p.doc,p.invoice,'Hàng khuyến mại',p.item,p.qty,0,null,null,'Bảng chi tiết ghi khuyến mại, thanh toán 0đ. Tiền hàng/VAT trên bảng tổng 0đ.']),
];
const last=adjustments.length+4;
other.getRange(`A1:J${last}`).format={font:{name:'Arial',size:10},verticalAlignment:'center',rowHeight:24};
[8,24,25,24,78,10,22,22,22,70].forEach((v,i)=>other.getRangeByIndexes(0,i,1,1).format.columnWidth=v);
other.getRange('A1').values=[['Đơn giá làm tròn và hàng khuyến mại']];
other.getRange('A1').format.font={name:'Arial',size:14,bold:true,color:'#24435C'};
other.getRange('A2').values=[['33 chứng từ làm tròn đơn giá nhưng giữ thành tiền; 2 HĐ khuyến mại có thanh toán 0đ trên bảng chi tiết.']];
other.getRange('A3').values=[['Nguồn: cùng bộ import và bảng kê chi tiết ngày 31/10–01/11/2026.']];
other.getRange('A3').format.font={name:'Arial',size:10,italic:true};
other.getRange('A4:J4').values=[['STT','Số chứng từ AMIS','Số hóa đơn','Khác biệt','Tên hàng','Số lượng','Đơn giá bộ import','Đơn giá đã xuất','Thành tiền chi tiết','Ghi chú']];
other.getRange(`A5:J${last}`).values=adjustments.map((r,i)=>[i+1,...r]);
other.getRange(`B5:C${last}`).setNumberFormat('@');
other.getRange(`G5:H${last}`).setNumberFormat('#,##0.00');
other.getRange(`I5:I${last}`).setNumberFormat('#,##0');
other.getRange(`E5:E${last}`).format.wrapText=true;
other.getRange(`J5:J${last}`).format.wrapText=true;
other.getRange(`A5:J${last}`).format.rowHeight=52;
const otherTable=other.tables.add(`A4:J${last}`,true,'RoundingAndPromotion');
otherTable.style='TableStyleMedium2';
otherTable.showFilterButton=true;
other.getRange('A4:J4').format={fill:'#24435C',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:42,wrapText:true,horizontalAlignment:'center',verticalAlignment:'center'};
other.freezePanes.freezeRows(4);
other.freezePanes.freezeColumns(3);
w.recalculate();
const inspect = await w.inspect({kind:'region',sheetId:'Kiem tra trung',range:'A4:D8',maxChars:2000,tableMaxRows:5,tableMaxCols:4});
const errors = await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30},summary:'Formula error scan'});
await fs.writeFile(`${work}/artifact-inspection.ndjson`,`${inspect.ndjson}\n${errors.ndjson}`);
await (await SpreadsheetFile.exportXlsx(w)).save(`${output}/KIEM_TRA_XUAT_MOI_2024.xlsx`);
for(const [sheetName,range,name] of [['Kiem tra trung','A1:I18','preview-top'],['Kiem tra trung','N14:S18','preview-detail-check'],['Don gia va hang KM','B1:J8','preview-price'],['Don gia va hang KM','B40:J42','preview-promotion']]){
 const png = await w.render({sheetName,range,scale:1,format:'png'});
 await fs.writeFile(`${output}/${name}.png`,new Uint8Array(await png.arrayBuffer()));
}
console.log(inspect.ndjson);
console.log(errors.ndjson);
console.log(JSON.stringify({output:`${output}/KIEM_TRA_XUAT_MOI_2024.xlsx`,pairs:pairs.length}));
