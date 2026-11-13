import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const out=path.dirname(fileURLToPath(import.meta.url));
const data=JSON.parse(await fs.readFile(path.join(out,'reverse-report-data.json'),'utf8'));
const wb=Workbook.create();const summary=wb.worksheets.add('Tổng hợp');
const col=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const fmt='#,##0;[Red](#,##0);0';
const specs=[['Cần rà bán hàng',data.main,'Bán hàng cần rà soát'],['Theo dõi riêng',data.special,null],['CQT cần kiểm',data.review,'CQT không hợp lệ']];
const ends={};
for(const [name,input,fixedGroup] of specs){
 const s=wb.worksheets.add(name),rows=input.map(r=>[...r]);const end=rows.length+6;ends[name]=end;
 for(const r of rows){
   r[1]=new Date(r[1]+'T00:00:00Z');
   if(fixedGroup)r[5]=fixedGroup;
   else if(!['Khác kỳ — AMIS2023','Ngoại lệ KT đã chốt'].includes(r[5])){r[12]+=' | '+r[5];r[5]='24 HĐ theo dõi riêng';}
   else r[5]=r[5]==='Khác kỳ — AMIS2023'?'Khác kỳ AMIS2023':'Ngoại lệ kế toán';
 }
 s.showGridLines=false;s.getRange('A2').values=[[name==='Cần rà bán hàng'?'5 HĐ bán hàng cần rà soát':name==='Theo dõi riêng'?'Các trường hợp đã theo dõi riêng':'HĐ có trạng thái CQT không hợp lệ']];
 s.getRange('A3').values=[['Nguồn meInvoice 13/11/2026; đối chiếu sổ AMIS 2024–2025.']];
 s.getRange('A4').values=[[name==='Cần rà bán hàng'?'BHH liên quan hiện đã gắn HĐ khác; chưa kết luận HĐ trùng.':name==='Theo dõi riêng'?'Phân biệt biếu tặng/công ty, khác kỳ và ngoại lệ kế toán.':'Tách ngoài nhóm 5 HĐ bán hàng; cần kiểm trạng thái CQT.']];
 s.getRange('A6:M6').values=[data.headers];s.getRangeByIndexes(6,0,rows.length,13).values=rows;
 s.tables.add(`A6:M${end}`,true,'Reverse'+(specs.findIndex(x=>x[0]===name)+1));
 s.getRange(`A1:M${end+1}`).format.font={name:'Arial',size:10,color:'#202124'};
 s.getRange(`A1:M${end+1}`).format.verticalAlignment='center';s.getRange(`A1:A${end+1}`).format.rowHeight=24;
 s.getRange('A2').format.font={name:'Arial',size:15,bold:true,color:'#202124'};
 s.getRange('A3:A4').format.font={name:'Arial',size:10,italic:true,color:'#59636F'};
 s.getRange('A6:M6').format={fill:'#253B53',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:42,horizontalAlignment:'center',verticalAlignment:'center'};
 [235,105,160,145,175,245,195,255,220,185,155,280,650].forEach((w,i)=>s.getRange(`${col(i+1)}1`).format.columnWidthPx=w);
 s.getRange(`B7:B${end}`).setNumberFormat('dd/mm/yyyy');s.getRange(`C7:E${end+1}`).setNumberFormat(fmt);
 s.getRange(`A${end+1}`).values=[['Tổng']];s.getRange(`C${end+1}:E${end+1}`).formulas=[['C','D','E'].map(c=>`=SUM(${c}7:${c}${end})`)];s.getRange(`A${end+1}:E${end+1}`).format.font.bold=true;
 s.freezePanes.freezeRows(6);s.freezePanes.freezeColumns(1);s.tabColor=name==='Cần rà bán hàng'?'#253B53':name==='CQT cần kiểm'?'#C08226':'#83909E';
}
summary.showGridLines=false;summary.tabColor='#253B53';
summary.getRange('A2').values=[['meInvoice còn giá trị chưa nối sổ AMIS 2024–2025']];
summary.getRange('A3').values=[['Đối chiếu theo toàn bộ chuỗi F0 → F1 → F2 và các bước sau; nguồn 13/11/2026.']];
summary.getRange('A5:E5').values=[['Nhóm','Số chuỗi','Tiền hàng còn (đ)','VAT còn (đ)','Thanh toán còn (đ)']];
const groups=[['Bán hàng cần rà soát','Cần rà bán hàng'],['24 HĐ theo dõi riêng','Theo dõi riêng'],['Khác kỳ AMIS2023','Theo dõi riêng'],['Ngoại lệ kế toán','Theo dõi riêng']];
for(const [i,[label,name]] of groups.entries()){
 const rn=i+6,end=ends[name];summary.getRange(`A${rn}`).values=[[label]];
 summary.getRange(`B${rn}:E${rn}`).formulas=[[`=COUNTIFS('${name}'!F7:F${end},A${rn})`,...['C','D','E'].map(c=>`=SUMIFS('${name}'!${c}7:${c}${end},'${name}'!F7:F${end},A${rn})`)]];
}
summary.getRange('A10').values=[['Tổng chưa nối sổ AMIS2024/25']];summary.getRange('B10:E10').formulas=[['B','C','D','E'].map(c=>`=SUM(${c}6:${c}9)`)];
summary.getRange('A12').values=[['CQT không hợp lệ — tách ngoài']];summary.getRange('B12:E12').formulas=[[`=COUNTA('CQT cần kiểm'!A7:A${ends['CQT cần kiểm']})`,...['C','D','E'].map(c=>`=SUM('CQT cần kiểm'!${c}7:${c}${ends['CQT cần kiểm']})`)]];
summary.getRange('A15').values=[['Phạm vi và cách đọc']];
const notes=[
 '5 HĐ bán hàng còn 9.721.681đ; BHH liên quan hiện đã gắn HĐ khác. Chưa kết luận trùng hoặc chỉ định xử lý.',
 '24 HĐ biếu tặng/KOL/công ty đã được theo dõi riêng; không mặc định là sai hoặc thiếu chứng từ kế toán.',
 'HĐ 1C24TUV/00000111 thuộc BHH2023005090: có HĐ năm 2024 nhưng đơn ghi nhận năm 2023.',
 'HĐ 1C25TUV/00001943 là ngoại lệ KT đã chốt không xử lý; giữ nguyên quyết định.',
 'HĐ năm 2026 chỉ được xét khi thuộc chuỗi hoặc có mã đơn/chứng từ liên quan AMIS 2024–2025.',
 'Bỏ giá trị HĐ bị thay thế/hủy; cộng các HĐ hiện hành và điều chỉnh. Ô tiền trống không tự coi là 0.',
 'Đã bổ sung 31 quan hệ lịch sử thiếu trong bảng kê mới. Chuỗi 31054 → 34893 → 2639 đã xác nhận về 0.',
 'Kiểm độc lập đủ 31 HĐ còn tiền trực tiếp trong ZIP nguồn. Đây là đối chiếu file, chưa truy vấn CQT trực tiếp.'
];notes.forEach((v,i)=>summary.getRange(`A${17+i}`).values=[[v]]);
summary.getRange('A27').values=[['Nguồn']];data.source.forEach((v,i)=>summary.getRange(`A${29+i}`).values=[[v]]);
summary.getRange('A35').values=[['Bản chốt 2025']];summary.getRange('B35').values=[['https://docs.google.com/spreadsheets/d/1jACpDPSlOCjq-25hF_g5-06eZKD3e2Fq88j472wjTiA']];
summary.getRange('A1:F35').format.font={name:'Arial',size:10,color:'#202124'};summary.getRange('A1:A35').format.rowHeight=25;summary.getRange('A1:F35').format.verticalAlignment='center';
summary.getRange('A2').format.font={name:'Arial',size:15,bold:true,color:'#202124'};summary.getRange('A3').format.font.italic=true;
summary.getRange('A5:E5').format={fill:'#253B53',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:40,wrapText:true,horizontalAlignment:'center'};
summary.getRange('A10:E10').format.font.bold=true;summary.getRange('A15').format.font.bold=true;
[350,115,180,160,190,80].forEach((w,i)=>summary.getRange(`${col(i+1)}1`).format.columnWidthPx=w);
summary.getRange('B6:B12').setNumberFormat('#,##0');summary.getRange('C6:E12').setNumberFormat(fmt);
wb.recalculate();
const vals=summary.getRange('B6:E10').values;
if(vals[0][0]!==5||vals[0][3]!==9721681||vals[4][0]!==31||vals[4][3]!==452682122)throw new Error('Summary failed '+JSON.stringify(vals));
console.log((await wb.inspect({kind:'region',sheetId:'Tổng hợp',range:'A5:E12',maxChars:2500,tableMaxRows:8,tableMaxCols:5})).ndjson);
for(const name of ['Tổng hợp',...specs.map(s=>s[0])]){
 const blob=await wb.render({sheetName:name,range:name==='Tổng hợp'?'A1:E12':'A1:E12',scale:1,format:'png'});
 await fs.writeFile(path.join(out,`reverse-preview-${name}.png`),new Uint8Array(await blob.arrayBuffer()));
}
const f=await SpreadsheetFile.exportXlsx(wb);await f.save(path.join(out,'MEINVOICE_CHUA_NOI_AMIS_2024_2025_1311.xlsx'));
console.log('EXPORTED reverse report');
