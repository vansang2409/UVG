import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const dir='D:/UVG/outputs/chuoi-sau-xu-ly-20261030';
const p=JSON.parse(await fs.readFile(`${dir}/prepared.json`,'utf8'));
function col(n){let out='';for(n++;n;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out;}
const w=Workbook.create();
for(const [index,sp] of p.sheets.entries()){
 const s=w.worksheets.add(sp.name),last=9+sp.rows.length,lastCol=col(8+6*sp.stages);
 s.showGridLines=false;s.tabColor=index?'#426B60':'#294D6A';
 s.getRange(`A1:${lastCol}${last}`).format={font:{name:'Arial',size:11},verticalAlignment:'center',rowHeight:26,columnWidth:17};
 s.getRange('A2').values=[[`Chuỗi hóa đơn ${sp.name.toLowerCase()} sau xử lý`]];
 s.getRange('A2').format.font={name:'Arial',size:15,bold:true,color:'#223A4B'};
 s.getRange('A2:H2').format.rowHeight=32;
 s.getRange('A3').values=[['Chuỗi xử lý']];s.getRange('B3').values=[[sp.processed]];
 s.getRange('D3').values=[['Đã về 0']];s.getRange('E3').formulas=[[`=COUNTIFS(H10:H${last},"Đã về 0")`]];
 s.getRange('G3').values=[['KT loại trừ']];s.getRange('H3').formulas=[[`=COUNTIFS(H10:H${last},"KT loại trừ")`]];
 s.getRange('A4').values=[['Phát hành']];s.getRange('B4').values=[[46324]];
 s.getRange('B4').setNumberFormat('dd/mm/yyyy');
 s.getRange('D4').values=[['Đơn vị: đồng']];
 s.getRange('A5').values=[[p.source_note]];
 s.getRange('A6').values=[[index?'Chuỗi thay thế nối từ gốc thật; số dư chỉ lấy HĐ cuối, không cộng các bản đã bị thay thế.':'F1–F7 xếp theo ngày/số hóa đơn; các HĐ có thể cùng tham chiếu F0, không mặc định F2 điều chỉnh F1.']];
 s.getRange('A7').values=[['Thanh toán nguồn trống chỉ dùng 0 ở các ô có căn cứ đã chốt; xem ghi chú cuối dòng.']];
 s.getRange('A5:H7').format={font:{name:'Arial',size:10,color:'#596772'},wrapText:false,rowHeight:20};
 const heads=['STT','Nhóm xử lý','Ký hiệu gốc','Số HĐ gốc','Dư tiền hàng','Dư VAT','Dư thanh toán','Kết quả'];
 const fields=['Ký hiệu','Số hóa đơn','Ngày','Tiền hàng','VAT','Thanh toán'];
 for(let f=0;f<sp.stages;f++)heads.push(...fields.map(x=>`${x} F${f}`));
 heads.push('Ghi chú');
 s.getRange(`A9:${lastCol}9`).values=[heads];
 s.getRange(`A10:${lastCol}${last}`).values=sp.rows.map(x=>x.values.slice(0,-1));
 const forms=sp.rows.map((r,i)=>{
  const rn=i+10;
  const sums=[0,1,2].map(c=>r.replacement?`=${col(8+6*(r.keys.length-1)+3+c)}${rn}`:`=SUM(${r.keys.map((_,f)=>`${col(8+6*f+3+c)}${rn}`).join(',')})`);
  return [...sums,`=IF(B${rn}="KT loại trừ","KT loại trừ",IF(AND(E${rn}=0,F${rn}=0,G${rn}=0),"Đã về 0","Còn chênh"))`];
 });
 s.getRange(`E10:H${last}`).formulas=forms;
 const table=s.tables.add(`A9:${lastCol}${last}`,true,index?'ChuoiThayThe':'ChuoiDieuChinh');
 table.style='TableStyleLight9';table.showFilterButton=true;
 s.getRange(`A9:${lastCol}9`).format={fill:'#294D6A',font:{name:'Arial',size:11,bold:true,color:'#FFFFFF'},rowHeight:48,wrapText:true,horizontalAlignment:'center'};
 s.getRange('A8:H8').format={fill:'#E7EDF1',font:{name:'Arial',size:11,bold:true},rowHeight:26};
 s.getRange('A8').values=[['Kết quả sau xử lý']];
 for(let f=0;f<sp.stages;f++){
  const start=8+6*f,end=start+5,color=f%2?'#426B60':'#294D6A';
  s.mergeCells(`${col(start)}8:${col(end)}8`);
  s.getRange(`${col(start)}8:${col(end)}8`).format={fill:color,font:{name:'Arial',size:11,bold:true,color:'#FFFFFF'},horizontalAlignment:'center'};
  s.getRange(`${col(start)}8`).values=[[f===0?'F0 — Hóa đơn gốc':`F${f} — ${index?'Hóa đơn thay thế':'Hóa đơn liên quan'}`]];
  s.getRange(`${col(start)}9:${col(end)}9`).format.fill=color;
  for(const c of [start,start+1]){
   s.getRange(`${col(c)}10:${col(c)}${last}`).setNumberFormat(c===start?'@':'00000000');
   s.getRange(`${col(c)}9:${col(c)}${last}`).format.columnWidth=c===start?15:16;
  }
  s.getRange(`${col(start+2)}10:${col(start+2)}${last}`).setNumberFormat('dd/mm/yyyy');
  s.getRange(`${col(start+2)}9:${col(start+2)}${last}`).format.columnWidth=15;
  s.getRange(`${col(start+3)}10:${col(end)}${last}`).setNumberFormat('#,##0;[Red]-#,##0;0');
  s.getRange(`${col(start+3)}9:${col(end)}${last}`).format.columnWidth=19;
 }
 s.getRange(`A9:A${last}`).format.columnWidth=9;
 s.getRange(`B9:B${last}`).format.columnWidth=23;
 s.getRange(`C9:D${last}`).format.columnWidth=16;
 s.getRange(`D10:D${last}`).setNumberFormat('00000000');
 s.getRange(`E9:G${last}`).format.columnWidth=19;
 s.getRange(`E10:G${last}`).setNumberFormat('#,##0;[Red]-#,##0;0');
 s.getRange(`H9:H${last}`).format.columnWidth=18;
 const note=col(8+6*sp.stages);
 s.getRange(`${note}9:${note}${last}`).format.columnWidth=85;
 s.getRange(`${note}10:${note}${last}`).format.wrapText=true;
 for(let i=0;i<sp.rows.length;i++)if(sp.rows[i].values[8+6*sp.stages])s.getRange(`A${i+10}:${lastCol}${i+10}`).format.rowHeight=48;
 s.getRange(`E10:G${last}`).conditionalFormats.addCustom('AND($B10<>"KT loại trừ",E10<>0)',{fill:'#FDE9E7',font:{color:'#A52828',bold:true}});
 s.getRange(`A10:H${last}`).conditionalFormats.addCustom('$B10="KT loại trừ"',{fill:'#FFF2CC',font:{color:'#80551D'}});
 s.freezePanes.freezeRows(9);s.freezePanes.freezeColumns(4);
 console.log(JSON.stringify({sheet:s.name,rows:sp.rows.length,stages:sp.stages}));
}
w.recalculate();
for(const sp of p.sheets){
 console.log((await w.inspect({kind:'table',range:`'${sp.name}'!A3:H4`,include:'values,formulas',tableMaxRows:2,tableMaxCols:8,maxChars:1800})).ndjson);
}
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:10},maxChars:1200})).ndjson);
for(const [sheet,range,name] of [['Điều chỉnh','A2:N13','dieu-chinh'],['Thay thế','A2:N15','thay-the'],['Thay thế','AA8:AG15','thay-the-cuoi'],['Điều chỉnh','BC8:BE13','dieu-chinh-cuoi'],['Điều chỉnh','A9432:H9434','loai-tru']]){
 const png=await w.render({sheetName:sheet,range,scale:1,format:'png'});
 await fs.writeFile(`${dir}/${name}.png`,new Uint8Array(await png.arrayBuffer()));
}
const out=`${dir}/CHUOI_F0_F1_SAU_XU_LY_20261029.xlsx`;
await (await SpreadsheetFile.exportXlsx(w)).save(out);
console.log(JSON.stringify({output:out}));
