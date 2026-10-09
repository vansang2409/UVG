import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const h='D:/UVG/scripts/data/check-review';const d=JSON.parse(await fs.readFile(`${h}/split-42-data.json`,'utf8'));
console.log('IMPORT');const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${h}/before-split-42.xlsx`));
const widths=[18,20,17,21,22,18,75,75,15,18,20,23,24,24,25,26,24,24,27,28,19,32,28,26,29,100,40];
for(const [name,raw] of Object.entries(d.groups)){
 const s=w.worksheets.add(name),last=raw.length+3;
 s.getRange('A1').values=[[name.includes('31')?'31 chuỗi khớp chi tiết và số tiền':'11 chuỗi cần kiểm tra cách điều chỉnh']];
 s.getRange('A2').values=[['Đối chiếu XML điều chỉnh với chi tiết gốc đã rà. Nguồn XML: data/Bo sung/meInvoice/2026-10-09-237-XML. Chưa xác minh hồ sơ nghiệp vụ và mật mã chữ ký.']];
 s.mergeCells('A1:AA1');s.mergeCells('A2:AA2');
 s.getRange('A3:AA3').values=[d.headers];
 const rows=raw.map(r=>{const v=[...r];v[2]=new Date(`${v[2]}T00:00:00Z`);v[5]=new Date(`${v[5]}T00:00:00Z`);return v;});
 s.getRange(`A4:AA${last}`).values=rows;
 s.getRange(`W4:Y${last}`).formulas=rows.map((_,i)=>{const n=i+4;return [`=O${n}+P${n}`,`=Q${n}+R${n}`,`=S${n}+T${n}`];});
 s.showGridLines=false;s.getRange(`A1:AA${last}`).format={font:{name:'Arial',size:10},verticalAlignment:'center'};
 s.getRange('A1').format={font:{name:'Arial',size:14,bold:true},rowHeight:28};s.getRange('A2').format={wrapText:true,rowHeight:30};
 s.getRange('A3:AA3').format={fill:'#243746',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:48,horizontalAlignment:'center'};
 s.getRange(`A4:AA${last}`).format.rowHeight=82;
 for(const range of [`G4:H${last}`,`V4:V${last}`,`Z4:AA${last}`])s.getRange(range).format.wrapText=true;
 widths.forEach((v,i)=>s.getRangeByIndexes(0,i,1,1).format.columnWidth=v);
 for(const col of ['B','E'])s.getRange(`${col}4:${col}${last}`).setNumberFormat('@');
 for(const col of ['C','F'])s.getRange(`${col}4:${col}${last}`).setNumberFormat('dd/mm/yyyy');
 s.getRange(`K4:U${last}`).setNumberFormat('#,##0;[Red](#,##0);0');s.getRange(`W4:Y${last}`).setNumberFormat('#,##0;[Red](#,##0);0');
 const t=s.tables.add(`A3:AA${last}`,true,name.includes('31')?'Chuoi42DaKhop31':'Chuoi42CanCheck11');t.style='TableStyleMedium2';t.showFilterButton=true;
 s.freezePanes.freezeRows(3);s.freezePanes.freezeColumns(2);
}
w.recalculate();
for(const [sheetName,file] of [['42 - Da khop 31','42-good.png'],['42 - Can check 11','42-review.png']]){
 const b=await w.render({sheetName,range:'V3:Z7',scale:1,format:'png'});await fs.writeFile(`${h}/${file}`,new Uint8Array(await b.arrayBuffer()));
 console.log((await w.inspect({kind:'table',range:`'${sheetName}'!W4:Y5`,include:'values',tableMaxRows:2,tableMaxCols:3,maxChars:700})).ndjson);
}
console.log('EXPORT');await(await SpreadsheetFile.exportXlsx(w)).save(`${h}/with-split-42.xlsx`);console.log('STAGED');
