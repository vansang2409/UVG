import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook, SpreadsheetFile, FileBlob} from '@oai/artifact-tool';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const out = path.join(root, 'outputs');
const helper = path.join(root, 'scripts/data');
const data = JSON.parse(await fs.readFile(path.join(helper, 'invoice_reports.json'), 'utf8'));
const colors = ['#243746', '#315F77', '#526579'];
const moneyFormat = '#,##0;[Red](#,##0);0';
const reports = [
  {key:'adjusted', sheet:'Dieu chinh', file:'HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx', identity:false},
  {key:'exceptions', sheet:'Ngoai le', file:'NGOAI_LE_F0_2024_T6_2025.xlsx', identity:true},
];
const requestedReport=process.argv.find(a=>a.startsWith('--report='))?.split('=')[1];
if(requestedReport && !reports.some(r=>r.key===requestedReport))throw new Error('Unknown report');
function statusNote(reportKey) {
  const general='Trạng thái nguồn là hiện tại, có thể phản ánh xử lý năm 2026. Xét phạm vi bằng ngày chứng từ và liên kết, không chỉ nhãn trạng thái.';
  if(reportKey!=='adjusted' || !data.user_confirmed_dates?.length)return general;
  return general+' Ngày 1C25TUV/00012885 ở bảng chính = 05/07/2025 theo xác nhận và ảnh danh sách meInvoice 07/10/2026; tab nguồn giữ ngày bảng tổng 18/08/2025. Chưa đối chiếu XML.';
}
async function saveReport(w,report) {
  const file=await SpreadsheetFile.exportXlsx(w), destination=path.join(out,report.file);
  await file.save(destination);
  try {await fs.rename(destination+'.inspect.ndjson',path.join(helper,report.file+'.inspect.ndjson'));}
  catch(error) {if(error.code!=='ENOENT')throw error;}
  console.log(JSON.stringify({saved:report.file,...data.summary[report.key]}));
}
function col(n) { let r=''; for(n++;n;n=Math.floor((n-1)/26)) r=String.fromCharCode(65+(n-1)%26)+r; return r; }
function day(t) { const [d,m,y]=t.split('/').map(Number); return new Date(Date.UTC(y,m-1,d)); }
function width(s,c,n,rows) { s.getRangeByIndexes(0,c,rows,1).format.columnWidth=n; }
function base(s,rows,cols) {
  s.showGridLines=false;
  const r=s.getRangeByIndexes(0,0,rows,cols);
  r.format.font.name='Arial'; r.format.font.size=10; r.format.verticalAlignment='center';
}
if(process.argv.includes('--preview-confirmed-date') || process.argv.includes('--apply-confirmed-date')) {
  const report=reports[0], groups=data.groups.adjusted;
  const index=groups.findIndex(g=>g.root[0]==='1C25TUV' && g.root[1]==='00001314');
  if(index<0)throw new Error('Missing chain 1314');
  const group=groups[index], row=index+3;
  const raw=groups.flatMap(g=>g.detail), rawIndex=raw.findIndex(r=>r[2]==='1C25TUV' && r[3]==='00012885');
  const w=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(out,report.file)));
  const s=w.worksheets.getItem('Dieu chinh'), t=w.worksheets.getItem('Chi tiet nguon'), g=w.worksheets.getItem('Huong dan');
  if(s.getRange(`B${row}`).values[0][0]!=='00001314' || t.getRange(`D${rawIndex+2}`).values[0][0]!=='00012885')
    throw new Error('Imported workbook row identities differ');
  const apply=process.argv.includes('--apply-confirmed-date');
  if(apply) {
    const invoice=group.invoices.find(r=>r[0]==='1C25TUV' && r[1]==='00012885');
    if(invoice[2]!=='05/07/2025' || !data.user_confirmed_dates?.length)throw new Error('Missing date confirmation');
    s.getRange(`I${row}`).values=[[day(invoice[2])]];
    s.getRange(`AB${row}`).values=[[group.notes]];
    s.getRange(`A${row}:AB${row}`).format.rowHeight=Math.max(28,Math.ceil(group.notes.length/145)*15+8);
    t.getRange(`M${rawIndex+2}`).values=[[raw[rawIndex][12]]];
    t.getRange(`A${rawIndex+2}:M${rawIndex+2}`).format.rowHeight=Math.max(28,Math.ceil(raw[rawIndex][12].length/155)*15+8);
    g.getRange('A9').values=[[statusNote('adjusted')]];g.getRange('A9').format.rowHeight=52;
    w.recalculate();
  }
  const suffix=apply?'after':'before';
  for(const [sheetName,range,label] of [['Dieu chinh',`A${row-1}:L${row+1}`,'date'],['Dieu chinh',`Y${row}:AB${row}`,'note'],['Chi tiet nguon',`H${rawIndex+2}:M${rawIndex+2}`,'source'],['Huong dan','A8:A10','guide']]) {
    const preview=await w.render({sheetName,range,scale:1,format:'png'});
    await fs.writeFile(path.join(helper,`12885_${suffix}_${label}.png`),new Uint8Array(await preview.arrayBuffer()));
  }
  console.log((await w.inspect({kind:'table',range:`Dieu chinh!G${row}:L${row}`,include:'values',tableMaxRows:1,tableMaxCols:6,maxChars:1200})).ndjson);
  if(apply) {
    console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!',options:{useRegex:true,maxResults:10},maxChars:500})).ndjson);
    await saveReport(w,report);
  }
  process.exit(0);
}
const prepared=[];
for(const report of reports.filter(r=>!requestedReport || r.key===requestedReport)) {
  const groups=data.groups[report.key], summary=data.summary[report.key];
  if(!groups.length) throw new Error(`Empty report: ${report.key}`);
  const blocks=summary.max_blocks, offset=report.identity?2:0;
  const headers=report.identity?['Tên người mua','MST/CCCD chủ hộ nguyên bản']:[];
  for(let i=0;i<blocks;i++) headers.push(...(i===0
    ?['Ký hiệu','Số hóa đơn','Ngày','Tiền hàng','Thuế','Thanh toán']
    :[`Ký hiệu sửa ${i}`,`Số hóa đơn sửa ${i}`,`Ngày sửa ${i}`,`Tiền hàng sửa ${i}`,`Thuế sửa ${i}`,`Thanh toán sửa ${i}`]));
  headers.push('Cộng dồn tiền hàng','Cộng dồn thuế','Cộng dồn thanh toán','Cần kiểm tra');
  const rows=groups.map(g=>{
    const r=report.identity?[g.name,g.tax]:[];
    for(let i=0;i<blocks;i++) {const b=g.invoices[i];r.push(...(b?[b[0],b[1],day(b[2]),...b.slice(4)]:Array(6).fill(null)));}
    r.push(...g.sums,g.notes||'');return r;
  });
  const last=rows.length+2, end=col(headers.length-1), sumStart=offset+blocks*6;
  const w=Workbook.create(), s=w.worksheets.add(report.sheet);
  s.getRange(`A2:${end}2`).values=[headers]; s.getRange(`A3:${end}${last}`).values=rows;
  const table=s.tables.add(`A2:${end}${last}`,true,report.identity?'NgoaiLeChuoi':'DieuChinhChuoi');
  table.style='TableStyleMedium2'; table.showFilterButton=true;
  base(s,last,headers.length);
  const head=s.getRange(`A1:${end}2`);
  head.format.fill=colors[0];head.format.font.color='#FFFFFF';head.format.font.bold=true;head.format.horizontalAlignment='center';
  s.getRange(`A2:${end}2`).format.wrapText=true;
  s.getRange(`A1:${end}1`).format.rowHeight=26;s.getRange(`A2:${end}2`).format.rowHeight=42;
  s.getRange(`A3:${end}${last}`).format.rowHeight=28;
  if(report.identity) {
    s.mergeCells('A1:B1');s.getRange('A1').values=[['Thông tin người mua']];
    width(s,0,60,last);width(s,1,28,last);
    s.getRange(`A3:A${last}`).format.wrapText=true;s.getRange(`B3:B${last}`).setNumberFormat('@');
  }
  for(let i=0;i<blocks;i++) {
    const start=offset+i*6, band=`${col(start)}1:${col(start+5)}1`;
    s.mergeCells(band);s.getRange(band).values=[[`F${i} — ${i===0?'Hóa đơn gốc':report.identity?'Hóa đơn liên quan':'Điều chỉnh'}`]];
    s.getRangeByIndexes(0,start,2,6).format.fill=colors[i]||colors[2];
    for(let j=0;j<6;j++)width(s,start+j,[16,14,16,20,18,21][j],last);
    s.getRangeByIndexes(2,start+1,groups.length,1).setNumberFormat('00000000');
    s.getRangeByIndexes(2,start+2,groups.length,1).setNumberFormat('dd/mm/yyyy');
    const money=s.getRangeByIndexes(2,start+3,groups.length,3);money.setNumberFormat(moneyFormat);money.format.horizontalAlignment='right';
  }
  s.mergeCells(`${col(sumStart)}1:${end}1`);s.getRange(`${col(sumStart)}1`).values=[['Giá trị cộng dồn và kiểm tra']];
  for(let i=0;i<3;i++)width(s,sumStart+i,25,last);width(s,sumStart+3,105,last);
  s.getRangeByIndexes(2,sumStart,groups.length,3).setNumberFormat(moneyFormat);
  s.getRangeByIndexes(2,sumStart,groups.length,3).format.horizontalAlignment='right';
  s.getRangeByIndexes(2,sumStart+3,groups.length,1).format.wrapText=true;
  for(let i=0;i<groups.length;i++) {
    const lines=Math.max(1,Math.ceil((groups[i].notes||'').length/145),report.identity?Math.ceil(groups[i].name.length/65):1);
    if(lines>1)s.getRange(`A${i+3}:${end}${i+3}`).format.rowHeight=Math.max(28,lines*15+8);
  }
  s.freezePanes.freezeRows(2);s.freezePanes.freezeColumns(offset+2);
  const notes=[
    `UVG – ${report.identity?'ngoại lệ giữ nguyên':'hóa đơn đã điều chỉnh ngoài ngoại lệ'}`,
    'F0 từ 01/01/2024 đến 30/06/2025; hóa đơn liên quan đến 31/12/2025.',
    `${summary.roots} F0, ${summary.invoices} hóa đơn riêng biệt; ${summary.roots_with_notes} chuỗi có ghi chú cần kiểm tra.`,
    `F0 năm 2024: ${summary.roots_by_year['2024']||0}; F0 tháng 1–6/2025: ${summary.roots_by_year['2025']||0}.`,
    report.identity?'Ngoại lệ nhận diện bằng tên công ty/doanh nghiệp hoặc định danh dạng MST trong nguồn; chưa xác thực MST.'
      :'Đã loại chuỗi ngoại lệ và chuỗi có liên kết thay thế trong kỳ; chỉ lấy chuỗi có liên kết điều chỉnh trong kỳ.',
    'F1/F2 và các khối tiếp theo xếp theo liên kết, rồi ngày. Không mặc định F2 sửa F1; xem loại/tham chiếu ở Chi tiet nguon.',
    'Cộng dồn là số tính sẵn theo nguồn. Không cộng bản bị thay thế khi có liên kết thay thế trong kỳ; không cộng hóa đơn nguồn ghi đã hủy.',
    report.identity?'Ô tiền thiếu giữ trống, không tự điền 0. Giá trị cộng dồn thiếu căn cứ được để trống và ghi ở Cần kiểm tra.'
      :'Ô tiền thiếu giữ trống, không tự điền 0. Chỉ riêng 1C24TUV/00000574 tạm tính thanh toán gốc trống = 0 theo xác nhận; ô nguồn vẫn trống.',
    statusNote(report.key),
    `${data.source_summary.files} file nguồn có SHA-256 trong SOURCE_MANIFEST.json; ${data.source_summary.historical_links_added} liên kết lịch sử bổ sung từ Git có vị trí nguồn.`,
    `${data.source_summary.unresolved_roots} F0 ngoài ngoại lệ có nhãn đã xử lý nhưng chưa có liên kết đến 31/12/2025; không tự đưa vào nhóm đã điều chỉnh.`,
    'Đơn vị: đồng. Đây là bảng rà soát; chưa chốt nghiệp vụ phát hành, chưa tạo import, ký hoặc phát hành hóa đơn.',
  ];
  const g=w.worksheets.add('Huong dan');g.getRange(`A1:A${notes.length}`).values=notes.map(v=>[v]);
  base(g,notes.length,1);width(g,0,145,notes.length);
  g.getRange(`A1:A${notes.length}`).format.wrapText=true;g.getRange(`A1:A${notes.length}`).format.rowHeight=34;g.getRange('A1').format.font.bold=true;
  if(report.key==='adjusted' && data.user_confirmed_dates?.length)g.getRange('A9').format.rowHeight=52;
  const detail=groups.flatMap(g=>g.detail).map(r=>[...r.slice(0,4),day(r[4]),...r.slice(5)]);
  const t=w.worksheets.add('Chi tiet nguon');
  const th=['Ký hiệu gốc','Số gốc','Ký hiệu','Số hóa đơn','Ngày','Tên người mua','MST/CCCD chủ hộ nguyên bản','Loại / tham chiếu',
    'Trạng thái nguồn hiện tại','Tiền hàng','Thuế','Thanh toán','Vị trí nguồn'];
  t.getRange('A1:M1').values=[th];t.getRange(`A2:M${detail.length+1}`).values=detail;
  const rawTable=t.tables.add(`A1:M${detail.length+1}`,true,report.identity?'NguonNgoaiLe':'NguonDieuChinh');
  rawTable.style='TableStyleMedium2';rawTable.showFilterButton=true;base(t,detail.length+1,13);
  for(let i=0;i<13;i++)width(t,i,[16,14,16,14,16,55,28,54,35,20,18,21,110][i],detail.length+1);
  t.getRange('A1:M1').format.fill=colors[0];t.getRange('A1:M1').format.font.color='#FFFFFF';t.getRange('A1:M1').format.font.bold=true;
  t.getRange('A1:M1').format.wrapText=true;t.getRange('A1:M1').format.rowHeight=38;
  t.getRange(`A2:M${detail.length+1}`).format.rowHeight=28;
  t.getRange(`B2:B${detail.length+1}`).setNumberFormat('00000000');t.getRange(`D2:D${detail.length+1}`).setNumberFormat('00000000');
  t.getRange(`E2:E${detail.length+1}`).setNumberFormat('dd/mm/yyyy');t.getRange(`G2:G${detail.length+1}`).setNumberFormat('@');
  t.getRange(`J2:L${detail.length+1}`).setNumberFormat(moneyFormat);
  t.getRange(`F2:I${detail.length+1}`).format.wrapText=true;t.getRange(`M2:M${detail.length+1}`).format.wrapText=true;
  for(let i=0;i<detail.length;i++) {
    const r=detail[i], lines=Math.max(1,Math.ceil((r[5]||'').length/60),Math.ceil((r[7]||'').length/65),Math.ceil((r[12]||'').length/155));
    if(lines>1)t.getRange(`A${i+2}:M${i+2}`).format.rowHeight=Math.max(28,lines*15+8);
  }
  t.freezePanes.freezeRows(1);w.recalculate();
  const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:20},maxChars:1500});
  console.log(JSON.stringify({report:report.key,errors:errors.ndjson}));
  for(const [sheetName,range,suffix] of [[report.sheet,`A1:${end}6`,'main'],['Huong dan',`A1:A${notes.length}`,'guide'],['Chi tiet nguon','A1:M5','source']]) {
    const image=await w.render({sheetName,range,scale:1,format:'png'});
    await fs.writeFile(path.join(helper,`${report.key}_${suffix}.png`),new Uint8Array(await image.arrayBuffer()));
  }
  const readable=await w.render({sheetName:report.sheet,range:report.identity?'A1:H6':'A1:L6',scale:1,format:'png'});
  await fs.writeFile(path.join(helper,`${report.key}_readable.png`),new Uint8Array(await readable.arrayBuffer()));
  const inspection=await w.inspect({kind:'table',range:`${report.sheet}!${col(sumStart)}2:${end}5`,include:'values,formulas',tableMaxRows:4,tableMaxCols:4,maxChars:1200});
  console.log(JSON.stringify({report:report.key,summary,inspection:inspection.ndjson}));
  prepared.push({report,w});
}
await fs.mkdir(out,{recursive:true});
for(const {report,w} of prepared) {
  await saveReport(w,report);
}
