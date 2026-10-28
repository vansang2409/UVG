import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const here=path.dirname(fileURLToPath(import.meta.url));
const root=path.resolve(here,'../../..');
await fs.mkdir(path.join(here,'previews'),{recursive:true});
if(process.argv.includes('--before')){
  for(const [name,views] of [
    ['HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx',[
      ['Import dieu chinh','J41796:X41800','before-import'],
      ['Dieu chinh','Y3:AI3','before-plan'],
      ['Cho KT xac nhan','G1:I5','before-pending']]],
    ['NGOAI_LE_F0_2024_T6_2025.xlsx',[
      ['Ngoai le','A78:N79','before-exceptions']]]]){
    console.log('LOAD_BEFORE',name);
    const w=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(here,'before',name)));
    for(const [sheetName,range,label] of views){
      const b=await w.render({sheetName,range,scale:1,format:'png'});
      await fs.writeFile(path.join(here,'previews',label+'.png'),new Uint8Array(await b.arrayBuffer()));
    }
  }
  console.log('BEFORE_PREVIEWS_DONE');
  process.exit(0);
}
if(process.argv.includes('--after')||process.argv.includes('--after-summary')){
  for(const [name,views] of [
    ['HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx',[
      ['Import dieu chinh','O41800:S41803','final-products'],
      ['Dieu chinh','BA917:BC917','final-zero'],
      ['Huong dan','A13:A15','final-guide'],
      ['Cho KT xac nhan','G3:H3','final-pending']]],
    ['NGOAI_LE_F0_2024_T6_2025.xlsx',[
      ['Ngoai le','A76:H79','final-exceptions'],['Huong dan','A3:A5','final-exceptions-guide']]],
    ['UVG_72_HOA_DON_KHONG_CAN_DIEU_CHINH_20261019.xlsx',[
      ['Tong hop','A5:D9','final-list72']]]]){
    if(process.argv.includes('--after-summary')&&name!=='UVG_72_HOA_DON_KHONG_CAN_DIEU_CHINH_20261019.xlsx')continue;
    console.log('LOAD_AFTER',name);
    const w=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(here,'staged',name)));
    for(const [sheetName,range,label] of views){
      const b=await w.render({sheetName,range,scale:1,format:'png'});
      await fs.writeFile(path.join(here,'previews',label+'.png'),new Uint8Array(await b.arrayBuffer()));
    }
  }
  console.log('AFTER_PREVIEWS_DONE');process.exit(0);
}
const p=JSON.parse(await fs.readFile(path.join(here,'plan.json'),'utf8'));
const col=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const cash='#,##0;[Red](#,##0);0';
const targetBooks=p.files.filter(f=>f.kind!=='compact');
await fs.mkdir(path.join(here,'authored'),{recursive:true});
for(const f of targetBooks){
  const w=f.kind==='list72'?await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(here,'before',f.name))):Workbook.create();
  if(f.kind!=='list72')for(const name of Object.keys(f.sheets))w.worksheets.add(name);
  for(const [name,sp] of Object.entries(f.sheets)){
    const s=w.worksheets.getItem(name);
    if(f.kind!=='list72'){
      s.showGridLines=false;
      if(sp.header?.length){s.getRange(`A${sp.header_row}:${col(sp.header.length-1)}${sp.header_row}`).values=[sp.header];s.getRange(`A${sp.header_row}:${col(sp.header.length-1)}${sp.header_row}`).format={fill:'#22364D',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:40,wrapText:true};}
      for(const [c,width] of Object.entries(sp.widths||{}))if(Number.isFinite(width))s.getRange(`${c}1`).format.columnWidth=width;
      for(const r of sp.display_rows||[]){
        s.getRange(`A${r.row}:AA${r.row}`).values=[r.values.map((v,i)=>[2,8,14,20].includes(i)&&v?new Date(v):v)];
      }
    }
    for(const b of sp.blocks){
      const values=b.values.map(row=>row.map((v,i)=>b.date_columns.includes(i)&&v?new Date(String(v).slice(0,10)+'T00:00:00Z'):v));
      s.getRange(b.range).values=values;
      s.getRange(b.range).format={font:{name:'Arial',size:10},verticalAlignment:'center',wrapText:true,rowHeight:70};
      const start=Number(b.range.match(/\d+/)[0]),end=Number(b.range.split(':').at(-1).match(/\d+/)[0]);
      const firstLetters=b.range.match(/^[A-Z]+/)[0];let first=0;for(const c of firstLetters)first=first*26+c.charCodeAt(0)-64;
      for(const i of b.date_columns)s.getRange(`${col(first-1+i)}${start}:${col(first-1+i)}${end}`).setNumberFormat('dd/mm/yyyy');
    }
    for(const c of sp.changes)s.getRange(c.cell).values=[[c.value]];
    for(const c of sp.formulas)s.getRange(c.cell).formulas=[[c.formula]];
    for(const [row,height] of Object.entries(sp.row_heights||{}))s.getRange(`A${row}`).format.rowHeight=height;
    if(name==='Import dieu chinh'){
      s.getRange('K41800:K42005').setNumberFormat('00000000');
      s.getRange('Q41800:W42005').setNumberFormat(cash);
      s.getRange('O41800:O42005').format.wrapText=true;
    }else if(name==='Dieu chinh'){
      for(const r of p.root_checks){
        for(const c of ['Y','Z','AA','AF','AG','AH','AL','AM','AN','BA','BB','BC'])s.getRange(`${c}${r.row}`).setNumberFormat(cash);
        s.getRange(`AB${r.row}:BC${r.row}`).format={font:{name:'Arial',size:10},wrapText:true,rowHeight:130};
      }
    }
  }
  w.recalculate();
  const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30}});
  await fs.writeFile(path.join(here,`errors-${f.kind}.ndjson`),errors.ndjson);
  const views=f.kind==='adjusted'?[['Import dieu chinh','O41800:S41803','after-products'],['Import dieu chinh','T41800:W41803','after-amounts'],['Dieu chinh','AC3:AH3','after-plan'],['Dieu chinh','BA917:BC917','after-zero'],['Cho KT xac nhan','G3:H3','after-pending']]:f.kind==='exceptions'?[['Ngoai le','A76:H79','after-exceptions']]:[['Tong hop','A5:D9','after-list72']];
  for(const [sheetName,range,label] of views){const b=await w.render({sheetName,range,scale:1,format:'png'});await fs.writeFile(path.join(here,'previews',label+'.png'),new Uint8Array(await b.arrayBuffer()));}
  await(await SpreadsheetFile.exportXlsx(w)).save(path.join(here,'authored',f.name));
  console.log('AUTHORED',f.kind);
}
