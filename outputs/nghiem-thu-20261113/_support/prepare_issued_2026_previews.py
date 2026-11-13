"""Lấy giá trị và định dạng thực từ file đã lưu để render các vùng kiểm tra nhỏ."""
from pathlib import Path
from zipfile import ZipFile
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter,range_boundaries
from datetime import datetime,date
from lxml import etree
import json,sys,re
sys.stdout.reconfigure(encoding='utf-8');O=Path(__file__).parent;CHAINS='--chains' in sys.argv;COMPACT='--compact' in sys.argv or CHAINS
results=O.parent if O.name=='_support' else O
p=(results if CHAINS else O)/('TONG_HOP_HD_2026_CHO_DON_2024_2025_CHUOI.xlsx' if CHAINS else 'TONG_HOP_HD_2026_CHO_DON_2024_2025_RUT_GON.xlsx' if COMPACT else 'TONG_HOP_HD_PHAT_HANH_2026_CHO_DON_2024_2025.xlsx')
d=json.loads((O/'issued-2026-report-data.json').read_text(encoding='utf-8'))
idx=next((i+6 for i,r in enumerate(d['bhh_rows']) if '\n' in r[12] and r[11]),6)
specs=[('Tổng hợp','Tổng hợp',list(range(1,22)),list(range(1,8))),('HĐ 2026','HĐ 2026',list(range(1,14)),list(range(1,12))),('BHH và CT liên quan','BHH và CT liên quan',list(range(1,13)),list(range(1,9))),('Chuỗi thiếu BHH','Chuỗi thiếu BHH',list(range(1,13)),list(range(1,8))),('Liên kết BHH','BHH và CT liên quan',[5,idx,idx+1,idx+2],[2,12,13])]
if COMPACT:specs=[('Tổng hợp','Tổng hợp',list(range(1,20)),list(range(1,8))),('HĐ 2026','HĐ 2026',list(range(1,14)),list(range(1,14))),('BHH và CT liên quan','BHH và CT liên quan',list(range(1,13)),list(range(1,8))),('Chuỗi thiếu BHH','Chuỗi thiếu BHH',list(range(1,13)),list(range(1,8))),('Liên kết BHH','BHH và CT liên quan',[5,idx,idx+1,idx+2],[2,5,6,7])]
if CHAINS:
 chain=json.loads((O/'issued-2026-chain-view.json').read_text(encoding='utf-8'))
 selected=[5,6]+[i+7 for i,r in enumerate(chain['rows']) if r[0] in [34968,46137,46139]]
 specs=[('Chuỗi F0 F1 F2','Chuỗi F0 F1 F2',list(range(1,12)),list(range(1,10))),('F0 F1 ví dụ','Chuỗi F0 F1 F2',selected,list(range(1,10))),('F2 F3 ví dụ','Chuỗi F0 F1 F2',selected,[1,2,3,10,11,12,13,14,15,16,17])]
w=load_workbook(p,read_only=True,data_only=True);sheets=[]
def col(color):
 if color and color.type=='rgb':return '#'+color.rgb[-6:]
 return None
with ZipFile(p) as z:
 merge_cache={}
 for name,source,rr,cc in specs:
  sn=w.sheetnames.index(source)+1;heights={};widths={};default=23
  with z.open(f'xl/worksheets/sheet{sn}.xml') as f:
   for _,e in etree.iterparse(f,events=('end',)):
    tag=etree.QName(e).localname
    if tag=='sheetFormatPr':default=float(e.get('defaultRowHeight','23'))
    elif tag=='col':
     for c in range(int(e.get('min')),int(e.get('max'))+1):widths[c]=float(e.get('width'))
    elif tag=='row':
     n=int(e.get('r'));heights[n]=float(e.get('ht',default))
     if n>=max(rr):break
     e.clear()
     while e.getprevious() is not None:del e.getparent()[0]
  if sn not in merge_cache:
   tail=b''
   with z.open(f'xl/worksheets/sheet{sn}.xml') as f:
    while True:
     data=f.read(1048576)
     if not data:break
     tail=(tail+data)[-32768:]
   merge_cache[sn]=[m.decode() for m in re.findall(rb'<mergeCell ref="([^"]+)"',tail)]
  merges=[]
  for merge in merge_cache[sn]:
   c1,r1,c2,r2=range_boundaries(merge)
   if set(range(c1,c2+1))<=set(cc) and set(range(r1,r2+1))<=set(rr):
    merges.append(f'{get_column_letter(cc.index(c1)+1)}{rr.index(r1)+1}:{get_column_letter(cc.index(c2)+1)}{rr.index(r2)+1}')
  cells={};ws=w[source]
  for rn,source_row in enumerate(ws.iter_rows(min_row=min(rr),max_row=max(rr),max_col=max(cc)),min(rr)):
   if rn not in rr:continue
   for c in cc:
    cell=source_row[c-1];v=cell.value
    if not hasattr(cell,'coordinate'):continue
    borders={edge:{'style':getattr(cell.border,edge).style,'color':col(getattr(cell.border,edge).color) or '#D7E0E8'} for edge in ['top','bottom','left','right'] if getattr(cell.border,edge).style}
    cells[cell.coordinate]={'value':v.isoformat() if isinstance(v,(datetime,date)) else v,'isDate':isinstance(v,(datetime,date)),'font':{'name':cell.font.name or 'Arial','size':cell.font.sz or 10,'bold':bool(cell.font.bold),'italic':bool(cell.font.italic),'color':col(cell.font.color) or '#233348'},'fill':col(cell.fill.fgColor) if cell.fill.patternType else None,'format':cell.number_format,'wrap':bool(cell.alignment.wrap_text),'align':cell.alignment.horizontal or 'left','borders':borders}
  out=[]
  from openpyxl.utils import get_column_letter
  for r in rr:out.append([cells.get(get_column_letter(c)+str(r),{'value':None}) for c in cc])
  sheets.append({'name':name,'source':source,'rows':out,'heights':[heights.get(r,default) for r in rr],'widths':[widths.get(c,12) for c in cc],'source_columns':cc,'source_rows':rr,'merges':merges})
w.close();(O/('issued-2026-chain-preview-cells.json' if CHAINS else 'issued-2026-compact-preview-cells.json' if COMPACT else 'issued-2026-preview-cells.json')).write_text(json.dumps(sheets,ensure_ascii=False),encoding='utf-8');print('Extracted actual workbook views',len(sheets),flush=True)
