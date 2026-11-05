"""Xuất theo luồng sau khi Artifact Tool export vượt giới hạn heap4GB và6GB."""
import json
from pathlib import Path
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree
P=Path(__file__).resolve().parents[1]/'outputs/phan-loai-2025-20261105'
d=json.loads((P/'workbook-data-compact.json').read_text(encoding='utf-8'))
wb=Workbook(write_only=True);summary=wb.create_sheet('Tổng hợp');specs=[]
body=Font(name='Arial',size=10,color='202124');bold=Font(name='Arial',size=10,bold=True,color='202124');title=Font(name='Arial',size=15,bold=True,color='202124')
gray=PatternFill('solid',fgColor='E8EAED');yellow=PatternFill('solid',fgColor='FFF8DC');number='#,##0;[Red](#,##0);0'
def cell(ws,value,font=body,fmt=None,fill=None,wrap=False):
    c=WriteOnlyCell(ws,value=value);c.font=font;c.alignment=Alignment(vertical='center',wrap_text=wrap)
    if isinstance(value,str) and not value.startswith('='):c.data_type='s'
    if isinstance(value,str) and value.startswith('https://'):
        c.hyperlink=value;c.font=Font(name='Arial',size=10,color='0563C1',underline='single')
    if fmt:c.number_format=fmt
    if fill:c.fill=fill
    return c
def setup(ws,widths):
    ws.sheet_view.showGridLines=False;ws.sheet_format.defaultRowHeight=36
    for i,w in enumerate(widths,1):ws.column_dimensions[get_column_letter(i)].width=w
    for r,h in [(1,12),(2,28),(3,24),(4,24),(5,12),(6,42)]:ws.row_dimensions[r].height=h
def add(ws,values,head=False,formats=None,editable=()):
    ws.append([cell(ws,v,font=bold if head else body,fmt=(formats or {}).get(i),fill=gray if head else yellow if i in editable else None,wrap=head or i in (7,8,18)) for i,v in enumerate(values)])
for t in d['tables']:
    ws=wb.create_sheet(t['name']);isreview=t['name']=='HĐ cần kiểm riêng';end=len(t['rows'])+6;specs.append({'name':t['name'],'rows':len(t['rows']),'columns':len(t['headers']),'end':end})
    widths=[25,15,26,21,19,30,21,48,60,22,34,55,36,43,50,25,25,23,52] if not isreview else [31,15,26,30,24,22,24,60,41,31,22,38,52,29,55,70]
    setup(ws,widths);ws.freeze_panes='B7';ws.auto_filter.ref=f'A6:{get_column_letter(len(t["headers"]))}{end}'
    dv=DataValidation(type='list',formula1='"Đã kiểm tra,Cần thêm hồ sơ,Đã thực hiện"',allow_blank=True)
    dv.sqref=f'{"K" if isreview else "J"}7:{"K" if isreview else "J"}{end}';ws.data_validations.append(dv)
    ws.append([]);ws.append([cell(ws,'Ứng viên xuất lại / bổ sung' if t['name']=='Xuất bổ sung' else t['name'],font=title)])
    note='Đề xuất theo mã đơn, tiền và VAT. Chưa so chi tiết hàng hóa.' if t['name']=='Giữ HĐ' else 'Chuỗi đã về 0. Cần xác minh nghiệp vụ và HĐ khác trước khi lập.' if t['name']=='Xuất bổ sung' else 'Kế toán ghi nhận tại các cột kết luận và ghi chú.'
    if t['name']=='Rà soát HĐ':note='Rà soát tiền/VAT và chuỗi. Sửa liên kết nghi ghép nhầm trước khi xử lý HĐ.'
    if isreview:note='Kiểm riêng theo HĐ, không cộng vào tổng chứng từ AMIS.'
    ws.append([cell(ws,note)]);ws.append([cell(ws,'Đối soát 05/11/2026. Tiền nhóm/chuỗi lặp theo CT, không cộng dồn.')]);ws.append([]);add(ws,t['headers'],head=True)
    fmts={1:'dd/mm/yyyy',2:'@'}
    for c in ([3,4,6,15,16,17] if not isreview else [4,5,6]):fmts[c]=number
    for r in t['rows']:add(ws,r,formats=fmts,editable=(10,11) if isreview else (9,10))
    print(json.dumps({'written':t['name'],'rows':len(t['rows'])},ensure_ascii=False),flush=True)
setup(summary,[44,20,27,27,27,24,23]);summary.sheet_format.defaultRowHeight=24
vals={1:[],2:['Đối soát và kế hoạch xử lý HĐ 2025'],3:['Theo chứng từ AMIS 2025. Giữ HĐ đúng, xử lý sai/thiếu. Phần 2024 giữ nguyên.'],4:['Bản phân loại 05/11/2026. Bảng làm việc, chưa phải lệnh phát hành.'],6:['Nhóm đề xuất','Số chứng từ','Thanh toán AMIS (đ)','VAT AMIS (đ)','Trước thuế AMIS (đ)','KT đã ghi nhận']}
for i,t in enumerate(specs[:5],7):
    ref="'"+t['name'].replace("'","''")+"'";e=t['end']
    vals[i]=['Ứng viên xuất lại / bổ sung' if t['name']=='Xuất bổ sung' else t['name'],f'=COUNTA({ref}!A7:A{e})',f'=SUM({ref}!D7:D{e})',f'=SUM({ref}!E7:E{e})',f'=C{i}-D{i}',f'=COUNTA({ref}!J7:J{e})']
vals[12]=['Tổng chứng từ AMIS','=SUM(B7:B11)','=SUM(C7:C11)','=SUM(D7:D11)','=SUM(E7:E11)','=SUM(F7:F11)']
vals[15]=['Theo tháng hạch toán AMIS'];vals[17]=['Tháng','Giữ HĐ','Sửa liên kết','Rà soát HĐ','Ứng viên xuất bổ sung','Cần KT xác minh','Tổng chứng từ']
groups=['Giữ HĐ','Sửa liên kết','Xử lý HĐ sai','Xuất bổ sung','Cần KT xác minh']
for m in range(1,13):
    counts=[d['monthly'].get(str(m),{}).get(c,{}).get('docs',0) for c in groups];vals[17+m]=[m,*counts,sum(counts)]
vals[32]=['Cách sử dụng và giới hạn']
for r,v in enumerate(d['limits'],34):vals[r]=[v]
vals.update({42:['Nguồn Google Sheets',d['source_url']],43:['Bảng kê meInvoice',d['source_cutoff']],44:['HĐ cần kiểm riêng',specs[5]['rows']],45:['HĐ KT đã note bỏ',25],46:['Tổng thanh toán nguồn AMIS',d['source_total']],47:['Tổng VAT nguồn AMIS',d['source_vat']]})
for r in range(1,48):
    values=vals.get(r,[])
    if r==2:summary.append([cell(summary,values[0],font=title)]);continue
    fmts={i:number for i in (2,3,4)} if 7<=r<=12 else {1:number} if r in (46,47) else {}
    add(summary,values,head=r in (6,17),formats=fmts)
wb.save(P/'PHAN_LOAI_XU_LY_HD_2025.xlsx')
# Gắn cache tổng hợp đã đối chiếu nguồn, vẫn giữ công thức để Excel tính lại khi sửa.
cache={}
for i,t in enumerate(specs[:5],7):
    cat='Xử lý HĐ sai' if t['name']=='Rà soát HĐ' else t['name'];s=d['summary'][cat]
    for c,v in zip('BCDEF',[s['docs'],s['money'],s['vat'],s['money']-s['vat'],0]):cache[f'{c}{i}']=v
for c,v in zip('BCDEF',[65494,d['source_total'],d['source_vat'],d['source_total']-d['source_vat'],0]):cache[f'{c}12']=v
file=P/'PHAN_LOAI_XU_LY_HD_2025.xlsx';tmp=P/'streaming-cache.tmp'
with ZipFile(file) as zi,ZipFile(tmp,'w',ZIP_DEFLATED) as zo:
    for item in zi.infolist():
        raw=zi.read(item.filename)
        if item.filename=='xl/worksheets/sheet1.xml':
            xml=etree.fromstring(raw);ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for c in xml.findall('.//s:c',ns):
                if c.get('r') in cache and c.find('s:f',ns) is not None:
                    v=c.find('s:v',ns)
                    if v is None:v=etree.SubElement(c,'{'+ns['s']+'}v')
                    v.text=str(cache[c.get('r')])
            raw=etree.tostring(xml)
        zo.writestr(item,raw)
tmp.replace(file)
(P/'workbook-specs.json').write_text(json.dumps(specs,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'exported':True,'bytes':(P/'PHAN_LOAI_XU_LY_HD_2025.xlsx').stat().st_size}))
