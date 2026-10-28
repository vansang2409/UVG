"""Tạo dữ liệu đối chiếu chênh, giữ riêng nhóm chưa có BHH và giá trị bằng 0."""
from pathlib import Path
from collections import defaultdict
import json,hashlib
ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/revenue-gap-20261028'
p=json.loads((WORK/'prepared.json').read_text(encoding='utf8'))
f=json.loads((ROOT/'scripts/data/revenue-forecast-20261028/prepared.json').read_text(encoding='utf8'))
linked=json.loads((ROOT/'scripts/data/amis-2024-linked-20261026/prepared.json').read_text(encoding='utf8'))
for m in linked['manifest']:
    assert hashlib.sha256((ROOT/m['file']).read_bytes()).hexdigest()==m['sha256'],m['file']
old_by_root=defaultdict(list)
for r in f['old_invoices']:old_by_root['/'.join(r['root'])].append(r)
headers=['Nhóm','HĐ gốc','HĐ trong chuỗi được cộng','BHH AMIS đã loại','Người mua','Tiền hàng cũ giữ','Tiền hàng AMIS loại','Chênh tiền hàng','VAT cũ giữ','VAT AMIS loại','Chênh VAT','Thanh toán cũ giữ','Thanh toán AMIS loại','Chênh thanh toán','Căn cứ / nhận xét','Hướng xử lý đề xuất','Nguồn chứng cứ','Tình trạng']
rows=[]
assert not p['outside_chains'],'Phát sinh chuỗi giữ ngoài danh sách cần rà lại'
for no in ['00000003','00000478','00000498','00000574']:
    rt='1C24TUV/'+no;old_rows=old_by_root[rt]
    adjustments=[r for r in f['adjustment_records'] if '/'.join(r['root'])==rt]
    balance=[sum(r['money'][i] for r in old_rows+adjustments) for i in range(3)]
    assert balance==[0,0,0],(rt,balance)
    evidence='; '.join(x['source']+' | dòng '+str(x.get('row','')) for x in old_rows)
    evidence+='; '+ '; '.join('outputs/'+x['file']+' | Import dieu chinh | dòng '+str(x['row'])+' | STT '+str(x['index']) for x in adjustments)
    evidence+='; accountant-decisions.json/latest_override_20261019; apply-13-20261019/final-verification.json'
    rows.append(['Đã điều chỉnh về 0',rt,'\n'.join('/'.join(x['key']) for x in old_rows),'',old_rows[0]['buyer'],0,0,None,0,0,None,0,0,None,'Bộ import vừa pull đã có khoản giảm cho toàn chuỗi theo quyết định 19/10.','Đã tính vào phần điều chỉnh về 0 của báo cáo mới; không giữ lại hoặc giảm thêm lần nữa.',evidence,'Đã chuẩn bị trong import; chưa xác minh đã phát hành'])
for g in sorted(p['groups'],key=lambda r:(-abs(r['difference'][0]),r['roots'])):
    cross='BHH2023005090' in g['codes']
    label='Khác kỳ' if cross else 'Đã ghép theo danh sách'
    basis='Đã ghép theo mã BHH trên danh sách người dùng, quy về gốc thật theo liên kết thay thế.'
    action='Giữ nguyên nguồn; đã khớp tiền.'
    state='Khớp'
    if any(g['difference']):
        if cross:
            basis='HĐ 1C24TUV/111 lập 15/03/2024, danh sách ghi BHH2023005090. Note: đơn ghi nhận 2023, xuất hóa đơn 2024 => sai thời điểm.'
            action='KT kiểm tra kỳ ghi nhận 2023/2024 trước khi quyết định xử lý; không tự thêm vào doanh số AMIS 2024.'
            state='Khác kỳ, cần KT kiểm tra'
        elif g['roots']==['1C24TUV/00000064']:
            basis='Chuỗi có 2609 (-482.407đ) và 9375 (+482.407đ), hai điều chỉnh triệt tiêu. Bảng HD cty chỉ trình bày đến 2609 nên số sau điều chỉnh tại bảng này chưa đủ chuỗi.'
            action='KT rà đủ HĐ 64/2609/9375 và giá AMIS trước khi chốt cách xử lý phần 482.407đ.'
            state='Lệch tiền hàng'
        elif abs(g['difference'][0])>1:
            basis+=' Tiền hóa đơn còn hiệu lực sau toàn chuỗi khác số AMIS.'
            action='KT đối chiếu giá/chiết khấu và đơn gốc; chốt căn cứ điều chỉnh hóa đơn hoặc sổ trước khi sửa.'
            state='Lệch tiền hàng'
        else:
            basis+=' Lệch làm tròn 1đ tại tiền hàng hoặc thuế/thanh toán.'
            action='Tách sai số làm tròn, giữ số nguồn để KT kiểm tra.'
            state='Lệch làm tròn'
    elif not g['amis_docs']:
        basis+=' HĐ 63 có giá trị 0; chưa có BHH, không tạo chênh tiền.'
        action='Không cần bù tiền; giữ dấu vết trường hợp thiếu BHH.'
    rs=[x for rt in g['roots'] for x in old_by_root[rt]]
    evidence='data/HD Xuất CTY 2024.xlsx | HD cty 2024 | dòng '+', '.join(map(str,g['exclusion_rows']))
    evidence+='; '+ '; '.join('AMIS '+d['doc']+' | dòng '+', '.join(map(str,d['rows'])) for d in g['amis_docs'])
    evidence+='; '+ '; '.join(x['source']+' | dòng '+str(x.get('row','')) for x in rs)
    im=g['invoice_money'];am=g['amis_money']
    rows.append([label,'\n'.join(g['roots']),'\n'.join('/'.join(x['key']) for x in rs),'\n'.join(d['doc'] for d in g['amis_docs']),'\n'.join(g['buyers']),im[0],am[0],None,im[1],am[1],None,im[2],am[2],None,basis,action,evidence,state])
assert len(rows)==118
assert sum(r[5]-r[6] for r in rows)==2368960
assert sum(r[8]-r[9] for r in rows)==189518
assert sum(r[11]-r[12] for r in rows)==2558478
cols={'retained':sum(r[5] for r in rows),'excluded':sum(r[6] for r in rows),'gap':sum(r[5]-r[6] for r in rows),'company_pair_diff':sum(r[5]-r[6] for r in rows if r[0]=='Đã ghép theo danh sách'),'cross_year':sum(r[5]-r[6] for r in rows if r[0]=='Khác kỳ')}
assert cols['company_pair_diff']==-1519929 and cols['cross_year']==3888889
(WORK/'review-sheet.json').write_text(json.dumps({'headers':headers,'rows':rows,'totals':cols},ensure_ascii=False),encoding='utf8')
SID=700000011
def cell(v):
    if v is None:return {}
    return {'userEnteredValue':{'numberValue':v} if isinstance(v,(int,float)) else {'stringValue':v}}
last=34+len(rows)
cross=[g for g in p['groups'] if 'BHH2023005090' in g['codes']]
assert len(cross)==1
cross=cross[0]
sixteen=sorted([g for g in p['groups'] if g['difference'][0] and g is not cross],key=lambda g:g['roots'][0])
assert len(sixteen)==16 and sum(g['difference'][0] for g in sixteen)==-1519929
detail_headers=['HĐ meInvoice (gốc)','Chứng từ AMIS','Khách hàng','Tiền HĐ sau điều chỉnh','Tiền AMIS năm 2024','Chênh HĐ − AMIS']
def detail(g):return ['\n'.join(g['roots']),'\n'.join(g['codes']),'\n'.join(g['buyers']),g['invoice_money'][0],g['amis_money'][0],None]
top=[
    ['CHI TIẾT CHÊNH DOANH THU CÒN LẠI: 2.368.960đ'],
    ['Đơn vị: đồng, chưa VAT. Tiền HĐ đã cộng đủ điều chỉnh liên quan. Chênh = tiền HĐ − tiền AMIS.'],
    ['Khoản','Số tiền chênh (đ)','Nội dung'],
    ['HĐ 111: đơn bán ghi nhận năm 2023',None,'Hóa đơn lập năm 2024; đơn bán theo danh sách thuộc năm 2023.'],
    ['16 nhóm khác tiền hàng',None,'Từng hóa đơn, mã BHH, khách hàng và số tiền trình bày ngay bên dưới.'],
    ['TỔNG CHÊNH CÒN LẠI',None,'3.888.889 − 1.519.929 = 2.368.960đ.'],
    [],
    ['1. HÓA ĐƠN LIÊN QUAN ĐƠN BÁN NĂM 2023'],
    detail_headers,
    detail(cross),
    ['AKIA: HĐ lập ngày 15/03/2024. Danh sách ghi BHH2023005090 thuộc năm 2023, không có trong sổ AMIS 2024; cột AMIS ở đây bằng 0 theo phạm vi năm 2024.'],
    [],
    ['2. CHI TIẾT 16 NHÓM KHÁC TIỀN HÀNG'],
    detail_headers,
]
top.extend(detail(g) for g in sixteen)
top.extend([
    ['TỔNG 16 NHÓM',None,None,None,None,None],
    ['TỔNG CHÊNH: HĐ 111 + 16 NHÓM',None,None,None,None,None],
    ['Riêng HĐ 64: điều chỉnh 1C25TUV/2609 giảm 482.407đ, sau đó 1C25TUV/9375 cộng lại 482.407đ. Cộng đủ chuỗi, tiền hàng HĐ còn 2.407.407đ.'],
])
assert len(top)==33
top.append(headers)
out=[{'values':[cell(v) for v in r]} for r in top]
formulas={3:'=F10',4:'=F31',5:'=SUM(B4:B5)'}
for i,fm in formulas.items():out[i]['values'][1]={'userEnteredValue':{'formulaValue':fm}}
for rn in [10]+list(range(15,31)):
    out[rn-1]['values'][5]={'userEnteredValue':{'formulaValue':f'=D{rn}-E{rn}'}}
for col,letter in [(3,'D'),(4,'E'),(5,'F')]:
    out[30]['values'][col]={'userEnteredValue':{'formulaValue':f'=SUM({letter}15:{letter}30)'}}
out[31]['values'][5]={'userEnteredValue':{'formulaValue':'=SUM(F10;F31)'}}
for rn,r in enumerate(rows,35):
    cells=[cell(v) for v in r]
    for col,fm in [(7,f'=F{rn}-G{rn}'),(10,f'=I{rn}-J{rn}'),(13,f'=L{rn}-M{rn}')]:cells[col]={'userEnteredValue':{'formulaValue':fm}}
    out.append({'values':cells})
for i in range(0,len(out),40):
    request={'updateCells':{'start':{'sheetId':SID,'rowIndex':i,'columnIndex':0},'rows':out[i:i+40],'fields':'userEnteredValue'}}
    (WORK/f'write-{i//40:02}.json').write_text(json.dumps(request,ensure_ascii=False,separators=(',',':')),encoding='utf8')
print(json.dumps({'data_rows':len(rows),'last_row':last,'chunks':(len(out)+39)//40,'totals':cols},ensure_ascii=False))
