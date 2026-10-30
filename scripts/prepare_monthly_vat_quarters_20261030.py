"""VAT theo nguồn và các bảng quý cho tab hàng hóa hiện hành; không sửa import."""
from pathlib import Path
from datetime import datetime
from decimal import Decimal
import json,re,hashlib
ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/monthly-vat-quarters-20261030'
f=json.loads((ROOT/'scripts/data/revenue-forecast-20261028/prepared.json').read_text(encoding='utf8'))
p=json.loads((ROOT/'scripts/data/amis-2024-linked-20261026/prepared.json').read_text(encoding='utf8'))
monthly=json.loads((ROOT/'scripts/data/monthly-goods-20261028/prepared.json').read_text(encoding='utf8'))
dec=lambda x:Decimal(str(x)) if x not in (None,'') else Decimal(0)
month=lambda s:datetime.fromisoformat(s[:19]).month
for s in f['manifest']:assert hashlib.sha256((ROOT/s['path']).read_bytes()).hexdigest()==s['sha256'],s['path']
for s in p['manifest']:assert hashlib.sha256((ROOT/s['file']).read_bytes()).hexdigest()==s['sha256'],s['file']
vat={m:{'amis':Decimal(0),'invoice':Decimal(0)} for m in range(1,13)}
docs={d['doc']:d for d in p['docs']}
for d in p['docs']:
    vat[month(d['date'])]['amis']+=sum(dec(l['values'][32]) for l in d['original_lines'])
for n in f['new_records']:vat[month(n['date'])]['invoice']+=dec(n['money'][1])
processed={tuple(r['root']) for r in f['adjustment_records']}|{tuple(r['root']) for r in f['manual_records']}
for h in f['old_invoices']:
    if tuple(h['root']) not in processed:vat[month(h['root_date'])]['invoice']+=dec(h['money'][1])
for r in f['returns_records']:
    match=re.findall(r'\bBTL\d+\b',r['reason'])
    assert len(set(match))==1,(r['index'],match)
    vat[month(docs[match[0]]['date'])]['invoice']+=dec(r['money'][1])
assert sum(x['amis'] for x in vat.values())==f['totals']['amis_net'][1]==3994254050
assert sum(x['invoice'] for x in vat.values())==f['totals']['forecast_after_manual'][1]==3994443568
assert sum(x['invoice']-x['amis'] for x in vat.values())==189518
data=[['2. VAT THEO THÁNG'],['Tháng','Sổ AMIS','Hóa đơn sau xử lý (dự tính)','Chênh HĐ − AMIS']]
for m,v in vat.items():
    rn=19+len(data)
    data.append([m,int(v['amis']),int(v['invoice']),f'=C{rn}-B{rn}'])
data.extend([['CẢ NĂM','=SUM(B21:B32)','=SUM(C21:C32)','=SUM(D21:D32)'],[]])
quarters=[]
for label,first,kind in [('3. DOANH THU THEO QUÝ',5,'revenue'),('4. VAT THEO QUÝ',21,'vat'),('5. GIÁ VỐN THEO QUÝ',61,'cost')]:
    data.extend([[label],['Quý','Sổ AMIS','Hóa đơn: phần đã có giá vốn' if kind=='cost' else 'Hóa đơn sau xử lý (dự tính)','Còn cần bổ sung' if kind=='cost' else 'Chênh HĐ − AMIS']])
    start=19+len(data)
    for q in range(1,5):
        a=first+(q-1)*3;b=a+2;rn=19+len(data)
        row=[f'Quý {q}',f'=SUM(B{a}:B{b})',f'=SUM(C{a}:C{b})',('Thiếu giá vốn HĐ 63 và 111' if q==1 else None) if kind=='cost' else f'=C{rn}-B{rn}']
        data.append(row)
    end=19+len(data)-1
    data.extend([['CẢ NĂM',f'=SUM(B{start}:B{end})',f'=SUM(C{start}:C{end})','Còn thiếu giá vốn HĐ 63 và 111' if kind=='cost' else f'=SUM(D{start}:D{end})'],[]])
    quarters.append({'kind':kind,'first_row':start,'last_row':end})
assert len(data)==40
def cell(v):return {} if v is None else {'userEnteredValue':{'numberValue':v} if isinstance(v,(int,float)) else {'formulaValue':v} if v.startswith('=') else {'stringValue':v}}
request={'updateCells':{'range':{'sheetId':700000020,'startRowIndex':18,'endRowIndex':58,'startColumnIndex':0,'endColumnIndex':4},'rows':[{'values':[cell(row[i] if i<len(row) else None) for i in range(4)]} for row in data],'fields':'userEnteredValue'}}
WORK.mkdir(parents=True,exist_ok=True)
(WORK/'request.json').write_text(json.dumps(request,ensure_ascii=False),encoding='utf8')
prepared={'monthly_vat':[{'month':m,**{k:int(v) for k,v in x.items()}} for m,x in vat.items()],'quarters':quarters,'rows':data,'totals':{'amis_vat':3994254050,'invoice_vat':3994443568,'vat_gap':189518},'scope':monthly['basis'],'source_hashes':f['manifest'],'missing_cost_invoices':[63,111]}
(WORK/'prepared.json').write_text(json.dumps(prepared,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps({'monthly_vat':prepared['monthly_vat'],'quarters':quarters,'insert_rows':40,'totals':prepared['totals']}))
