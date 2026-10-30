"""Bổ sung tiền hàng F0/F1/F2 và số HĐ cho 16 nhóm, không sửa số tổng hiện hành."""
from pathlib import Path
import json,hashlib
ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/gap-f-stages-20261030'
f=json.loads((ROOT/'scripts/data/revenue-forecast-20261028/prepared.json').read_text(encoding='utf8'))
gap=json.loads((ROOT/'scripts/data/revenue-gap-20261028/prepared.json').read_text(encoding='utf8'))
current=json.loads((ROOT/'scripts/data/report-post-issue-20261030/prepared.json').read_text(encoding='utf8'))
for path,digest in current['source_hashes'].items():assert hashlib.sha256((ROOT/path).read_bytes()).hexdigest()==digest,path
evidence={tuple(x['root']):x for x in current['evidence']}
groups=sorted([g for g in gap['groups'] if g['difference'][0] and 'BHH2023005090' not in g['codes']],key=lambda g:g['roots'][0])
assert len(groups)==16
headers=['F0 — Tiền hàng','F1 — Tiền hàng','F2 — Tiền hàng','F0 — Số HĐ','F1 — Số HĐ','F2 — Số HĐ']
matrix=[headers];audit=[]
for rn,g in enumerate(groups,15):
    assert len(g['roots'])==1
    root=g['roots'][0];root_key=tuple(root.split('/'))
    es=evidence[root_key]
    nodes=sorted([x for x in f['old_invoices'] if tuple(x['root'])==root_key],key=lambda x:(tuple(x['key'])!=root_key,x['date'],x['key']))
    assert {tuple(x['key']) for x in nodes}==set(map(tuple,es['all_keys']))==set(map(tuple,es['active']))
    assert es['kind']=='retained' and not es['new']
    assert len(nodes)<=3 and tuple(nodes[0]['key'])==root_key
    assert sum(x['money'][0] for x in nodes)==g['invoice_money'][0]==es['after_goods']
    matrix.append([nodes[i]['money'][0] if i<len(nodes) else None for i in range(3)]+['/'.join(nodes[i]['key']) if i<len(nodes) else None for i in range(3)])
    audit.append({'row':rn,'root':root,'nodes':nodes,'total':es['after_goods']})
assert sum(x['total'] for x in audit)==29066273
def cell(v):return {} if v is None else {'userEnteredValue':{'numberValue':v} if isinstance(v,(int,float)) else {'stringValue':v}}
request={'updateCells':{'range':{'sheetId':700000011,'startRowIndex':13,'endRowIndex':30,'startColumnIndex':7,'endColumnIndex':13},'rows':[{'values':[cell(v) for v in r]} for r in matrix],'fields':'userEnteredValue'}}
WORK.mkdir(parents=True,exist_ok=True)
(WORK/'prepared.json').write_text(json.dumps({'headers':headers,'rows':matrix[1:],'audit':audit,'source_hashes':current['source_hashes'],'meaning':'F0 là HĐ gốc; F1/F2 là từng HĐ điều chỉnh theo ngày, số tiền mang dấu nguồn. Không phát sinh để trống.'},ensure_ascii=False,indent=2),encoding='utf8')
(WORK/'request.json').write_text(json.dumps(request,ensure_ascii=False),encoding='utf8')
print(json.dumps({'groups':16,'invoices':sum(len(x['nodes']) for x in audit),'stages':3,'sum':29066273,'source_hashes_verified':len(current['source_hashes'])}))
