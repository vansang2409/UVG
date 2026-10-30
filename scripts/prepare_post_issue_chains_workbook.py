"""Chuẩn bị chuỗi sau xử lý, chỉ đọc nguồn; giữ riêng HĐ KT loại trừ."""
from pathlib import Path
from collections import defaultdict,Counter
from datetime import datetime
import sys,json,hashlib
sys.stdout.reconfigure(encoding='utf-8')
from openpyxl.utils.datetime import to_excel
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from prepare_invoice_reports import rows,invoice_key,invoice_day,amount
OUT=ROOT/'outputs/chuoi-sau-xu-ly-20261030'
OUT.mkdir(parents=True,exist_ok=True)
A=ROOT/'scripts/data/post-issue-20261029/audit.json'
a=json.loads(A.read_text(encoding='utf-8'))
for path,h in a['manifest'].items():
 p=ROOT/path
 assert p.exists() and hashlib.sha256(p.read_bytes()).hexdigest()==h,('Nguồn thay đổi',path)
recs={tuple(r['key']):r for r in a['invoices']}
edge_map={(tuple(e['old']),tuple(e['new'])):e for e in a['edges']}
historical=ROOT/'data/Du lieu Meinvoice/Bang_ke_hoa_don_thay_the_xuat_theo_HD_bi_thay_the_1793158686512.xls'
historical_parents={};hist={}
for rn,r in rows(historical):
 if rn<7 or amount(r.get('A')) is None:continue
 old=invoice_key(r.get('B'),r.get('D'));new=invoice_key(r.get('M'),r.get('O'))
 if not old or not new:continue
 assert new not in historical_parents or historical_parents[new]==old
 historical_parents[new]=old
 for k,dc,cols in [(old,'C',['J','K','L']),(new,'N',['U','V','W'])]:
  money=[int(amount(r.get(c))) for c in cols]
  if k not in hist:hist[k]={'key':k,'date':invoice_day(r.get(dc)).isoformat(),'money':money,'resolved_money':money,'source':{'file':str(historical.relative_to(ROOT)),'row':rn},'fallbacks':[]}
  else:assert hist[k]['money']==money
# Ngày ở phía HĐ bị thay thế có thể khác ngày lập ở chặng trước; ưu tiên bảng tổng hiện hành.
upstream={k for g in a['groups']['replacement_six'] for k in [tuple(g['root'])]}
for k in list(upstream):
 while k in historical_parents:
  k=historical_parents[k];upstream.add(k)
for source in a['files']:
 if source['kind']!='total' or source['date_min']>'2025-05-31':continue
 p=ROOT/source['file']
 for rn,r in rows(p):
  if rn<6 or amount(r.get('A')) is None:continue
  key=invoice_key(r.get('B'),r.get('C'))
  if key not in upstream:continue
  m=[amount(r.get(c)) for c in ['P','Q','R']]
  assert all(x is not None for x in m)
  if key in hist:assert m==hist[key]['money']
  if key not in recs:
   recs[key]={'key':key,'date':invoice_day(r.get('D')).isoformat(),'money':[int(v) for v in m],'resolved_money':[int(v) for v in m],'status':r.get('U') or '', 'cqt_status':r.get('S') or '', 'source':{'file':source['file'],'row':rn},'fallbacks':[]}
 for k in upstream:
  if k not in recs and k in hist:continue

def serial(day):return to_excel(datetime.fromisoformat(day[:10]))
def ordered(g):
 root=tuple(g['root']);keys={tuple(k) for k in g['all_keys']}
 assert root in keys
 return [root]+sorted(keys-{root},key=lambda k:(recs[k]['date'],k))
def stage(k):
 r=recs[k];m=r['resolved_money'];assert None not in m
 return [k[0],k[1],serial(r['date']),*m]
def explanation(keys):
 notes=[]
 for i,k in enumerate(keys):
  r=recs[k]
  if r['money'][2] is None:notes.append(f'Thanh toán F{i}: nguồn trống, dùng 0 theo căn cứ đã chốt.')
 return ' '.join(notes)
def sources(keys):
 return '; '.join(f"F{i}: {recs[k]['source']['file']} | dòng {recs[k]['source']['row']}" for i,k in enumerate(keys))
def make_row(num,label,g,keys,maxstages,replacement=False,excluded=False):
 root=keys[0];note=explanation(keys)
 if excluded:note='KT loại khỏi đợt phát hành vì TĐ không hợp lệ; giữ số tiền nguồn, không tính vào tổng đã xử lý.'
 elif replacement:note='Chỉ tính giá trị HĐ cuối; các HĐ bị thay thế không cộng vào số dư. '+note
 vals=[num,label,root[0],root[1],None,None,None,None]
 for k in keys:vals+=stage(k)
 vals += [None]*(6*(maxstages-len(keys)))
 vals += [note or None,sources(keys)]
 expected=[597269,47782,645051] if excluded else [0,0,0]
 if replacement:assert stage(keys[-1])[-3:]==expected
 else:assert [sum(stage(k)[-3+c] for k in keys) for c in range(3)]==expected
 return {'values':vals,'keys':keys,'expected':expected,'excluded':excluded,'replacement':replacement}

adjusted=[('Đã bị điều chỉnh',g) for g in a['groups']['adjusted']]+[('Chưa điều chỉnh',g) for g in a['groups']['untreated']]
adjusted.sort(key=lambda x:(x[1]['root'][0],x[1]['root'][1]))
ex=a['accountant_exclusions'][0]['source_result']
maxadj=max(len(g['all_keys']) for _,g in adjusted)
adjrows=[make_row(i,label,g,ordered(g),maxadj) for i,(label,g) in enumerate(adjusted,1)]
adjrows.append(make_row(len(adjrows)+1,'KT loại trừ',ex,ordered(ex),maxadj,excluded=True))
replacement=[]
for g in a['groups']['replacement_six']:
 old=tuple(g['root']);anc=[old]
 while anc[0] in historical_parents:anc.insert(0,historical_parents[anc[0]])
 new=[tuple(k) for k in g['direct_replacements']];assert len(new)==1
 keys=anc+new
 assert all(k in recs for k in keys)
 assert all(recs[k]['date']<=recs[l]['date'] for k,l in zip(keys,keys[1:]))
 replacement.append((g,keys))
replacement.sort(key=lambda x:x[1][0])
maxrep=max(len(k) for _,k in replacement)
reprows=[make_row(i,'Thay thế',g,k,maxrep,replacement=True) for i,(g,k) in enumerate(replacement,1)]
assert len(adjrows)==9425 and len(reprows)==6 and maxadj==8 and maxrep==4
result={'date':'2026-10-29','source_note':'Bảng kê sau phát hành đến 29/10/2026; đối soát chốt 30/10/2026. Đơn vị: đồng.',
 'sheets':[{'name':'Điều chỉnh','stages':maxadj,'rows':adjrows,'processed':9424,'excluded':1}, {'name':'Thay thế','stages':maxrep,'rows':reprows,'processed':6,'excluded':0}],
 'source_hashes':{**a['manifest'],str(A.relative_to(ROOT)):hashlib.sha256(A.read_bytes()).hexdigest(),str(historical.relative_to(ROOT)):hashlib.sha256(historical.read_bytes()).hexdigest()}}
(OUT/'prepared.json').write_text(json.dumps(result,ensure_ascii=False),encoding='utf-8')
print(json.dumps({'sheets':[{'name':s['name'],'rows':len(s['rows']),'stages':s['stages'],'processed':s['processed']} for s in result['sheets']],'replacement_keys':[r['keys'] for r in reprows]},ensure_ascii=False))
