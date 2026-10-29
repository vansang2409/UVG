"""Rà chênh 1–2 đồng ở nhóm chưa điền mã; chỉ đọc snapshot đã xác minh."""
import json,re,unicodedata,sys
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal,ROUND_HALF_UP
ROOT=Path(__file__).resolve().parents[1]
DIR=ROOT/'scripts/data/nb-fill-20261022'
def norm(x):return re.sub(r'\s+',' ',unicodedata.normalize('NFC',str(x or '')).strip()).lower()
def inv(x):
 s=str(x or '').strip()
 if s.endswith('.0'):s=s[:-2]
 return s.lstrip('0').zfill(8) if s.isdigit() else None
def dong(x):
 if x is None or x=='':return None
 try:return int(Decimal(str(x)).quantize(Decimal('1'),rounding=ROUND_HALF_UP))
 except Exception:return None
main={r['row']:r for p in sorted((DIR/'source').glob('main-*.json')) for r in json.loads(p.read_text(encoding='utf-8')) if r['row']>1}
nb=[r for p in sorted((DIR/'source').glob('nb-*.json')) for r in json.loads(p.read_text(encoding='utf-8')) if r['row']>1]
skip=json.loads((DIR/'summary-before.json').read_text(encoding='utf-8'))['skips']
idx=defaultdict(list)
for r in nb:
 v=r['v']
 if len(v)<18:continue
 idx[(inv(v[0]),norm(v[5]))].append(r)
result=[];seen_same_key=0
for s in skip:
 r=main[s['row']];v=r['v'];m=list(map(dong,[v[8],v[9],v[10]]))
 if any(x is None for x in m):continue
 same=idx.get((inv(v[6]),norm(v[0])),[])
 seen_same_key+=bool(same)
 candidates=[]
 for n in same:
  nv=n['v'];nm=list(map(dong,[nv[15],nv[16],nv[17]]))
  if any(x is None for x in nm):continue
  delta=[a-b for a,b in zip(m,nm)]
  if max(map(abs,delta))<=2 and any(delta) and str(nv[6] or '').strip():
   candidates.append(dict(nb_row=n['row'],code=str(nv[6]).strip(),nb_money=nm,delta_meinvoice_minus_nb=delta,max_delta=max(map(abs,delta))))
 if candidates:
  codes=sorted({x['code'] for x in candidates})
  # Nếu có khớp tiền chính xác nhưng nhiều mã, không dùng lệch nhỏ để bỏ qua mâu thuẫn đó.
  exact_codes={str(n['v'][6]).strip() for n in same if n['v'][6] and [dong(n['v'][k]) for k in (15,16,17)]==m}
  result.append(dict(row=r['row'],invoice=[v[5],v[6]],buyer=v[0],main_money=m,
   previous_reason=s['reason'],candidate_codes=codes,exact_codes=sorted(exact_codes),
   unique_code=len(set(codes)|exact_codes)==1,candidates=candidates))
stats=dict(remaining_checked=len(skip),same_number_and_buyer=seen_same_key,
 near_1_or_2_dong=len(result),unique_code=sum(x['unique_code'] for x in result),
 ambiguous=sum(not x['unique_code'] for x in result),
 within_1_dong=sum(any(c['max_delta']==1 for c in x['candidates']) for x in result),
 requires_2_dong=sum(all(c['max_delta']==2 for c in x['candidates']) for x in result))
out=dict(rule='Same invoice number and NFC/case/space-normalized buyer; compare all three independently rounded integer VND amounts; each absolute delta <=2 and at least one nonzero. No Sheet edits.',stats=stats,cases=result)
(DIR/'rounding-audit.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
sys.stdout.reconfigure(encoding='utf-8');print(json.dumps(stats,ensure_ascii=False));print(json.dumps(result[:5],ensure_ascii=False,indent=2))
