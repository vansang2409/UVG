"""Tách nhóm chưa điều chỉnh bằng cùng quy tắc của hai report hiện có, không sửa nguồn."""
import sys,json,ast,hashlib
from pathlib import Path
from collections import Counter
r=Path(r'D:\UVG');h=r/'scripts/data/untreated';sys.path.insert(0,str(r/'scripts'))
import prepare_invoice_reports as core
src=(r/'scripts/prepare_invoice_reports.py').read_text(encoding='utf8')
src=src.replace("groups = {'exceptions': [], 'adjusted': []}","groups = {'exceptions': [], 'adjusted': [], 'untreated': []}")
old="""                continue
        pending, ordered, flags = set(nodes), [], []"""
new="""                    continue
                if 'huy' in status:
                    exclusions.append({'root': list(root), 'reason': 'Hóa đơn gốc nguồn ghi đã hủy; không đưa vào nhóm chưa điều chỉnh'})
                    continue
        pending, ordered, flags = set(nodes), [], []"""
assert src.count(old)==1
src=src.replace(old,new)
src=src.replace("groups['exceptions' if is_exception else 'adjusted'].append(group)","groups['exceptions' if is_exception else 'adjusted' if has_adjustment else 'untreated'].append(group)")
# Không cho workflow gốc ghi manifest hoặc dữ liệu helper đang dùng.
end=src.index('    HELPERS.mkdir(parents=True, exist_ok=True)')
src=src[:end]+'    return payload\n'
space={**core.__dict__,'__name__':'readonly_untreated'}
exec(compile(src,str(r/'scripts/prepare_invoice_reports.py'),'exec'),space)
data=space['main']()
assert data['summary']['exceptions']['roots']==238
assert data['summary']['adjusted']['roots']==1047
entries=data['groups']['untreated'];assert entries and all(len(x['invoices'])==1 for x in entries)
# Đối chiếu nhóm mới với hai bảng hiện dùng.
import openpyxl
for fn,sig,num in [('HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx',0,1),('NGOAI_LE_F0_2024_T6_2025.xlsx',2,3)]:
 p=r/'outputs'/fn
 w=openpyxl.load_workbook(p,read_only=True,data_only=True)
 keys={(v[sig],v[num]) for v in w.worksheets[0].iter_rows(min_row=3,values_only=True)}
 assert not keys&{tuple(x['root']) for x in entries}
counts=Counter(); money=[0,0,0]
for g in entries:
 m=g['invoices'][0][4:7]
 kind='Thiếu số tiền' if any(v is None for v in m) else 'Cả 3 khoản bằng 0' if all(v==0 for v in m) else 'Có giá trị khác 0'
 counts[kind]+=1
 for i,v in enumerate(m):
  if v is not None:money[i]+=v
 if all(v is not None for v in m) and m[0]+m[1]!=m[2]:counts['Lệch số học']+=1
 if 'huy' in core.normal(g['detail'][0][8]):raise AssertionError('Cancelled leaked')
payload={'groups':entries,'summary':data['summary']['untreated'],'scope':data['scope'],'source_summary':data['source_summary'],'hashes':{x['path']:x['sha256'] for x in data['manifest']},'historical_sources':data['historical_sources'],'exclusions':data['excluded'],'unresolved':data['unresolved'],'counts':dict(counts),'total_money':money,'existing_counts':{k:v for k,v in data['summary'].items() if k!='untreated'}}
(h/'untreated.json').write_text(json.dumps(payload,ensure_ascii=False),encoding='utf8')
print(json.dumps({k:payload[k] for k in ['summary','counts','total_money','source_summary']},ensure_ascii=False))
