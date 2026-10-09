from pathlib import Path
from decimal import Decimal as D
from datetime import datetime
import json, shutil, hashlib, openpyxl
h=Path(__file__).resolve().parent;root=h.parents[2]
p=root/'outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx'
source=json.loads((h/'adjusted-import-data.json').read_text(encoding='utf-8'))
xml=json.loads((h/'zip-237-xml-review-20261009.json').read_text(encoding='utf-8'))
selected=[r for r in xml['records_237'] if r['line_goods_vs_header_mismatch']];assert len(selected)==195
w=openpyxl.load_workbook(p,read_only=True,data_only=True)
main={tuple(v[:2]):(n,v) for n,v in enumerate(w['Dieu chinh'].iter_rows(min_row=3,max_col=55,values_only=True),3)}
imp=w['Import dieu chinh'];start=10+sum(1 for v in imp.iter_rows(min_row=10,max_col=1,values_only=True))
seq=max(v[0] or 0 for v in imp.iter_rows(min_row=10,max_col=1,values_only=True));assert start==40787 and seq==1600
existing={tuple(v) for v in imp.iter_rows(min_row=10,min_col=10,max_col=11,values_only=True)}
rows=[];changes=[];meta=[];chains=[]
clean=lambda v:None if v is None or str(v).strip()=='' else str(v).strip()
date=lambda s:datetime.fromisoformat(s[:10]).strftime('%d/%m/%Y')
def rate(v):return float(str(v).strip().rstrip('%'))
for record in sorted(selected,key=lambda r:main[tuple(r['related'])][0]):
 k=tuple(record['related']);adjust=tuple(record['key']);n,v=main[k]
 assert k not in existing and all(vv==0 for vv in v[24:27])
 assert tuple(v[6:8])==adjust and not v[12] and not v[18]
 assert record['header_matches_report'] and record['reference_matches_report'] and record['date_matches_current_workbook']
 key='/'.join(k);old='/'.join(adjust);hd=source['headers'][key];lines=source['detail'][key];common=lines[0]
 original=[D(str(x)) for x in hd['money']];reverse=[-D(str(x)) for x in record['money']]
 assert reverse==original and original[0]+original[1]==original[2]
 assert [sum(D(str(l[c] or 0)) for l in lines) for c in ['goods','tax','pay']]==original
 assert all(l['qty'] not in [None,0] and l['price'] is not None and l['goods'] is not None and l['name'] and l['method'] and l['discount'] in [None,0] for l in lines)
 rates={rate(l['rate']) for l in lines if l['rate'] not in [None,'']};assert len(rates)==1
 tr=next(iter(rates));assert tr==8
 desc=f'Điều chỉnh tăng thành tiền để đảo khoản giảm trên HĐ {old}; tham chiếu HĐ gốc {key}'
 seq+=1
 rows.append([seq,'2026-10-09',clean(common['buyer']),clean(common['address']),clean(common['taxid']),clean(common['buyer']),clean(common['email']),clean(common['method']),None,*k,date(str(v[2])),None,desc,desc,None,None,None,float(reverse[0]),float(reverse[0]),tr,float(reverse[1]),float(reverse[2]),None])
 meta.append({'root':k,'source':adjust,'kind':'reverse','seq':seq,'first':True});reverse_seq=seq;seq+=1
 reason=f'Điều chỉnh giảm toàn bộ chi tiết HĐ gốc {key}; làm lại sau khi đảo khoản trên HĐ {old}'
 for i,l in enumerate(lines):
  first=i==0
  rows.append([seq,'2026-10-09',clean(l['buyer']),clean(l['address']),clean(l['taxid']),clean(l['buyer']),clean(l['email']),clean(l['method']),None,*k,date(str(v[2])),None,reason,clean(l['name']),clean(l['unit']),-l['qty'],l['price'],-l['goods'],-float(original[0]) if first else None,tr if first else None,-float(original[1]) if first else None,-float(original[2]) if first else None,1 if l['goods']==0 else None])
  meta.append({'root':k,'source':k,'kind':'reduce','seq':seq,'first':first})
 note=f'CHỐT LÀM LẠI 09/10/2026: XML {old} có tiền hàng chi tiết 0 nhưng tổng tiền hàng khác 0. F4 đảo khoản giảm cũ theo tổng XML, F5 giảm đủ chi tiết F0 theo mẫu 00000003/00000508. Đã thêm import STT {reverse_seq}–{seq}. Sau dự kiến tiền hàng/thuế/thanh toán vẫn 0. '
 note+='Ghi chú trước khi chốt: '+str(v[27] or '')
 if k==('1C25TUV','00001314'):note=note.replace('Chưa kiểmXML.','XML đã kiểm, xác nhận ngày 18/08/2025.')
 changes.append({'row':n,'values':{'AB':note,'AC':f'Đảo khoản {old}; tham chiếu gốc {key} (đã chốt theo XML)','AE':'2026-10-09','AI':f'Giảm toàn bộ chi tiết gốc {key} (đã chốt làm lại)','AK':'2026-10-09'}})
 chains.append({'root':k,'adjustment':adjust,'row':n,'reverse_seq':reverse_seq,'reduce_seq':seq,'original_money':[float(x) for x in original],'line_count':len(lines),'sources':list(dict.fromkeys(l['source'] for l in lines)),'xml_sha256':record['sha256']})
assert seq==1990 and len(chains)==195
w.close();shutil.copy2(p,h/'before-195-xml.xlsx')
payload={'base_sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'start_row':start,'last_row':start+len(rows)-1,'rows':rows,'meta':meta,'changes':changes,'chains':chains,'invoices':1990,'roots':986,'total_lines':40777+len(rows),'added_lines':len(rows),'promo_added':sum(r[23]==1 for r in rows)}
(h/'195-xml-data.json').write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:payload[k] for k in ['start_row','last_row','invoices','roots','total_lines','added_lines','promo_added']}))
