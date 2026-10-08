import json,sys,hashlib
from pathlib import Path
from collections import Counter
from decimal import Decimal
import openpyxl
r=Path(r'D:\UVG');h=r/'scripts/data/check-review';sys.path.insert(0,str(r/'scripts'))
from prepare_invoice_reports import rows
p=r/'outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx';w=openpyxl.load_workbook(p,data_only=True);s=w['Dieu chinh']
roots=[{'row':n,'key':(v[0],v[1]),'old_note':v[27],'original_money':list(v[3:6])} for n,v in enumerate(s.iter_rows(min_row=3,values_only=True),3) if all(a==0 for a in v[24:27])]
assert len(roots)==237
wanted={x['key'] for x in roots}; detail={k:[] for k in wanted};hashes={}
for name in ['Bang_ke_chi_tiet_HD_da_su_dung_2024.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls']:
 src=r/'data/Du lieu Meinvoice'/name;hashes[str(src.relative_to(r))]=hashlib.sha256(src.read_bytes()).hexdigest();mtt='_MTT' in name
 for n,v in rows(src):
  key=(v.get('F' if mtt else 'D'),v.get('B'))
  if key not in wanted:continue
  def num(c):
   z=v.get(c);return None if z in (None,'') else float(Decimal(z))
  detail[key].append({'row':n,'source':name,'name':v.get('S' if mtt else 'L'),'qty':num('U' if mtt else 'N'),'price':num('V' if mtt else 'O'),'goods':num('W' if mtt else 'P'),'discount':num('Y' if mtt else 'R'),'tax':num('AA' if mtt else 'T'),'payment':num('AB' if mtt else 'U'),'promo_flag':v.get('AD' if mtt else 'W')})
counts=Counter(); flags=Counter(); linecount=0
for x in roots:
 d=detail[x['key']];zero=sum(a['goods']==0 for a in d);sale=sum(a['goods'] is not None and a['goods']!=0 for a in d);missing=sum(a['goods'] is None for a in d);flagged=[a for a in d if a['promo_flag'] not in (None,'','0',0,'False','false')]
 flags.update(str(a['promo_flag']) for a in d);linecount+=len(d)
 category='Không có chi tiết' if not d else 'Thiếu thành tiền chi tiết' if missing else 'Toàn bộ dòng gốc bằng 0' if zero==len(d) else 'Có bán và dòng bằng 0' if zero else 'Toàn bộ dòng gốc khác 0'
 counts[category]+=1
 text=f"Rà 237 chuỗi: F0–F3 đã cộng dồn 0. Gốc có {len(d)} dòng chi tiết: {sale} dòng thành tiền khác 0, {zero} dòng bằng 0, {missing} dòng thiếu thành tiền. "
 if not d:text+='Chưa tìm được chi tiết gốc; chờ đối chiếu. '
 elif missing:text+='Chưa kết luận đủ dòng bán/khuyến mãi do thiếu số. '
 elif sale==0:text+='Toàn bộ thành tiền gốc = 0: theo quy tắc đã chốt, nhập 1 cột Hàng khuyến mãi cho từng dòng. '
 elif zero:text+='Có chi tiết bán và dòng 0; chỉ dòng thành tiền gốc = 0 mới nhập 1 cột Hàng khuyến mãi. '
 else:text+='Có chi tiết bán; không đánh dấu khuyến mãi theo quy tắc thành tiền = 0. '
 text+=f"Cờ KM nguyên bản: {len(flagged)}/{len(d)} dòng có giá trị. "
 refs=[]
 for name in dict.fromkeys(a['source'] for a in d):
  ns=[a['row'] for a in d if a['source']==name];refs.append(name+' | dòng '+(str(ns[0]) if len(ns)==1 else str(min(ns))+'–'+str(max(ns))))
 text+='Nguồn: '+'; '.join(refs)+'.'
 x.update(category=category,detail=d,note=(str(x['old_note'])+'; ' if x['old_note'] else '')+text)
 x['key']=list(x['key'])
assert all(x['detail'] for x in roots)
(h/'zero-detail-review.json').write_text(json.dumps({'roots':roots,'counts':dict(counts),'linecount':linecount,'flags':dict(flags),'hashes':hashes},ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps({'roots':len(roots),'linecount':linecount,'counts':dict(counts),'source_flags':dict(flags)},ensure_ascii=False))
print('SAMPLES',[(x['key'],x['category'],len(x['detail'])) for x in roots[:5]])
