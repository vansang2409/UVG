"""Chuẩn bị bảng làm việc và kiểm chứng đầy đủ chứng từ, không sửa nguồn."""
import json,hashlib
from pathlib import Path
from collections import Counter,defaultdict
from datetime import datetime
ROOT=Path(__file__).resolve().parents[1];P=ROOT/'outputs/phan-loai-2025-20261105'
d=json.loads((P/'classification.json').read_text(encoding='utf-8'))
source=json.loads((P/'source/plan25AMIS.json').read_text(encoding='utf-8'))
groups=['Giữ HĐ','Sửa liên kết','Xử lý HĐ sai','Xuất bổ sung','Cần KT xác minh']
bydoc={x['doc']:x for x in d['docs']}
assert len(bydoc)==len(source)-1==65494
assert set(bydoc)=={x[2] for x in source[1:]}
assert sum(x['money'] or 0 for x in d['docs'])==sum(x[7] for x in source[1:])
assert sum(x['vat'] or 0 for x in d['docs'])==sum(x[6] for x in source[1:])
checks={}
for rn,s in enumerate(source[1:],2):
    x=bydoc[s[2]]
    assert (x['row'],x['serial'],x['money'],x['vat'],x['order'])==(rn,s[0],s[7],s[6],str(s[8]).strip() if s[8] is not None else '')
    if x['category']=='Giữ HĐ':
        assert not x['protected'] and all(v is not None for v in x['balance'])
        assert abs(x['group_money']-x['balance'][2])<=2 and abs(x['group_vat']-x['balance'][1])<=2
        assert abs(x['group_money']-x['group_vat']-x['balance'][0])<=2
    if x['category']=='Xuất bổ sung':assert not x['protected'] and all(v is not None and abs(v)<=2 for v in x['balance']) and x['group_money']>0
assert bydoc[source[12720][2]]['row']==12721
assert bydoc[source[12720][2]]['category']=='Sửa liên kết'
assert bydoc[source[357][2]]['balance']==[0,0,0]
assert bydoc[source[357][2]]['category']=='Xuất bổ sung'
assert bydoc[source[36229][2]]['category']=='Sửa liên kết'
assert len({r['root'] for r in d['invoice_review'] if r['decision']})==25
checks={'all_65494_documents_preserved':True,'money_and_vat_preserved':True,'kept_chains_match_three_amounts':True,'supplement_candidates_have_verified_zero_balance':True,'protected_roots_not_exported':True,'known_2026_zero_chain_not_kept':True,'known_wrong_link_not_kept':True,'all_25_accountant_pairs_visible':True}
headers=['Số chứng từ','Ngày AMIS','Mã đơn hàng','Thanh toán AMIS (đ)','VAT AMIS (đ)','HĐ gốc đối soát','Lệch nhóm (đ)','Vấn đề','Bước tiếp theo','Kết luận KT','Ghi chú KT','Link AMIS','Số HĐ đang gắn','Ký hiệu đang gắn','HĐ còn hiệu lực','Ngày HĐ gốc','Số CT cùng chuỗi','Thanh toán AMIS nhóm (đ)','VAT AMIS nhóm (đ)','Thanh toán chuỗi (đ)','VAT chuỗi (đ)','Lệch VAT nhóm (đ)','Căn cứ ghép','Ghi chú AMIS gốc','HĐ đề xuất / ứng viên','Chi tiết nguồn chuỗi']
def plain(x):
    if x is None:return None
    s=str(x)
    return "'"+s if s.startswith('=') else s
def dateserial(s):return (datetime.fromisoformat(s)-datetime(1899,12,30)).days if s else None
tables=[]
suspect_roots={x['root'] for x in d['docs'] if x['category']=='Sửa liên kết' and x['root']}
for cat in groups:
    data=[]
    for x in sorted((r for r in d['docs'] if r['category']==cat),key=lambda r:(r['month'],r['doc'])):
        bal=x['balance'];known=bal[2] is not None and x['group_money'] is not None;knownvat=bal[1] is not None and x['group_vat'] is not None
        data.append([x['doc'],x['serial'],x['order'],x['money'],x['vat'],x['root'],bal[2]-x['group_money'] if known else None,x['reason'],x['action'],None,None,d['source_url']+f"#gid=1264954379&range=A{x['row']}:M{x['row']}",plain(x['old_invoice']),x['old_symbol'],'; '.join(x['active']),dateserial(x['root_date']),x['group_count'],x['group_money'],x['group_vat'],bal[2],bal[1],bal[1]-x['group_vat'] if knownvat else None,'; '.join(x['evidence']),plain(x['note']),x['proposal'] or '; '.join(x['hints']),'Chuỗi: '+'; '.join(x['chain_keys']) if x['chain_keys'] else 'Chưa có chuỗi xác minh'])
    compact_headers=headers[:12]+['HĐ AMIS đang gắn','HĐ đề xuất / ứng viên','HĐ còn hiệu lực','Thanh toán AMIS nhóm (đ)','Thanh toán chuỗi (đ)','VAT chuỗi (đ)','Ghi chú AMIS gốc']
    data=[r[:12]+[(str(r[13])+'/'+str(r[12])) if r[13] else r[12],r[24],r[14],r[17],r[19],r[20],r[23]] for r in data]
    if cat=='Xử lý HĐ sai':
        for r in data:
            if r[5] in suspect_roots:
                r[7]='Nhóm có CT nghi ghép nhầm; chưa kết luận HĐ sai'
                r[8]='Sửa liên kết nhóm trước, rồi so lại tiền và VAT HĐ'
    tables.append({'name':'Rà soát HĐ' if cat=='Xử lý HĐ sai' else cat,'headers':compact_headers,'rows':data})
invoice_headers=['HĐ cần kiểm tra','Ngày HĐ','Mã đơn nguồn','Người mua','Thanh toán chuỗi (đ)','VAT chuỗi (đ)','Trước thuế chuỗi (đ)','Vấn đề','Kết luận KT hiện có','HĐ KT chọn giữ','Ghi nhận tiếp theo','Ghi chú tiếp theo','HĐ còn hiệu lực','Trạng thái HĐ gốc','Thiếu sót nguồn','Nguồn bảng kê']
rs=[]
for r in sorted(d['invoice_review'],key=lambda x:(not bool(x['decision']),x['root'])):
    b=r['balance'];src=r['source'] or {}
    rs.append([r['root'],dateserial(r['date']),r['order'],r['buyer'],b[2],b[1],b[0],r['reason'],r['decision'],r['pair_keep'],None,None,'; '.join(r['active']),r['state'],'; '.join(r['issues']),f"{src.get('file','')} / dòng {src.get('row','')}"])
tables.append({'name':'HĐ cần kiểm riêng','headers':invoice_headers,'rows':rs})
out={'as_of':d['as_of'],'source_url':d['source_url'],'source_cutoff':d['source_cutoff'],'summary':d['summary'],'monthly':d['monthly'],'tables':tables,'verification':checks,'source_total':sum(x[7] for x in source[1:]),'source_vat':sum(x[6] for x in source[1:]),'review_pair_count':25,'limits':['Giữ HĐ là đề xuất theo mã đơn, tiền và VAT; chưa đối chiếu chi tiết hàng AMIS2025.','Xuất bổ sung là ứng viên có chuỗi về0; chưa phải danh sách được phép phát hành.','357 chứng từ chưa tìm được HĐ giữ tại Cần KT xác minh; chưa kết luận chưa xuất.','Tiền nhóm/chuỗi lặp trên các CT cùng nhóm: không cộng các cột này. Tổng hợp chỉ cộng tiền từng CT.','Bảng kê nguồn đến29/10/2026, bổ sung phát hành01/11/2026; cần làm mới trước khi phát hành.','Không sửa AMIS/Ktoan, không xử lý chuỗi2024 đã chốt; giữ nguyên ghi chú nguồn.']}
(P/'workbook-data-compact.json').write_text(json.dumps(out,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
manifest={str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in (P/'source').glob('*.json')}
(P/'verification.json').write_text(json.dumps({'checks':checks,'docs':len(bydoc),'total_money':out['source_total'],'total_vat':out['source_vat'],'manifest':manifest,'workbook_data_sha256':hashlib.sha256((P/'workbook-data-compact.json').read_bytes()).hexdigest()},ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'checks':checks,'summary':d['summary'],'invoice_review':len(rs),'max_text_length':max(len(v) for t in tables for row in t['rows'] for v in row if isinstance(v,str)),'data_bytes':(P/'workbook-data.json').stat().st_size},ensure_ascii=False))
