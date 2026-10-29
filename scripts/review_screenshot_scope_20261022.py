"""Tái kiểm snapshot chỉ đọc; lưu kết quả phân tích riêng, không sửa nguồn/import."""
from pathlib import Path
from collections import Counter, defaultdict
from decimal import Decimal
import json, hashlib, sys

ROOT = Path(__file__).resolve().parents[1]
sys.stdout.reconfigure(encoding='utf-8')
base = ROOT / 'scripts/data/link-audit-20261021'
saved = json.loads((base / 'summary.json').read_text(encoding='utf-8'))
# Chạy lại phần tính toán trước mọi thao tác ghi của script gốc.
script = ROOT / 'scripts/audit_order_links_20261021.py'
source = script.read_text(encoding='utf-8').split("payload=dict(summary=summary,results=results,amis=A,nb=N,ktReview=review)")[0]
ns = {'__file__': str(script), '__name__': '__read_only_review__'}
exec(compile(source, str(script), 'exec'), ns)
current = ns['summary']
assert current == saved, 'Kết quả tái tính khác báo cáo đã lưu'
assert current['sourceHashes'] == saved['sourceHashes']
rows, amis = ns['results'], ns['A']
old = [r for r in rows if r['oldReceive'] and r['category'] == 'corroborated']
assert len(old) == 116
assert all(r['oldAmis'] and not r['duplicateOrder'] and r['status'] == 'Hoá đơn mới' for r in old)

chain_data = json.loads((ROOT/'scripts/data/h2-chain-readiness-20261022/workbook-input.json').read_text(encoding='utf-8'))
root_of = {}
invoice_data = {}
for c in chain_data['chains']:
    for v in c['invoices']:
        k = tuple(v['key'])
        assert k not in root_of
        root_of[k] = tuple(c['root'])
        invoice_data[k] = v
assert len(root_of) == 37758 and len(chain_data['chains']) == 37087
groups = defaultdict(list)
for r in rows:
    if r['code']:
        groups[r['code']].append(r)
duplicates = [g for g in groups.values() if len(g) > 1]
dup_class = Counter()
for g in duplicates:
    keys = [(r['series'],r['invoice']) for r in g]
    if all(k in root_of for k in keys):
        dup_class['one_known_chain' if len({root_of[k] for k in keys}) == 1 else 'multiple_known_roots'] += 1
    else:
        dup_class['not_all_in_h2_chain_scope'] += 1

from openpyxl import load_workbook
books = []
existing_source_keys = set()
for file in ['HOA_DON_CHUA_DIEU_CHINH_IMPORT_2_SHEET.xlsx','HOA_DON_DA_DIEU_CHINH_IMPORT_2_SHEET.xlsx']:
    p = ROOT/'outputs'/file
    w = load_workbook(p, read_only=True, data_only=True)
    sheet = w['Import dieu chinh']
    # Đếm số hóa đơn dự kiến theo cột được nhận diện từ tiêu đề thực tế.
    preview = list(sheet.iter_rows(min_row=1,max_row=8,values_only=True))
    import_ids = {r[0] for r in sheet.iter_rows(min_row=10,values_only=True) if isinstance(r[0],(int,float))}
    context = w.worksheets[1]
    for r in context.iter_rows(min_row=3,values_only=True):
        for offset in ([0] if context.title=='F0' else [0,6,12,18]):
            if offset+1 < len(r) and isinstance(r[offset],str) and r[offset].startswith('1C') and r[offset+1]:
                existing_source_keys.add((r[offset],str(r[offset+1])))
    books.append({'file': str(p.relative_to(ROOT)), 'sha256':hashlib.sha256(p.read_bytes()).hexdigest(), 'sheets':w.sheetnames, 'rows':sheet.max_row,'invoice_count':len(import_ids),'header_preview':[list(r[:8]) for r in preview]})
    w.close()
google_keys = {(r['series'],r['invoice']) for r in rows}
h2_keys = {k for k,v in invoice_data.items() if '2025-07-01' <= v['date'] <= '2025-12-31'}
# H2 bao gồm các hóa đơn con của F0 trước tháng 7: lấy bảng tổng đầy đủ ở bước sau khi cần.
old_keys = {(r['series'],r['invoice']) for r in old}
old_chain_states = Counter()
for k in old_keys:
    v = invoice_data.get(k)
    old_chain_states[v['status'] if v else 'outside_h2_chain_scope'] += 1

result = {
 'verified_snapshot_date':'2026-10-21', 'rerun_matches_saved_summary':True,
 'source_hashes_verified':len(current['sourceHashes']),
 'source_counts':current['sourceCounts'], 'categories':current['categories'],
 'amis_date_range':[min(a['date'] for a in amis if a['date']), max(a['date'] for a in amis if a['date'])],
 'amis_years':dict(Counter(a['date'][:4] or 'blank' for a in amis)),
 'old_candidates':len(old), 'old_candidate_total':str(sum((r['total'] for r in old),Decimal(0))),
 'old_candidate_amis_date_range':[min(r['amisDates'][0] for r in old),max(r['amisDates'][0] for r in old)],
 'old_candidate_receipt_range':[min(r['received'] for r in old),max(r['received'] for r in old)],
 'old_candidate_invoice_range':[min(r['date'] for r in old),max(r['date'] for r in old)],
 'old_candidate_masked_names':sum(r['maskedBuyer'] for r in old),
 'old_candidate_current_excel_statuses':dict(old_chain_states),
 'old_candidate_multiple_invoice_chains':sum(sum(rt == root_of.get(k) for rt in root_of.values())>1 for k in old_keys if k in root_of),
 'old_candidates_already_in_existing_import_source':len(old_keys & existing_source_keys),
 'duplicate_codes':len(duplicates), 'duplicate_chain_classification':dict(dup_class),
 'duplicate_groups_at_least_two_new_invoices':sum(sum(r['status']=='Hoá đơn mới' for r in g)>=2 for g in duplicates),
 'new_invoice_direct_money_conflicts':current['directMismatchNew'],
 'books':books,
 'candidate_details':[{'key':[r['series'],r['invoice']],'row':r['row'],'nb_rows':r['nbMatchingRows'],'amis_rows':r['amisMatchingRows'],'amis_documents':r['amisDocs'],'received':r['received'],'amis_dates':r['amisDates'],'invoice_date':r['date'],'total':str(r['total'])} for r in old],
}
out = ROOT/'outputs/doi-chieu-van-de-anh-20261022'
out.mkdir(parents=True,exist_ok=True)
(out/'ket_qua_tai_kiem.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
print(json.dumps({k:v for k,v in result.items() if k not in ('candidate_details','books')},ensure_ascii=False,indent=2))
print(json.dumps(books,ensure_ascii=False,default=str))
