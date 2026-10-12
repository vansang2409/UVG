"""Build source-preserving, confirmed action proposals for the review workbook."""
from collections import Counter
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
import json

import openpyxl

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
read = lambda n: json.loads((HERE / n).read_text(encoding='utf-8'))
data = read('review-data.json')
audit = read('confirmed-decisions-audit.json')
manifest = read('review-manifest.json')
decisions = read('accountant-decisions.json')
assert decisions['remaining_points'] == []
pending = {tuple(c['root']): c for c in data['audit']['pending']}
review = {tuple(c['root']): c for c in data['audit']['review_11']}
xml = {tuple(x['key']): x for x in data['xmls']}
raw = {tuple(r['key']): r for r in manifest['related']}
extra = {tuple(x['key']): x for x in data['extra']}
actions = []
case_plans = []
kept_cases = {'HH02', 'HH03', 'HH05', 'CC01', 'CC02', 'CC03', 'CC04', 'CC05', 'CC06'}

def number(value):
    if value is None:
        return None
    n = Decimal(str(value))
    return int(n) if n == n.to_integral_value() else float(n)

def key_text(key):
    return '/'.join(key)

def add(case, target, action, method, money=None, detail=None):
    target = tuple(target)
    src = raw[target]
    amounts = xml[target]['money'] if action == 'Đảo điều chỉnh cũ' else src['money']
    source_kind = 'XML đã phát hành' if action == 'Đảo điều chỉnh cũ' else 'Bảng tổng xuất lại'
    values = [(-number(v) if v is not None else None) for v in amounts] if money is None else money
    assert all(v is not None for v in values)
    actions.append({'id': f"PA{len(actions)+1:02d}", 'case': case['id'], 'root': case['root'],
                    'target': target, 'date': src['date'], 'action': action, 'method': method,
                    'money': values, 'source_money': [number(v) for v in amounts], 'source_kind': source_kind,
                    'detail': detail or method, 'source_xml': xml.get(target, {}).get('file'),
                    'difference': values[0] + values[1] - values[2]})

for case in audit['review_cases']:
    assert not case['protected']
    root = tuple(case['root'])
    ident, group = case['id'], case['group']
    if ident in kept_cases:
        case_plans.append({'id': ident, 'sheet': group, 'row': case['row'], 'mode': 'Giữ nguyên',
                           'decision': 'Đã xác nhận cùng sản phẩm thực tế. Giữ nguyên cặp hóa đơn, không lập thêm điều chỉnh trong đợt này.'})
        continue
    if group == 'Lech tien':
        text = 'Đảo từng chỉ tiêu đúng số trên XML đã phát hành của mọi điều chỉnh cũ trong chuỗi, giữ nguyên độ lớn; đồng thời chuẩn bị giảm đủ chi tiết hóa đơn gốc.'
        for inv in pending[root]['invoices']:
            if tuple(inv['key']) != root:
                add(case, inv['key'], 'Đảo điều chỉnh cũ', 'Đảo từng số tiền đã phát hành; không tính lại cho cân bằng.')
        add(case, root, 'Giảm hóa đơn gốc', 'Giảm đủ dòng gốc: đảo số lượng/thành tiền, giữ đơn giá nguồn.')
    elif group == 'So luong can kiem':
        if root in pending:
            for inv in pending[root]['invoices']:
                if tuple(inv['key']) != root:
                    add(case, inv['key'], 'Đảo điều chỉnh cũ', 'Đảo số tiền đã phát hành của điều chỉnh cũ.')
        if ident == 'SL01':
            text = 'Gốc 1C24TUV/00000555: dòng 6 giữ số lượng và đơn giá trống, chỉ đảo thành tiền -46.296. Các dòng còn lại giảm theo chi tiết gốc. Đảo riêng điều chỉnh 00002632.'
        elif ident == 'SL02':
            text = 'Bỏ dòng 45 KA600 FT không có giá trị; giảm các dòng còn lại của gốc 00000113. Đảo riêng điều chỉnh 00000635 theo số đã phát hành.'
        elif ident in {'SL03','SL04','SL05','SL06'}:
            text = 'Giữ dòng chú thích hàng biếu/tặng không thu tiền, tính chất ghi chú/diễn giải; giữ số lượng, đơn giá, tiền dòng trống như bảng xuất mới. Giảm các dòng hàng có số liệu.'
        else:
            product = 'KM002' if ident == 'SL07' else 'KM003'
            text = f'Bỏ đúng dòng 2 {product} không có giá trị; giảm các dòng hàng còn lại của {key_text(root)}.'
        add(case, root, 'Giảm hóa đơn gốc', text)
    elif group == 'Thanh toan trong':
        child = tuple(case['focus'])
        text = f'Gốc {root[1]} chỉ giảm số lượng từng dòng theo phản hồi KT; đơn giá để trống, không thêm giá trị tiền. Đảo riêng tiền hàng/thuế/thanh toán XML của {child[1]}. Ô thanh toán gốc trống vẫn giữ nguyên ở nguồn.'
        add(case, child, 'Đảo điều chỉnh cũ', 'Đảo riêng tiền hàng, thuế, thanh toán đã phát hành.')
        add(case, root, 'Giảm số lượng gốc', 'Chỉ giảm số lượng; không có đơn giá/giá trị tiền.', money=[0,0,0], detail=case['reply']['Số liệu/nội dung đúng (KT điền)'])
    elif group == 'Hang va mo ta':
        child = tuple(case['focus'])
        text = f'Giảm đúng sản phẩm trên gốc {root[1]}; tăng lại sản phẩm đã giảm sai trên điều chỉnh {child[1]}, theo từng mã hàng và số lượng đã chốt.'
        add(case, child, 'Đảo điều chỉnh cũ', 'Đảo số lượng của hàng đã giảm sai, giữ đơn giá nguồn; thành tiền/thuế/thanh toán tăng trở lại.')
        add(case, root, 'Giảm hóa đơn gốc', 'Giảm sản phẩm đúng của hóa đơn gốc, số lượng -1 theo KT, giữ đơn giá nguồn.')
    else:
        child = tuple(case['focus'])
        if ident == 'DC01':
            text = 'Giảm gốc 00001061: KM300, SL -1, đơn giá499.074, thành tiền-499.074. Đảo00012835: tiền hàng+499.074, thuế+39.926, thanh toán+539.000; phần đảo là khoản tiền.'
            method = 'Đảo khoản tiền: thành tiền, thuế, thanh toán đổi dấu; số lượng và đơn giá không tự thêm.'
        else:
            price = number(review[root]['original_detail'][0]['price'])
            text = f'Cặp đúng {root[1]}/{child[1]}. Giảm gốc: SL -1, giá +{price:,}, thành tiền -{price:,}. Đảo điều chỉnh cũ: SL -1, giữ giá -{price:,}, thành tiền +{price:,}; thuế/thanh toán đảo theo XML.'
            method = f'SL -1, giữ đơn giá âm {number(xml[child]["lines"][0]["price"]):,}; thành tiền dương, đảo thuế và thanh toán theo XML.'
            line = xml[child]['lines'][0]
            assert -number(line['qty']) * number(line['price']) == -number(line['goods'])
        add(case, child, 'Đảo điều chỉnh cũ', method)
        add(case, root, 'Giảm hóa đơn gốc', 'SL -1, giữ đơn giá dương; thành tiền/thuế/thanh toán giảm theo gốc.')
    case_plans.append({'id': ident, 'sheet': group, 'row': case['row'], 'mode': 'Điều chỉnh', 'decision': text})

known_plans = []
for row, case in enumerate(manifest['resolved'], 8):
    if case['id'] in kept_cases:
        decision = 'Giữ cả cặp và mọi hóa đơn liên quan ngoài đợt xử lý theo xác nhận điểm 6. Không lập điều chỉnh mới.'
        mode = 'Giữ nguyên'
    else:
        decision = 'Giảm số lượng hàng hóa gốc 00009072 theo phản hồi KT; giữ nguyên điều chỉnh tên hàng00009186, không đảo hóa đơn này.'
        mode = 'Điều chỉnh'
        add(case, case['root'], 'Giảm hóa đơn gốc', decision)
    known_plans.append({'id':case['id'], 'row':row, 'decision':decision, 'mode':mode})

assert len(actions) == 48
assert len({tuple(a['root']) for a in actions}) == 27
assert len({tuple(a['target']) for a in actions}) == 48
assert Counter(c['mode'] for c in case_plans) == {'Điều chỉnh':26, 'Giữ nguyên':3}
assert Counter(c['mode'] for c in known_plans) == {'Giữ nguyên':6, 'Điều chỉnh':1}
assert sum(a['difference'] != 0 for a in actions) == 8
assert sum(a['action'] == 'Đảo điều chỉnh cũ' for a in actions) == 21
targets = {tuple(a['target']) for a in actions}
assert len(set(raw) - targets) == 19
for a in actions:
    assert not any(tuple(a['target']) in set(map(tuple, h['invoices'])) for h in audit['held_chains'])
    if a['action'] == 'Đảo điều chỉnh cũ':
        assert [x+y for x,y in zip(a['source_money'], a['money'])] == [0,0,0]

source_book = openpyxl.load_workbook(audit['preserved_path'], read_only=False, data_only=False)
source_views = {s.title: {'freeze_panes': s.freeze_panes, 'show_grid_lines': s.sheet_view.showGridLines} for s in source_book}
source_book.close()
result = {'source': audit['preserved_path'], 'source_sha256': audit['source_sha256'],
          'source_views': source_views,
          'case_plans':case_plans, 'known_plans':known_plans, 'actions':actions,
          'held_chains':audit['held_chains'], 'exclusion_summary':audit['summary'],
          'invoice_dates':audit['invoice_dates'],
          'summary':{'active_cases_confirmed':29, 'reference_cases_confirmed':7, 'new_actions':48,
                     'action_roots':27, 'reversals':21, 'root_reductions':27, 'source_invoices_kept':19,
                     'reversal_arithmetic_differences_preserved':8, 'import_workbooks_changed':False},
          'workbook_hashes':manifest['workbookHashes']}
(HERE / 'confirmed-plan-data.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result['summary'],ensure_ascii=False))
