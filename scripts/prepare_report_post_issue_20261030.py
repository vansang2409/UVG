"""Prepare a scoped update of the existing 2024 revenue report from issued data."""
from pathlib import Path
from collections import defaultdict, Counter
from decimal import Decimal
import hashlib
import json
import sys

from prepare_invoice_reports import rows, invoice_key, invoice_day, amount, native
from filter_invoice_exceptions import normal

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / 'scripts/data/report-post-issue-20261030'
BASE = ROOT / 'data/Meinv_SAU_PHAT_HANH_20261029'
SID = 700000010


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    snapshot = json.loads((WORK / 'before-cells.json').read_text(encoding='utf-8'))
    live = next(s for s in snapshot['sheets'] if s['properties']['sheetId'] == SID)['data'][0]['rows']
    prior = json.loads((ROOT / 'scripts/data/revenue-chain-audit-20261028/prepared.json').read_text(encoding='utf-8'))
    audit = json.loads((ROOT / 'scripts/data/post-issue-20261029/audit.json').read_text(encoding='utf-8'))
    baseline = json.loads((ROOT / 'scripts/data/revenue-chain-audit-20261028/review-sheet.json').read_text(encoding='utf-8'))
    policy = json.loads((ROOT / 'scripts/data/post-issue-20261029/confirmed-blank-payment-zero.json').read_text(encoding='utf-8'))
    confirmed_keys = {tuple(k) for k in policy['invoice_keys']}
    assert len(live[14:]) == len(baseline['rows']) == 778
    assert [r[0] for r in live[14:]] == [r[0] for r in baseline['rows']]

    records = {}
    edges = {}
    children = defaultdict(set)
    parents = defaultdict(set)
    replaced = set()
    hashes = {}
    for p in sorted(BASE.rglob('*.xls')):
        rel = p.relative_to(ROOT).as_posix()
        hashes[rel] = hashlib.sha256(p.read_bytes()).hexdigest()
        kind = 'replacement' if p.parent.name == 'THAY_THE' else 'adjustment' if p.parent.name == 'DIEU_CHINH' else 'total'
        for rn, r in rows(p):
            if rn < (6 if kind == 'total' else 7) or amount(r.get('A')) is None:
                continue
            if kind == 'total':
                key = invoice_key(r.get('B'), r.get('C'))
                assert key not in records
                records[key] = {'key': key, 'date': invoice_day(r.get('D')), 'money': [amount(r.get(c)) for c in ['P', 'Q', 'R']],
                                'status': r.get('U') or '', 'cqt': r.get('S') or '', 'source': rel, 'row': rn}
            else:
                old, new = invoice_key(r.get('B'), r.get('D')), invoice_key(r.get('M'), r.get('O'))
                assert old and new
                children[old].add(new)
                parents[new].add(old)
                edges[(old, new, kind)] = {'file': rel, 'row': rn, 'legacy': False}
                if kind == 'replacement':
                    replaced.add(old)
    # Retain verified historical links omitted by the current export, without
    # changing the dates or amounts of any current invoice.
    recovered = []
    for e in prior['edges']:
        old, new, kind = tuple(e['parent']), tuple(e['child']), e['kind']
        if (old, new, kind) not in edges:
            edges[(old, new, kind)] = {'legacy': True, 'evidence': e['evidence']}
            recovered.append({'old': old, 'new': new, 'kind': kind})
            children[old].add(new)
            parents[new].add(old)
            if kind == 'replacement':
                replaced.add(old)
    for k in confirmed_keys:
        if k in records:
            assert records[k]['money'] == [0, 0, None]
            records[k]['money'] = [Decimal(0)] * 3
            records[k]['payment_basis'] = 'user_confirmed_blank_zero'
    processed = {tuple(r['root']): r for g in ['adjusted', 'untreated'] for r in audit['groups'][g] if r['root'][0] == '1C24TUV'}
    replacement_roots = {('1C24TUV', f'{root:08}'): ('1C25TUV', f'{last:08}') for root, last in [(80,2601),(87,2610),(116,2602),(117,2603),(118,2604),(184,2617)]}
    assert len(processed) == 654 and sum(r['planned_count'] for r in processed.values()) == 1253

    def reach(root):
        found = set()
        todo = [root]
        while todo:
            k = todo.pop()
            if k in found:
                continue
            found.add(k)
            todo.extend(children[k])
        return found

    updates = []
    evidence = []
    used_keys = set()
    for rn, row in enumerate(live[14:], 15):
        root = invoice_key(*row[0].split('/'))
        nodes = reach(root)
        assert set(nodes) <= records.keys(), (root, 'missing current issued record', set(nodes)-records.keys())
        active = [k for k in sorted(nodes) if k not in replaced and 'bi thay the' not in normal(records[k]['status']) and 'huy' not in normal(records[k]['status'])]
        assert active or all('huy' in normal(records[k]['status']) for k in nodes), (root, 'no contributing invoice without a cancellation basis')
        current_goods = sum(records[k]['money'][0] for k in active)
        old_value = row[11]['v']
        assert current_goods == old_value, (root, current_goods, old_value)
        kind = 'adjustment' if root in processed else 'replacement' if root in replacement_roots else 'retained'
        new = [k for k in active if records[k]['date'].year == 2026]
        if kind == 'adjustment':
            assert processed[root]['balance'] == [0, 0, 0]
            assert set(map(tuple, processed[root]['all_keys'])) == nodes
            actual_adjustment = sum(records[k]['money'][0] for k in new)
            assert actual_adjustment == row[9]
            status = 'Đã phát hành, về 0'
        elif kind == 'replacement':
            target = replacement_roots[root]
            result = next(r for r in audit['groups']['replacement_six'] if tuple(r['root']) == target)
            assert result['balance'] == [0, 0, 0]
            assert set(map(tuple, result['active'])) == set(active)
            assert current_goods - row[6] == row[10]
            status = 'Đã thay thế, về 0'
        else:
            # The retained cohort is independently rechecked, not normalized to0.
            assert not new, (root, 'unexpected new adjustment in retained cohort')
            status = 'Giữ lại, đã kiểm' if active else 'Đã hủy, không cộng DT'
        assert all(records[k]['cqt'] == 'HĐ hợp lệ' for k in new)
        fresh_edges = [(a,b,k,v) for (a,b,k),v in edges.items() if a in nodes and b in nodes]
        link_text = '\n'.join('/'.join(a) + (' — thay thế → ' if k == 'replacement' else ' — điều chỉnh → ') + '/'.join(b) for a,b,k,_ in sorted(fresh_edges))
        current_sources = '\n'.join('/'.join(k) + ' | ' + records[k]['source'] + ' | dòng ' + str(records[k]['row']) for k in (active or sorted(nodes)))
        edge_sources = '\n'.join(v['file'] + ' | dòng ' + str(v['row']) for _,b,_,v in fresh_edges if not v['legacy'] and records[b]['date'].year == 2026)
        if edge_sources:
            current_sources += '\n' + edge_sources
        if kind == 'retained':
            note = 'Đã kiểm lại bảng kê đến 29/10/2026; giữ nguyên giá trị theo danh sách ngoài phạm vi đưa về 0.' if active else 'Hóa đơn đã hủy theo bảng kê; không cộng vào doanh thu.'
        else:
            note = 'Đã phát hành ngày 29/10/2026: ' + ', '.join('/'.join(k) for k in new) + '; chuỗi về 0 đủ 3 khoản.'
            if any(k in confirmed_keys for k in nodes):
                note += ' Thanh toán trống = 0 theo xác nhận 30/10.'
        if len(row) > 17 and row[17]:
            note = str(row[17]) + '\n' + note
        updates += [(rn,13,status),(rn,15,link_text),(rn,16,current_sources),(rn,17,note)]
        used_keys.update(nodes)
        evidence.append({'row':rn,'root':root,'kind':kind,'active':active,'all_keys':sorted(nodes),'new':new,'after_goods':current_goods,'before_goods':row[6], 'adjustment_goods':row[9], 'replacement_reduction':row[10], 'status':status})
    kinds = Counter(e['kind'] for e in evidence)
    assert kinds == {'adjustment':654,'replacement':6,'retained':118}, kinds
    assert sum(e['after_goods'] for e in evidence) == 622963131
    assert sum(e['before_goods'] for e in evidence) == 47402157888
    assert sum(e['adjustment_goods'] for e in evidence) == -46407850480
    assert sum(e['replacement_reduction'] for e in evidence) == -371344277
    assert sum(e['after_goods'] != 0 for e in evidence if e['kind'] == 'retained') == 116
    assert all(e['after_goods'] == 0 for e in evidence if e['kind'] != 'retained')
    top = {
        'A1':'BÁO CÁO DOANH THU 2024 — CẬP NHẬT 30/10/2026',
        'A2':'Đơn vị: đồng, chưa VAT. Gốc năm 2024; kiểm phát hành đến 29/10/2026.',
        'A4':'1. Hóa đơn cũ trước xử lý', 'C4':'Theo gốc năm 2024 và các điều chỉnh/thay thế trước đợt này.',
        'A5':'2. Điều chỉnh đã phát hành', 'C5':'Đã phát hành 1.253 HĐ/654 gốc; các chuỗi về 0 đủ 3 khoản.',
        'A6':'3. Thay thế đã phát hành', 'C6':'Đã ghép đủ 6 cặp; hóa đơn thay thế mới 0đ.',
        'A7':'4. Hóa đơn cũ còn giữ lại', 'C7':'Cũ còn giữ 622.963.131đ; 660 chuỗi thuộc diện xử lý đã về 0.',
        'A8':'5. Bán mới AMIS — dự kiến', 'C8':'Theo bộ AMIS đã chuẩn bị; chưa đối chiếu phần bán mới đã phát hành.',
        'A9':'6. Trả hàng — dự kiến', 'C9':'163 phiếu; chưa đối chiếu HĐ phát hành cho phần trả hàng.',
        'A10':'SAU BÁN MỚI/TRẢ HÀNG — DỰ TÍNH', 'C10':'Phần cũ đã kiểm thực tế; bán mới và trả hàng vẫn theo phương án.',
        'C12':'Chênh với sổ AMIS, không phải số dư của các chuỗi đã xử lý.',
        'C13':'HĐ 111 khác kỳ +3.888.889đ; 16 nhóm khác tiền hàng −1.519.929đ.',
        'D14':'Số HĐ trước xử lý','E14':'HĐ trước xử lý được cộng','F14':'HĐ trước xử lý bị loại',
        'J14':'Điều chỉnh đã phát hành','K14':'Thay thế đã phát hành: giảm DT',
        'N14':'Kết quả sau phát hành','Q14':'Nguồn kiểm tra mới','R14':'Ghi chú sau kiểm tra',
    }
    for addr, value in top.items():
        import re
        m = re.fullmatch(r'([A-Z]+)(\d+)',addr)
        col = 0
        for c in m[1]: col = col*26 + ord(c)-64
        updates.append((int(m[2]),col-1,value))
    assert all((row,col) not in {(r,1) for r in range(1,14)} for row,col,_ in updates)
    def cell(value):
        return {'userEnteredValue':{'numberValue':value}} if isinstance(value,(int,float)) else {'userEnteredValue':{'stringValue':value}}
    # Coalesce vertical blocks, leaving the existing formulas and number cells.
    requests = []
    for col in [13,15,16,17]:
        values = [(r,v) for r,c,v in updates if c == col and r>=15]
        assert [r for r,_ in values] == list(range(15,793))
        for start in range(0,len(values),40):
            part = values[start:start+40]
            requests.append({'updateCells':{'start':{'sheetId':SID,'rowIndex':part[0][0]-1,'columnIndex':col},'rows':[{'values':[cell(v)]} for _,v in part],'fields':'userEnteredValue'}})
    for row,col,value in updates:
        if row>=15: continue
        requests.append({'updateCells':{'start':{'sheetId':SID,'rowIndex':row-1,'columnIndex':col},'rows':[{'values':[cell(value)]}],'fields':'userEnteredValue'}})
    for name,data in [('prepared.json',{'updates':updates,'evidence':evidence,'source_hashes':hashes,'recovered_legacy_edges':recovered,'counts':dict(kinds),'old_retained':622963131,'forecast_after':49930540760,'amis':49928171800,'gap':2368960,'scope':'2024 roots only. Existing-invoice processing confirmed; new AMIS sales/returns remain forecast.'}),('requests.json',requests)]:
        (WORK/name).write_text(json.dumps(data,ensure_ascii=False,indent=2,default=native),encoding='utf-8')
    print(json.dumps({'counts':dict(kinds),'cells':len(updates),'requests':len(requests),'retained':622963131,'unchanged_forecast':49930540760,'unchanged_gap':2368960,'recovered_edges':len(recovered)},ensure_ascii=False))


if __name__=='__main__':main()
