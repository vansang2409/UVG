"""Rebuild UVG chain reports from the refreshed sources without editing source files."""
from __future__ import annotations

import hashlib
import io
import json
import re
import subprocess
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from openpyxl import load_workbook
from filter_invoice_exceptions import is_company, normal, tax_shaped

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'data/Du lieu Meinvoice'
HELPERS = ROOT / 'scripts/data'
START, END, CUTOFF = date(2024, 1, 1), date(2025, 6, 30), date(2025, 12, 31)
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
ZERO = Decimal(0)
HISTORICAL_REVISION = '9c50b0424a242695fa31182ab826613bf0fa699f'
USER_CONFIRMED_DATES = {
    ('1C25TUV', '00012885'): {
        'date': date(2025, 7, 5),
        'confirmed_on': '07/10/2026',
        'evidence': 'Người dùng xác nhận và gửi ảnh danh sách Hóa đơn bán hàng có mã CQT trên meInvoice; chưa đối chiếu XML độc lập',
    },
}


def text(v):
    return '' if v is None else str(v)


def invoice_key(signature, number):
    signature, number = text(signature).strip().upper(), text(number).strip()
    if not signature or not number:
        return None
    if number.endswith('.0') and number[:-2].isdigit():
        number = number[:-2]
    return signature, number.zfill(8) if number.isdigit() else number


def invoice_day(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.strptime(text(v).strip(), fmt).date()
        except ValueError:
            pass
    if v not in (None, ''):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
        except (ValueError, OverflowError):
            pass
    return None


def amount(v):
    if v in (None, ''):
        return None
    try:
        return Decimal(str(v))
    except Exception:
        return None


def native(v):
    if isinstance(v, Decimal):
        return int(v) if v == v.to_integral_value() else float(v)
    return v


def rows(path):
    """Read OOXML .xls/.xlsx, preserving identifier strings and source row numbers."""
    with ZipFile(path) as archive:
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            xml = etree.fromstring(archive.read('xl/sharedStrings.xml'))
            strings = [''.join(t.text or '' for t in entry.iter(NS + 't')) for entry in xml]
        with archive.open('xl/worksheets/sheet1.xml') as stream:
            for _, row in etree.iterparse(stream, events=('end',), tag=NS + 'row'):
                values = {}
                for cell in row:
                    column = re.sub(r'\d', '', cell.get('r', ''))
                    v = cell.find(NS + 'v')
                    if cell.get('t') == 's' and v is not None:
                        value = strings[int(v.text)]
                    elif cell.get('t') == 'inlineStr':
                        value = ''.join(t.text or '' for t in cell.iter(NS + 't'))
                    else:
                        value = v.text if v is not None else None
                    values[column] = value
                yield int(row.get('r', '0')), values
                row.clear()
                while row.getprevious() is not None:
                    del row.getparent()[0]


def main():
    files = sorted(p for p in (ROOT / 'data').rglob('*')
                   if p.is_file() and p.suffix.lower() in ('.xls', '.xlsx') and not p.name.startswith('~$'))
    manifest = [{'path': p.relative_to(ROOT).as_posix(), 'bytes': p.stat().st_size,
                 'sha256': hashlib.sha256(p.read_bytes()).hexdigest()} for p in files]
    source_hashes = {entry['path']: entry['sha256'] for entry in manifest}
    records = defaultdict(lambda: {'observations': [], 'date': None, 'header': None, 'detail': []})
    edges = {}
    read_counts = {}

    def observe(k, name, tax, day, source, family):
        if not k:
            return
        rec = records[k]
        if rec['date'] is None and day is not None:
            rec['date'] = day
        obs = {'name': text(name), 'tax': text(tax), 'date': day, 'source': source, 'family': family}
        if not any((o['name'], o['tax'], o['family'], o['date']) == (obs['name'], obs['tax'], family, day)
                   for o in rec['observations']):
            rec['observations'].append(obs)

    def add_edge(old, new, kind, child_day, evidence, historical=False):
        if not old or not new or not child_day or child_day > CUTOFF:
            return
        k = (old, new, kind)
        if k not in edges:
            edges[k] = {'old': old, 'new': new, 'kind': kind, 'date': child_day,
                        'sources': [], 'historical': historical}
        assert edges[k]['date'] == child_day, (k, 'Conflicting relationship dates')
        if evidence not in edges[k]['sources']:
            edges[k]['sources'].append(evidence)

    totals = sorted((SOURCE / 'Bang_ke_hoa_don_da_su_dung_2024_2025').glob('*.xlsx'))
    assert len(totals) == 10, 'Expected all ten total-report parts'
    for p in totals:
        count = 0
        for rn, r in rows(p):
            if rn < 6 or amount(r.get('A')) is None:
                continue
            k = invoice_key(r.get('B'), r.get('C'))
            if not k:
                continue
            day = invoice_day(r.get('D'))
            assert day is not None and START <= day <= CUTOFF, (p.name, rn, day)
            ref = f'{p.relative_to(ROOT).as_posix()} | dòng {rn}'
            observe(k, r.get('H') or r.get('L'), r.get('J'), day, ref, 'Bảng tổng')
            h = {'date': day, 'name': text(r.get('H') or r.get('L')), 'tax': text(r.get('J')),
                 'status': text(r.get('U')), 'money': [amount(r.get(c)) for c in ['P', 'Q', 'R']],
                 'source': ref}
            if records[k]['header']:
                previous = records[k]['header']
                assert all(previous[c] == h[c] for c in ['date', 'name', 'tax', 'status', 'money']), k
                raise AssertionError(f'Duplicate invoice in total parts: {k}')
            records[k]['header'] = h
            records[k]['date'] = day
            count += 1
        read_counts[p.relative_to(ROOT).as_posix()] = count

    detail_keys = set()
    for name in ['Bang_ke_chi_tiet_HD_da_su_dung_2024.xls', 'Bang_ke_chi_tiet_HD_da_su_dung_2025.xls',
                 'Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls']:
        p = SOURCE / name
        is_mtt = '_MTT' in name
        identities = ('F', 'B', 'D', 'H', 'I') if is_mtt else ('D', 'B', 'C', 'E', 'G')
        money_cols = ['W', 'AA', 'AB'] if is_mtt else ['P', 'T', 'U']
        discount_col = 'Y' if is_mtt else 'R'
        count = 0
        for rn, r in rows(p):
            if rn < 7 or amount(r.get('A')) is None:
                continue
            sig, number, day_col, name_col, tax_col = identities
            k = invoice_key(r.get(sig), r.get(number))
            if not k:
                continue
            day = invoice_day(r.get(day_col))
            assert k in records and records[k]['header'], (name, rn, k, 'Missing total invoice')
            assert day is not None, (name, rn, k, 'Unparseable detail date')
            ref = f'{p.relative_to(ROOT).as_posix()} | dòng {rn}'
            observe(k, r.get(name_col), r.get(tax_col), day, ref, 'Chi tiết')
            line = [amount(r.get(c)) for c in money_cols]
            discount = amount(r.get(discount_col))
            if line[0] is not None and discount is not None:
                line[0] -= discount
            records[k]['detail'].append(line)
            detail_keys.add(k)
            count += 1
        read_counts[p.relative_to(ROOT).as_posix()] = count
    header_keys = {k for k, rec in records.items() if rec['header']}
    assert detail_keys == header_keys, 'Detail coverage differs from totals'
    print(json.dumps({'phase': 'current_sources', 'total_invoices': len(header_keys)}, ensure_ascii=False), flush=True)

    link_files = [('Điều chỉnh', 'Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_dieu_chinh_2024.xls'),
                  ('Điều chỉnh', 'Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_dieu_chinh_2025.xls'),
                  ('Thay thế', 'Bang_ke_hoa_don_thay_the_xuat_theo_HD_thay_the_2024_2025.xls')]
    for kind, name in link_files:
        p = SOURCE / name
        count = 0
        for rn, r in rows(p):
            if rn < 7 or amount(r.get('A')) is None:
                continue
            old, new = invoice_key(r.get('B'), r.get('D')), invoice_key(r.get('M'), r.get('O'))
            if not old or not new:
                continue
            ref = f'{p.relative_to(ROOT).as_posix()} | dòng {rn}'
            observe(old, r.get('F'), r.get('G'), invoice_day(r.get('C')), ref, 'Liên kết')
            observe(new, r.get('Q'), r.get('R'), invoice_day(r.get('N')), ref, 'Liên kết')
            add_edge(old, new, kind, invoice_day(r.get('N')), ref)
            count += 1
        read_counts[p.relative_to(ROOT).as_posix()] = count
    new_edge_count = len(edges)

    # Preserve proven 2024 relations omitted by the new export; never restore/edit source files.
    revision = HISTORICAL_REVISION
    old_manifest = json.loads(subprocess.check_output(['git', 'show', f'{revision}:SOURCE_MANIFEST.json'], cwd=ROOT))
    old_hashes = {e['path']: e['sha256'] for e in old_manifest}
    historical_sources = []
    for name, kind in [('DS bị điều chỉnh 2024.xls', 'Điều chỉnh'), ('DS bị điều chỉnh 2025.xls', 'Điều chỉnh'),
                       ('DS bị thay thế 2024.xls', 'Thay thế')]:
        rel = 'data/Du lieu Meinvoice/' + name
        content = subprocess.check_output(['git', 'show', f'{revision}:{rel}'], cwd=ROOT)
        digest = hashlib.sha256(content).hexdigest()
        assert digest == old_hashes[rel], rel
        historical_sources.append({'git_revision': revision, 'path': rel, 'sha256': digest})
        w = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        for rn, r in enumerate(w.worksheets[0].values, 1):
            if not isinstance(r[0], (int, float)):
                continue
            old, new = invoice_key(r[1], r[3]), invoice_key(r[12], r[14])
            child_day = invoice_day(r[13])
            if not child_day or child_day > CUTOFF or (old, new, kind) in edges:
                continue
            assert old in header_keys and new in header_keys, (rel, rn, old, new)
            ref = f'Git {revision[:8]}:{rel} | dòng {rn}'
            observe(old, r[5], r[6], invoice_day(r[2]), ref, 'Git lịch sử')
            observe(new, r[16], r[17], child_day, ref, 'Git lịch sử')
            add_edge(old, new, kind, child_day, ref, historical=True)
        w.close()

    children, parents, undirected = defaultdict(set), defaultdict(set), defaultdict(set)
    for edge in edges.values():
        old, new = edge['old'], edge['new']
        children[old].add(new)
        parents[new].add(old)
        undirected[old].add(new)
        undirected[new].add(old)
    direct_exception = {k for k, rec in records.items()
                        if any(is_company(o['name']) or tax_shaped(o['tax']) for o in rec['observations'])}
    protected, seen = set(), set()
    for start in records:
        if start in seen:
            continue
        component, queue = set(), [start]
        while queue:
            k = queue.pop()
            if k in component:
                continue
            component.add(k)
            queue.extend(undirected[k] - component)
        seen.update(component)
        if component & direct_exception:
            protected.update(component)

    roots = sorted(k for k in header_keys if START <= records[k]['date'] <= END and not parents[k]
                   and normal(records[k]['header']['status']) not in ('hoa don dieu chinh', 'hoa don thay the'))
    groups = {'exceptions': [], 'adjusted': []}
    exclusions, unresolved = [], []
    for root in roots:
        nodes, queue = set(), [root]
        while queue:
            k = queue.pop()
            if k in nodes:
                continue
            if records[k]['date'] is None or records[k]['date'] > CUTOFF:
                continue
            nodes.add(k)
            queue.extend(children[k])
        linked = [e for e in edges.values() if e['old'] in nodes and e['new'] in nodes]
        is_exception = root in protected
        has_replacement = any(e['kind'] == 'Thay thế' for e in linked)
        has_adjustment = any(e['kind'] == 'Điều chỉnh' for e in linked)
        status = normal(records[root]['header']['status'])
        if not is_exception:
            if has_replacement:
                exclusions.append({'root': list(root), 'reason': 'Chuỗi thay thế trong kỳ; chưa xử lý ở nhóm điều chỉnh'})
                continue
            if not has_adjustment:
                if status in ('hoa don da bi dieu chinh', 'hoa don da bi thay the'):
                    unresolved.append({'root': list(root), 'date': str(records[root]['date']),
                                       'status': records[root]['header']['status'],
                                       'reason': 'Chưa có liên kết xác nhận xử lý đến 31/12/2025; trạng thái có thể phản ánh lần xử lý sau kỳ'})
                continue
        pending, ordered, flags = set(nodes), [], []
        while pending:
            available = [k for k in pending if not (parents[k] & pending)]
            if not available:
                flags.append('Vòng lặp liên kết; chưa chốt cộng dồn')
                available = list(pending)
            chosen = min(available, key=lambda k: (records[k]['date'] or date.max, k))
            ordered.append(chosen)
            pending.remove(chosen)
        assert ordered[0] == root
        multiparent = any(len(parents[k]) > 1 for k in nodes)
        if multiparent:
            flags.append('Hóa đơn sửa tham chiếu nhiều gốc; chưa phân bổ, cộng dồn giữ trống')
        if not has_adjustment and status == 'hoa don da bi dieu chinh':
            flags.append('Nhãn hiện tại đã bị điều chỉnh; chưa có liên kết điều chỉnh trong kỳ')
        if not has_replacement and any(normal(records[k]['header']['status']) == 'hoa don da bi thay the'
                                       for k in nodes if records[k]['header']):
            flags.append('Nhãn hiện tại đã bị thay thế; chưa xác nhận ngày thay thế trong kỳ')
        replaced = {e['old'] for e in linked if e['kind'] == 'Thay thế'}
        invoices, active_money, detail = [], [], []
        for k in ordered:
            rec, h = records[k], records[k]['header']
            money = h['money'][:] if h else [None, None, None]
            if h is None:
                flags.append(f'{k[0]}/{k[1]} thiếu bảng tổng')
            refs = [e for e in linked if e['new'] == k]
            invoice_kind = 'Hóa đơn gốc' if k == root else '; '.join(sorted(set(
                f"{e['kind']} cho {e['old'][0]}/{e['old'][1]}" for e in refs)))
            observations = rec['observations']
            best = next((o for o in observations if o['family'] == 'Bảng tổng'), observations[0])
            if is_exception:
                best = next((o for o in observations if is_company(o['name']) or tax_shaped(o['tax'])), best)
            chosen_name, chosen_tax = best['name'], best['tax']
            source_refs = [h['source']] if h else []
            source_refs.append(best['source'])
            source_refs += [source for e in refs for source in e['sources']]
            if k in replaced:
                flags.append(f'{k[0]}/{k[1]} có bản thay thế trong kỳ; không cộng bản bị thay thế')
            cancelled = h and 'huy' in normal(h['status'])
            if cancelled:
                flags.append(f'{k[0]}/{k[1]} nguồn ghi đã hủy; loại tiền khỏi cộng dồn')
            used = money[:]
            if k == ('1C24TUV', '00000574') and used[2] is None:
                used[2] = ZERO
                flags.append('Riêng 1C24TUV/00000574: tạm tính thanh toán gốc = 0 theo xác nhận; ô nguồn giữ trống')
            elif any(v is None for v in money):
                flags.append(f'{k[0]}/{k[1]} có ô tiền bảng tổng trống; không tự điền 0')
            if all(v is not None for v in money) and money[0] + money[1] != money[2]:
                flags.append(f'{k[0]}/{k[1]} tiền hàng + thuế lệch thanh toán')
            if len({o['tax'].strip() for o in observations if o['tax'].strip()}) > 1:
                flags.append(f'{k[0]}/{k[1]} có khác biệt định danh giữa nguồn; xem chi tiết')
            observed_days = sorted({o['date'] for o in observations if o['date']})
            confirmation = USER_CONFIRMED_DATES.get(k)
            if len(observed_days) > 1:
                date_authority = (f"bảng chính dùng {confirmation['date'].strftime('%d/%m/%Y')} theo xác nhận và ảnh danh sách meInvoice ngày {confirmation['confirmed_on']}; chưa đối chiếu XML"
                                  if confirmation else 'đang theo bảng tổng')
                flags.append(f"{k[0]}/{k[1]} ngày giữa nguồn khác nhau ({', '.join(d.strftime('%d/%m/%Y') for d in observed_days)}); {date_authority}")
                for observed_day in observed_days:
                    observation = next(o for o in observations if o['date'] == observed_day)
                    source_refs.append(f"{observation['source']} (ngày {observed_day.strftime('%d/%m/%Y')})")
            if confirmation:
                source_refs.append(f"Ngày bảng chính {confirmation['date'].strftime('%d/%m/%Y')}: {confirmation['evidence']} ({confirmation['confirmed_on']})")
            if not cancelled and k not in replaced:
                active_money.append(used)
            display_date = confirmation['date'] if confirmation else rec['date']
            invoices.append([*k, display_date.strftime('%d/%m/%Y'), invoice_kind, *map(native, money)])
            detail.append([*root, *k, rec['date'].strftime('%d/%m/%Y'), chosen_name, chosen_tax,
                           invoice_kind, h['status'] if h else 'Không có trong bảng tổng',
                           *map(native, money), '; '.join(dict.fromkeys(source_refs))])
        sums = [native(sum((m[i] for m in active_money), ZERO))
                if not multiparent and all(m[i] is not None for m in active_money) else None for i in range(3)]
        root_obs = records[root]['observations']
        preferred = next((o for o in root_obs if is_company(o['name']) or tax_shaped(o['tax'])), root_obs[0])
        if is_exception and root not in direct_exception:
            flags.append('Ngoại lệ theo liên kết với hóa đơn có công ty/định danh dạng MST trong cùng chuỗi')
        group = {'root': list(root), 'name': preferred['name'], 'tax': preferred['tax'], 'invoices': invoices,
                 'sums': sums, 'notes': '; '.join(dict.fromkeys(flags)), 'detail': detail}
        groups['exceptions' if is_exception else 'adjusted'].append(group)

    summaries = {}
    for name, entries in groups.items():
        invoice_keys = {tuple(row[:2]) for g in entries for row in g['invoices']}
        summaries[name] = {'roots': len(entries), 'invoices': len(invoice_keys),
                           'displayed_invoice_rows': sum(len(g['invoices']) for g in entries),
                           'max_blocks': max([3, *[len(g['invoices']) for g in entries]]),
                           'roots_with_notes': sum(bool(g['notes']) for g in entries),
                           'roots_by_year': dict(Counter(g['invoices'][0][2][-4:] for g in entries))}
    assert not ({tuple(g['root']) for g in groups['exceptions']} & {tuple(g['root']) for g in groups['adjusted']})
    assert all(START <= datetime.strptime(g['invoices'][0][2], '%d/%m/%Y').date() <= END
               for entries in groups.values() for g in entries)
    assert all(datetime.strptime(r[2], '%d/%m/%Y').date() <= CUTOFF
               for entries in groups.values() for g in entries for r in g['invoices'])
    historical_edges = [e for e in edges.values() if e['historical']]
    payload = {'scope': {'f0_start': str(START), 'f0_end': str(END), 'related_end': str(CUTOFF)},
               'groups': groups, 'summary': summaries,
               'user_confirmed_dates': [{'invoice': list(k), **v, 'date': v['date'].strftime('%d/%m/%Y')}
                                        for k, v in USER_CONFIRMED_DATES.items()],
               'source_summary': {'files': len(manifest), 'total_invoices': len(header_keys),
                                  'detail_invoices': len(detail_keys), 'fresh_unique_links': new_edge_count,
                                  'historical_links_added': len(historical_edges), 'candidate_f0': len(roots),
                                  'unresolved_roots': len(unresolved), 'excluded_replacement_roots': len(exclusions),
                                  'date_conflict_invoices': sum(len({o['date'] for o in r['observations'] if o['date']}) > 1
                                                               for r in records.values())},
               'read_counts': read_counts, 'manifest': manifest, 'historical_sources': historical_sources,
               'historical_edges': [{**e, 'old': list(e['old']), 'new': list(e['new']), 'date': str(e['date'])}
                                    for e in historical_edges],
               'excluded': exclusions, 'unresolved': unresolved}
    # Confirm source bytes did not change during the read before recording the new manifest.
    for rel, digest in source_hashes.items():
        assert hashlib.sha256((ROOT / rel).read_bytes()).hexdigest() == digest, rel
    HELPERS.mkdir(parents=True, exist_ok=True)
    (HELPERS / 'invoice_reports.json').write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')
    (ROOT / 'SOURCE_MANIFEST.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'reports': summaries, 'sources': payload['source_summary']}, ensure_ascii=False), flush=True)


if __name__ == '__main__':
    main()
