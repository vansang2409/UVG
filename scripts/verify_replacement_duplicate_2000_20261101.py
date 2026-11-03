"""Independently verify the replacement workbook against raw invoice exports."""
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from io import BytesIO
from pathlib import Path
import hashlib
import json
import sys

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / 'scripts/data/replacement-duplicate-2000-20261101'
OUTPUT = ROOT / 'outputs/thay-the-trung-2000-20261101/IMPORT_THAY_THE_VE_0_2000_HD_TRUNG_2024.xlsx'


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    prepared = json.loads((WORK / 'prepared.json').read_text(encoding='utf-8'))
    summary = json.loads((ROOT / 'scripts/data/new-sales-issued-20261101/summary.json').read_text(encoding='utf-8'))
    detail_audit = json.loads((ROOT / 'scripts/data/new-sales-issued-20261101/detail-summary.json').read_text(encoding='utf-8'))
    target = {f'1C26MUV/{n:08}' for n in range(133274, 135274)}
    retained = {f'1C26MUV/{n:08}' for n in range(131274, 133274)}
    heads, detail = {}, defaultdict(list)
    for entry in summary['source_manifest']:
        wb = load_workbook(BytesIO((ROOT / entry['file']).read_bytes()), read_only=True, data_only=True)
        sh = wb.active
        headers = {v: i for i, v in enumerate(next(sh.iter_rows(min_row=5, max_row=5, values_only=True))) if v}
        for r in sh.iter_rows(min_row=6, values_only=True):
            if not isinstance(r[0], (int, float)):
                continue
            get = lambda k: r[headers[k]]
            key = f"{get('Ký hiệu')}/{str(get('Số hóa đơn')).zfill(8)}"
            if key in target:
                assert key not in heads
                heads[key] = {'date': get('Ngày hóa đơn'), 'cqt': get('Mã của CQT'), 'money': [get(c) for c in ['Doanh số bán chưa thuế', 'Tiền thuế GTGT', 'Tổng tiền']]}
        wb.close()
    for entry in detail_audit['source_manifest']:
        wb = load_workbook(BytesIO((ROOT / entry['file']).read_bytes()), read_only=True, data_only=True)
        sh = wb.active
        headers = {v: i for i, v in enumerate(next(sh.iter_rows(min_row=5, max_row=5, values_only=True))) if v}
        for r in sh.iter_rows(min_row=6, values_only=True):
            if not isinstance(r[0], (int, float)):
                continue
            get = lambda k: r[headers[k]]
            key = f"{get('Ký hiệu')}/{str(get('Số hóa đơn')).zfill(8)}"
            if key in target:
                detail[key].append({k: get(k) for k in ['Tên khách hàng', 'MST/CCCD chủ hộ', 'Địa chỉ', 'Người mua hàng', 'Căn cước công dân', 'Hình thức TT', 'Tên hàng', 'ĐVT', 'Thuế suất (%)']})
        wb.close()
    assert set(heads) == set(detail) == target
    wb = load_workbook(OUTPUT, read_only=True, data_only=False)
    assert wb.sheetnames == ['Hóa đơn GTGT']
    sh = wb.active
    template = load_workbook(BytesIO(Path(prepared['template']).read_bytes()), read_only=True, data_only=True)
    assert list(next(sh.iter_rows(min_row=9, max_row=9, max_col=22, values_only=True))) == list(next(template.active.iter_rows(min_row=9, max_row=9, max_col=22, values_only=True)))
    assert list(next(sh.iter_rows(min_row=9, max_row=9, min_col=23, max_col=24, values_only=True))) == ['Lý do thay thế', 'Là dòng ghi chú/diễn giải']
    template.close()
    output_rows = list(sh.iter_rows(min_row=10, max_row=2009, max_col=24, values_only=False))
    groups = defaultdict(list)
    for rn, cells in enumerate(output_rows, 10):
        values = [c.value for c in cells]
        assert isinstance(values[0], int) and 1 <= values[0] <= 2000
        assert not any(c.data_type == 'f' for c in cells)
        assert isinstance(values[1], datetime) and values[1] == datetime(2026, 11, 1)
        assert values[17:22] == [0, 0, 0, 8, 0] and values[23] == 1
        assert values[15] == values[22] and values[16] is None
        assert values[6] is None and values[7] is None and values[10] is None
        expected = prepared['rows'][rn - 10].copy()
        for index in [1, 13]:
            if isinstance(values[index], datetime):
                values[index] = (values[index] - datetime(1899, 12, 30)).days
        assert values == expected, (rn, values, expected)
        groups[values[0]].append(cells)
    assert list(groups) == list(range(1, 2001))
    references = set()
    for seq, cells_rows in groups.items():
        first = [c.value for c in cells_rows[0]]
        key = f'{first[11]}/{first[12]}'
        assert key in target and key not in retained and key not in references
        assert key == f'1C26MUV/{133273 + seq:08}'
        references.add(key)
        assert len(cells_rows) == 1
        assert prepared['audit'][seq - 1]['source_lines'] == len(detail[key])
        assert isinstance(first[12], str) and len(first[12]) == 8 and cells_rows[0][12].number_format == '@'
        assert first[14] == heads[key]['cqt'] and cells_rows[0][14].data_type == 's'
        assert first[20:22] == [8, 0]
        assert isinstance(first[13], datetime) and first[13].date() == datetime.strptime(heads[key]['date'], '%d/%m/%Y').date()
        source = detail[key][0]
        reason = f"Thay thế hóa đơn số {first[12]} ngày {heads[key]['date']} do xuất trùng đơn hàng"
        assert first[15] == first[22] == reason and first[23] == 1 and first[16] is None
        assert first[1] == datetime(2026, 11, 1)
        assert all(line['Thuế suất (%)'] == '8%' for line in detail[key])
        for index, label in [(2, 'Tên khách hàng'), (3, 'MST/CCCD chủ hộ'), (4, 'Địa chỉ'), (5, 'Người mua hàng'), (8, 'Căn cước công dân'), (9, 'Hình thức TT')]:
            assert first[index] == (source[label] or None)
    assert references == target
    assert not any(c.value is not None for row in sh.iter_rows(min_row=2010, max_row=2730, max_col=25) for c in row)
    assert not any(c.value is not None for row in sh.iter_rows(min_row=9, max_row=2009, min_col=25, max_col=25) for c in row)
    wb.close()
    assert [sum(Decimal(str(h['money'][i])) for h in heads.values()) for i in range(3)] == [2883691447, 230695306, 3114386753]
    assert all(sha(ROOT / entry['file']) == entry['sha256'] for entry in prepared['source_manifest'])
    assert sha(Path(prepared['template'])) == prepared['template_sha256']
    result = {'passed': True, 'independent_source_reader': 'openpyxl', 'invoices': len(groups), 'lines': len(output_rows),
              'first_reference': min(references), 'last_reference': max(references), 'reference_date': '2026-11-01', 'reference_cqt_codes_checked': 2000,
              'retained_batch_excluded': True, 'new_dates_blank': False, 'new_date': '2026-11-01', 'one_note_line_per_invoice': True,
              'reason_matches_description': True, 'note_flag_all_one': True, 'buyer_and_reference_data_preserved': True, 'all_quantity_price_goods_vat_zero': True,
              'vat_rate_numeric': 8, 'original_template_columns_exact': 22, 'total_columns': 24,
              'no_example_rows_or_formulas': True, 'sources_unchanged': True, 'before': prepared['before'], 'replacement_totals': [0, 0, 0],
              'sha256': sha(OUTPUT), 'not_imported_or_issued': True}
    (WORK / 'verification.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False))


if __name__ == '__main__':
    main()
