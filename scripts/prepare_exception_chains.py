"""Chuyển ngoại lệ năm 2024 thành mỗi dòng một gốc, các hóa đơn sửa đặt ngang."""
import json
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs/ngoai-le-meinvoice-2024'
payload = json.loads((OUT / 'workbook_data.json').read_text(encoding='utf-8'))
records = {(r[0], r[1]): r for s in payload['sheets'] for r in s['rows']}
children, parents = defaultdict(set), defaultdict(set)
evidence = []
for path in sorted((ROOT / 'data/Du lieu Meinvoice').glob('DS*.xls')):
    with path.open('rb') as stream:
        wb = load_workbook(stream, read_only=True, data_only=True)
        for i, r in enumerate(wb.worksheets[0].values, 1):
            if not isinstance(r[0], (int, float)):
                continue
            old, new = (str(r[1]), str(r[3])), (str(r[12]), str(r[14]))
            if old in records and new in records:
                children[old].add(new)
                parents[new].add(old)
                evidence.append([*old, *new, 'Điều chỉnh' if 'điều chỉnh' in path.name else 'Thay thế',
                                 path.relative_to(ROOT).as_posix(), i])
        wb.close()
roots = sorted(k for k, r in records.items() if not parents[k] and datetime.strptime(r[2], '%d/%m/%Y').year == 2024)
assert len(roots) == 118
groups = []
covered = set()
for root in roots:
    descendants, queue = set(), [root]
    while queue:
        node = queue.pop()
        if node in descendants:
            continue
        descendants.add(node)
        queue.extend(children[node])
    # Xếp theo phụ thuộc trước, rồi ngày/số; các điều chỉnh cùng gốc không bị coi là sửa nối tiếp.
    remaining, ordered = set(descendants), []
    while remaining:
        available = [k for k in remaining if not (parents[k] & remaining)]
        assert available, 'Vòng lặp liên kết'
        chosen = min(available, key=lambda k: (datetime.strptime(records[k][2], '%d/%m/%Y'), k))
        ordered.append(chosen)
        remaining.remove(chosen)
    assert ordered[0] == root
    assert not (covered & descendants), 'Có hóa đơn thuộc nhiều gốc, cần kiểm tra'
    covered.update(descendants)
    groups.append(ordered)
assert covered == set(records) and len(covered) == 190
block_headers = ['Ký hiệu', 'Số hóa đơn', 'Ngày hóa đơn', 'Tên người mua', 'MST/CCCD chủ hộ nguyên bản',
                 'Trạng thái theo bảng tổng', 'Phân loại', 'Căn cứ nhận diện', 'Tiền hàng theo bảng tổng',
                 'Thuế theo bảng tổng', 'Thanh toán theo bảng tổng']
max_blocks = max(map(len, groups))
headers = ['Tên người mua (gốc)', 'MST/CCCD chủ hộ nguyên bản (gốc)']
for i in range(max_blocks):
    headers.extend(f'F{i} – {h}' for h in block_headers)
rows = []
for group in groups:
    first = records[group[0]]
    row = [first[3], first[4]]
    for i, key in enumerate(group):
        r = records[key]
        kind = 'Hóa đơn gốc' if i == 0 else '; '.join(f'{e[4]} cho {e[0]}/{e[1]}' for e in evidence if (e[2], e[3]) == key)
        # Thiếu trạng thái bảng tổng ghi đúng nguồn DS, không coi là chưa phát hành.
        status = 'Không có trong bảng tổng' if r[5] == 'Chưa có trạng thái bảng tổng' else r[5]
        row.extend([r[0], r[1], r[2], r[3], r[4], status, kind, r[7], *r[8:11]])
    row.extend([None] * (len(headers)-len(row)))
    rows.append(row)
sources = [[*k, records[k][14]] for k in sorted(records)]
result = {'headers': headers, 'rows': rows, 'block_headers': block_headers, 'max_blocks': max_blocks,
          'roots': len(roots), 'invoice_count': len(records), 'edges': evidence, 'sources': sources}
(OUT / 'chains_data.json').write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')
print(json.dumps({k: result[k] for k in ['roots', 'invoice_count', 'max_blocks']}))
