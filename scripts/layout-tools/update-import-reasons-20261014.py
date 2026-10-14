from pathlib import Path
from openpyxl import load_workbook
from collections import Counter
import json, shutil, hashlib, sys
sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / 'scripts/data/import-reasons-20261014'
OUT.mkdir(parents=True, exist_ok=True)
files = []
for name in ['HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx', 'HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx']:
    src = ROOT / 'outputs' / name
    baseline = OUT / 'before' / name
    baseline.parent.mkdir(exist_ok=True)
    if baseline.exists():
        raise RuntimeError('Bản trước sửa đã tồn tại; không ghi đè')
    shutil.copy2(src, baseline)
    w = load_workbook(src, read_only=True, data_only=True)
    data = [(i, r) for i, r in enumerate(w['Import dieu chinh'].values, 1) if i >= 10 and isinstance(r[0], (int, float))]
    groups = {}
    for i, r in data:
        groups.setdefault(r[0], []).append((i, r))
    changes = []
    counts = Counter()
    for seq, rows in groups.items():
        first = rows[0][1]
        # Ưu tiên chiều tổng thanh toán, sau đó tiền hàng/thuế; dòng chỉ giảm SL vẫn là giảm.
        amount = next((first[j] for j in [22, 19, 21] if isinstance(first[j], (int, float)) and first[j] != 0), None)
        if amount is None:
            quantities = [r[16] for _, r in rows if isinstance(r[16], (int, float)) and r[16] != 0]
            direction = 'tăng' if quantities and sum(quantities) > 0 else 'giảm'
        else:
            direction = 'tăng' if amount > 0 else 'giảm'
        counts[direction] += 1
        for i, r in rows:
            assert r[9:12] == first[9:12], (seq, i, 'Tham chiếu không đồng nhất')
            text = f'Điều chỉnh {direction} cho hóa đơn ký hiệu {r[9]}, số {r[10]}, ngày {r[11]} do xuất sai nội dung'
            changes.append({'row': i, 'old': r[13], 'new': text})
    files.append({'name': name, 'sha256_before': hashlib.sha256(src.read_bytes()).hexdigest(), 'changes': changes, 'counts': dict(counts), 'invoices': len(groups)})
    w.close()
(OUT / 'plan.json').write_text(json.dumps({'files': files}, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps([{k: v for k, v in f.items() if k != 'changes'} for f in files], ensure_ascii=False))
