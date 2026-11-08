"""Prepare a bounded append/cell patch; this script never writes to Google."""
from pathlib import Path
import json

BASE = Path(__file__).resolve().parents[1] / 'outputs/import-2025-20261109'
OLD = BASE / 'archive/before-zero-chain-fix-20261109/google-upload'
NEW = BASE / 'google-upload'

def read(path):
    return json.loads(path.read_text(encoding='utf-8'))

def cell(v):
    return {} if v is None else {'userEnteredValue': {'numberValue' if isinstance(v, (int, float)) else 'stringValue': v}}

def update(sid, row, col, values):
    return {'updateCells': {'start': {'sheetId': sid, 'rowIndex': row, 'columnIndex': col}, 'rows': [{'values': [cell(v) for v in line]} for line in values], 'fields': 'userEnteredValue'}}

def main():
    old, new = read(OLD / '0-matrix.json'), read(NEW / '0-matrix.json')
    before, after = read(OLD / '700000103-matrix.json'), read(NEW / '700000103-matrix.json')
    assert len(old) == 34521 and len(new) == 34990
    assert old[9:] == new[9:len(old)]
    assert [(i,j) for i in range(9) for j in range(22) if old[i][j] != new[i][j]] == [(1,0)]
    changed = [(i,j) for i in range(len(before)) for j in range(31) if before[i][j] != after[i][j]]
    rownums = [215,216,250,251,286]
    assert changed == [(3,0)] + [(i,j) for i in rownums for j in [28,29,30]]
    checks = [{'sheetId': 0, 'title': 'Import HD moi', 'range': 'A2', 'startRow': 1, 'startColumn': 0, 'values': [[old[1][0]]]},
              {'sheetId': 0, 'title': 'Import HD moi', 'range': 'A34521:V34521', 'startRow': 34520, 'startColumn': 0, 'values': [old[-1]]},
              {'sheetId': 700000103, 'title': 'Dieu chinh BTL - cho so HD', 'range': 'A4', 'startRow': 3, 'startColumn': 0, 'values': [[before[3][0]]]}]
    requests = [{'appendDimension': {'sheetId': 0, 'dimension': 'ROWS', 'length': len(new)-len(old)}},
                {'copyPaste': {'source': {'sheetId': 0, 'startRowIndex': len(old)-1, 'endRowIndex': len(old), 'startColumnIndex': 0, 'endColumnIndex': 22}, 'destination': {'sheetId': 0, 'startRowIndex': len(old), 'endRowIndex': len(new), 'startColumnIndex': 0, 'endColumnIndex': 22}, 'pasteType': 'PASTE_FORMAT'}},
                {'updateDimensionProperties': {'range': {'sheetId': 0, 'dimension': 'ROWS', 'startIndex': len(old), 'endIndex': len(new)}, 'properties': {'pixelSize': 38}, 'fields': 'pixelSize'}},
                update(0, len(old), 0, new[len(old):]), update(0, 1, 0, [[new[1][0]]]), update(700000103, 3, 0, [[after[3][0]]])]
    for i in rownums:
        checks.append({'sheetId': 700000103, 'title': 'Dieu chinh BTL - cho so HD', 'range': f'AC{i+1}:AE{i+1}', 'startRow': i, 'startColumn': 28, 'values': [before[i][28:31]]})
        requests.append(update(700000103, i, 28, [after[i][28:31]]))
    plan = {'spreadsheetId': read(NEW / 'manifest.json')['spreadsheetId'], 'oldRows': len(old), 'newRows': len(new), 'existingValueChanges': len(changed)+1, 'appendedRows': len(new)-len(old), 'checks': checks, 'requests': requests}
    (NEW / 'zero-chain-patch.json').write_text(json.dumps(plan, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    # Transfer plain values separately to avoid verbose CellData being truncated by the CLI bridge.
    (NEW / 'zero-chain-appended-values.json').write_text(json.dumps(new[len(old):], ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    compact = {**plan, 'requests': [r for i, r in enumerate(requests) if i != 3]}
    (NEW / 'zero-chain-patch-meta.json').write_text(json.dumps(compact, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    print(json.dumps({k: plan[k] for k in ['oldRows','newRows','existingValueChanges','appendedRows']}))

if __name__ == '__main__':
    main()
