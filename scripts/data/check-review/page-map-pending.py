"""Map pending invoices to 100-row pages using the current export and received first page."""
from pathlib import Path
from zipfile import ZipFile
from collections import Counter
from datetime import date
import io
import json
import sys

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / "scripts"))
from prepare_invoice_reports import rows, invoice_day, invoice_key, amount

sys.stdout.reconfigure(encoding="utf-8")
folder = ROOT / "data/Du lieu Meinvoice/Xuat_lai_2026-10-09"
records = {}
for path in [folder / "Bang_ke_hoa_don_da_su_dung_1791532882546.xlsx",
             folder / "Bang_ke_hoa_don_da_su_dung_1791532908545.zip"]:
    if path.suffix == ".zip":
        with ZipFile(path) as z:
            books = [(name, z.read(name)) for name in z.namelist() if name.endswith((".xlsx", ".xls"))]
    else:
        books = [(path.name, path.read_bytes())]
    for name, data in books:
        for rn, row in rows(io.BytesIO(data)):
            if amount(row.get("A")) is None:
                continue
            key, day = invoice_key(row.get("B"), row.get("C")), invoice_day(row.get("D"))
            if key and key[0] == "1C25TUV" and day and date(2025, 5, 4) <= day <= date(2025, 6, 29):
                records[key] = {"date": day.isoformat(), "status": row.get("U"), "source_row": rn, "source_file": name}

ordered = sorted(records)
positions = {key: n for n, key in enumerate(ordered)}
pending = json.loads((ROOT / "scripts/data/check-review/pending-current-audit-20261011.json").read_text(encoding="utf-8"))
received = json.loads((ROOT / "scripts/data/check-review/pending-xml-3a74a6eb65a1.json").read_text(encoding="utf-8"))
first_page = sorted(tuple(x["key"]) for x in received["records"])
by_page = {}
for raw in pending["pending_adjustment_xml_not_in_repo"]:
    key = tuple(raw)
    if key not in positions:
        raise ValueError(("Target not found", key))
    page = positions[key] // 100 + 1
    by_page.setdefault(str(page), []).append({"number": key[1], "date": records[key]["date"],
                                           "row_in_page": positions[key] % 100 + 1})
result = {"filters": {"signature": "1C25TUV", "date_from": "2025-05-04", "date_to": "2025-06-29",
                      "rows_per_page": 100, "sort": "invoice number ascending"},
          "invoice_count": len(ordered), "page_count": (len(ordered) + 99) // 100,
          "first_page_exactly_matches_received_zip": ordered[:100] == first_page,
          "statuses": dict(Counter(x["status"] for x in records.values())), "needed_pages": by_page,
          "page_ranges": {page: [ordered[(int(page)-1)*100][1], ordered[min(int(page)*100, len(ordered))-1][1]] for page in by_page}}
(ROOT / "scripts/data/check-review/pending-page-map-20261011.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
print(json.dumps(result, ensure_ascii=False, indent=2))
