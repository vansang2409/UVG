"""Preserve and audit a newly supplied XML/PDF ZIP without editing workbooks."""
from pathlib import Path
from zipfile import ZipFile
import argparse
import hashlib
import json
import re
import sys
import xml.etree.ElementTree as ET

from audit_pending_review import first, value, arithmetic, compare_xml_money
from prepare_invoice_reports import amount, invoice_day, invoice_key

ROOT = Path(__file__).resolve().parents[1]
DEST = ROOT / "data/Bo sung/meInvoice/2026-10-11-pending-XML"
REPORTS = ROOT / "scripts/data/check-review"


def digest(data):
    return hashlib.sha256(data).hexdigest()


def preserve(path, data):
    if path.exists():
        assert path.read_bytes() == data, ("Existing source differs", path)
    else:
        path.write_bytes(data)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("archive", type=Path)
    args = parser.parse_args()
    archive_data = args.archive.read_bytes()
    archive_sha = digest(archive_data)
    pending = json.loads((REPORTS / "pending-current-audit-20261011.json").read_text(encoding="utf-8"))
    wanted = {tuple(i["key"]): {"root": c["root"], **i} for c in pending["pending"] for i in c["invoices"]}
    originals = set()
    for path in (ROOT / "data/Bo sung/meInvoice").rglob("*.xml"):
        document = ET.fromstring(path.read_bytes())
        tt = first(first(document, "DLHDon"), "TTChung")
        originals.add(invoice_key((value(tt, "KHMSHDon") or "") + (value(tt, "KHHDon") or ""), value(tt, "SHDon")))
    before = {path.name: digest(path.read_bytes()) for path in (ROOT / "outputs").glob("*.xlsx")}
    records, pdfs, errors, seen = [], [], [], set()
    with ZipFile(args.archive) as archive:
        for entry in archive.infolist():
            if entry.is_dir():
                continue
            if entry.filename.lower().endswith(".pdf"):
                pdfs.append({"entry": entry.filename, "bytes": entry.file_size,
                             "sha256": digest(archive.read(entry))})
                continue
            if not entry.filename.lower().endswith(".xml"):
                continue
            data = archive.read(entry)
            try:
                document = ET.fromstring(data)
                hd = first(document, "DLHDon")
                tt, pay = first(hd, "TTChung"), first(hd, "TToan")
                rel = first(tt, "TTHDLQuan")
                key = invoice_key((value(tt, "KHMSHDon") or "") + (value(tt, "KHHDon") or ""), value(tt, "SHDon"))
                assert key and re.fullmatch(r"[A-Z0-9]+", key[0]) and re.fullmatch(r"\d{8}", key[1]), key
                assert key not in seen, ("Duplicate invoice in archive", key)
                seen.add(key)
                reference = invoice_key((value(rel, "KHMSHDCLQuan") or "") + (value(rel, "KHHDCLQuan") or ""),
                                        value(rel, "SHDCLQuan")) if rel is not None else None
                money = [amount(value(pay, f)) for f in ("TgTCThue", "TgTThue", "TgTTTBSo")]
                lines = []
                for line in hd.iter():
                    if line.tag.rsplit("}", 1)[-1] != "HHDVu":
                        continue
                    extra = {value(x, "TTruong"): value(x, "DLieu") for x in line.iter() if x.tag.rsplit("}", 1)[-1] == "TTin"}
                    lines.append({"name": value(line, "THHDVu"), "nature": value(line, "TChat"),
                                  "qty": amount(value(line, "SLuong")), "price": amount(value(line, "DGia")),
                                  "goods": amount(value(line, "ThTien")), "rate": value(line, "TSuat"),
                                  "tax_extra": amount(extra.get("VATAmount"))})
                row = {"key": key, "date": value(tt, "NLap"), "related": reference, "money": money,
                       "related_note": value(rel, "GChu"),
                       "header_arithmetic_difference": arithmetic(money), "lines": lines,
                       "line_goods_sum": sum((x["goods"] for x in lines if x["goods"] is not None), amount(0)),
                       "entry": entry.filename, "sha256": digest(data), "bytes": len(data),
                       "file": (DEST / ("_".join(key) + ".xml")).relative_to(ROOT).as_posix()}
                if key in wanted:
                    old = wanted[key]
                    row["pending_role"] = old["role"]
                    row["header_matches_current_export"] = money == [amount(x) for x in old["header"]["money"]]
                    row["money_comparison"] = compare_xml_money(money, [amount(x) for x in old["header"]["money"]])
                    row["date_matches_current_export"] = invoice_day(row["date"]) == invoice_day(old["header"]["date"])
                    row["reference_matches_current_chain"] = old["role"] == "original" or reference == tuple(old["root"])
                records.append(row)
            except Exception as exc:
                errors.append({"entry": entry.filename, "error": str(exc)})
        assert not errors, errors
        assert records, "No XML read"
        DEST.mkdir(parents=True, exist_ok=True)
        preserve(DEST / args.archive.name, archive_data)
        for row in records:
            preserve(ROOT / row["file"], archive.read(row["entry"]))

    received_adjustments = {tuple(r["key"]) for r in records if r.get("pending_role") == "adjustment"}
    covered = originals | seen
    remaining = [k for k, row in wanted.items() if row["role"] == "adjustment" and k not in covered]
    summary = {"xml": len(records), "pdf": len(pdfs), "unique_invoices": len(seen),
               "date_range": [min(r["date"] for r in records), max(r["date"] for r in records)],
               "number_range": [min(k[1] for k in seen), max(k[1] for k in seen)],
               "pending_adjustments_in_this_archive": len(received_adjustments),
               "new_pending_adjustments": len(received_adjustments - originals),
               "remaining_pending_adjustment_xml_not_in_repo": len(remaining),
               "pending_exact_header_matches": sum(r.get("pending_role") is not None and r["money_comparison"] == "exact" for r in records),
               "pending_xml_zero_export_blank": [r["key"] for r in records if r.get("money_comparison") == "xml_zero_export_blank"],
               "pending_source_mismatches": [r["key"] for r in records if r.get("pending_role") and
                   (r["money_comparison"] not in ("exact", "xml_zero_export_blank") or not r["date_matches_current_export"] or not r["reference_matches_current_chain"])],
               "pending_arithmetic_differences_in_xml": [r["key"] for r in records if r.get("pending_role") and r["header_arithmetic_difference"] not in (None, 0)]}
    after = {path.name: digest(path.read_bytes()) for path in (ROOT / "outputs").glob("*.xlsx")}
    assert before == after and digest(args.archive.read_bytes()) == archive_sha
    result = {"review_date": "2026-10-11", "archive": {"original_path": str(args.archive.resolve()),
              "saved_file": (DEST / args.archive.name).relative_to(ROOT).as_posix(), "sha256": archive_sha,
              "bytes": len(archive_data)}, "summary": summary, "records": records, "pdf_entries": pdfs,
              "remaining_pending_adjustment_xml_not_in_repo": remaining, "workbooks_unchanged": True,
              "workbook_hashes": before, "errors": errors,
              "limits": ["XML contents read without cryptographic signature or live CQT verification.",
                         "All archive XML and the original XML/PDF ZIP are preserved byte for byte.",
                         "No source totals, workbook contents, import plan or issuance decision changed."]}
    report = REPORTS / ("pending-xml-" + archive_sha[:12] + ".json")
    report.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8")
    provenance = DEST / "PROVENANCE.json"
    history = json.loads(provenance.read_text(encoding="utf-8")) if provenance.exists() else {"archives": []}
    item = {**result["archive"], "xml": len(records), "pdf": len(pdfs),
            "audit_report": report.relative_to(ROOT).as_posix(),
            "xml_files": [{"path": r["file"], "sha256": r["sha256"], "archive_entry": r["entry"]} for r in records]}
    history["archives"] = [x for x in history["archives"] if x["sha256"] != archive_sha] + [item]
    provenance.write_text(json.dumps(history, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({**summary, "archive_sha256": archive_sha, "report": str(report),
                      "matched_pending": [r for r in records if r.get("pending_role")]}, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
