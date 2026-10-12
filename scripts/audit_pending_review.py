"""Recheck the 19 pending chains and 11 review chains without editing Excel."""
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile
import hashlib
import io
import json
import sys
import xml.etree.ElementTree as ET

import openpyxl

from prepare_invoice_reports import amount, invoice_day, invoice_key, rows

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx"
REPORT = ROOT / "scripts/data/check-review/pending-current-audit-20261011.json"
NEW = ROOT / "data/Du lieu Meinvoice/Xuat_lai_2026-10-09"


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def norm(value):
    return " ".join(str(value or "").split())


def first(element, name):
    return next((e for e in element.iter() if e.tag.rsplit("}", 1)[-1] == name), None)


def value(element, name):
    if element is None:
        return None
    return next((e.text for e in element if e.tag.rsplit("}", 1)[-1] == name), None)


def arithmetic(money):
    return None if any(x is None for x in money) else money[0] + money[1] - money[2]


def compare_xml_money(xml_money, export_money):
    """Record blank/zero differences without changing either source's values."""
    if xml_money == export_money:
        return "exact"
    if len(xml_money) == len(export_money) and all(
        a == b or (a == Decimal(0) and b is None)
        for a, b in zip(xml_money, export_money)
    ):
        return "xml_zero_export_blank"
    return "different"


def detail_totals(lines):
    result = []
    for field in ("goods", "tax", "payment"):
        values = [line[field] for line in lines]
        result.append(sum(values, Decimal(0)) if values and all(v is not None for v in values) else None)
    return result


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    before = sha(WORKBOOK)
    w = openpyxl.load_workbook(WORKBOOK, read_only=True, data_only=True)
    main_rows = {tuple(v[:2]): (rn, v) for rn, v in enumerate(
        w["Dieu chinh"].iter_rows(min_row=3, max_col=55, values_only=True), 3) if v[0]}
    pending_rows = [(rn, v) for rn, v in enumerate(
        w["Cho KT xac nhan"].iter_rows(min_row=3, max_col=12, values_only=True), 3) if v[0]]
    review_rows = [(rn, v) for rn, v in enumerate(
        w["42 - Can check 11"].iter_rows(min_row=4, max_col=27, values_only=True), 4) if v[0]]
    imported = {tuple(v) for v in w["Import dieu chinh"].iter_rows(
        min_row=10, min_col=10, max_col=11, values_only=True) if v[0]}
    w.close()
    assert len(pending_rows) == 19 and len(review_rows) == 11
    assert len({tuple(v[:2]) for _, v in pending_rows + review_rows}) == 30
    targets = set()
    for _, v in pending_rows + review_rows:
        root = tuple(v[:2])
        assert root not in imported, root
        for i in (0, 6, 12, 18):
            source = main_rows[root][1]
            if source[i]:
                targets.add(tuple(source[i:i + 2]))

    headers, details, edges, source_hashes = {}, defaultdict(list), set(), {}
    conflicts = []
    for path in sorted(NEW.iterdir()):
        if path.suffix not in (".xlsx", ".xls", ".zip") or "1791532943641" in path.name:
            continue
        data = path.read_bytes()
        source_hashes[path.relative_to(ROOT).as_posix()] = hashlib.sha256(data).hexdigest()
        if path.suffix == ".zip":
            with ZipFile(io.BytesIO(data)) as archive:
                books = [(n, archive.read(n)) for n in archive.namelist() if n.endswith((".xlsx", ".xls"))]
        else:
            books = [(path.name, data)]
        for name, content in books:
            label = path.relative_to(ROOT).as_posix() + (" / " + name if path.suffix == ".zip" else "")
            kind = "detail" if "chi_tiet" in name else (
                "link" if "dieu_chinh_xuat" in name or "thay_the_xuat" in name else "total")
            for rn, row in rows(io.BytesIO(content)):
                if amount(row.get("A")) is None:
                    continue
                if kind == "link":
                    parent = invoice_key(row.get("B"), row.get("D"))
                    child = invoice_key(row.get("M"), row.get("O"))
                    if parent in targets or child in targets:
                        edges.add((parent, child))
                    continue
                key = invoice_key(row.get("B"), row.get("C"))
                if key not in targets:
                    continue
                evidence = {"file": label, "row": rn}
                if kind == "total":
                    observation = {"date": invoice_day(row.get("D")),
                                   "money": [amount(row.get(c)) for c in ("P", "Q", "R")],
                                   "evidence": evidence}
                    if key in headers and any(headers[key][f] != observation[f] for f in ("date", "money")):
                        conflicts.append({"key": key, "old": headers[key], "new": observation})
                    headers[key] = observation
                else:
                    details[key].append({"name": row.get("L"), "unit": row.get("M"),
                                         "qty": amount(row.get("N")), "price": amount(row.get("O")),
                                         "goods": amount(row.get("P")), "rate": row.get("Q"),
                                         "tax": amount(row.get("R")), "payment": amount(row.get("S")),
                                         "evidence": evidence})
    print("Read current export sources for", len(targets), "target invoices", flush=True)

    xmls = defaultdict(list)
    for path in sorted((ROOT / "data/Bo sung/meInvoice").rglob("*.xml")):
        document = ET.fromstring(path.read_bytes())
        hd = first(document, "DLHDon")
        tt = first(hd, "TTChung")
        key = invoice_key((value(tt, "KHMSHDon") or "") + (value(tt, "KHHDon") or ""), value(tt, "SHDon"))
        if key not in targets:
            continue
        related, pay = first(tt, "TTHDLQuan"), first(hd, "TToan")
        reference = invoice_key((value(related, "KHMSHDCLQuan") or "") + (value(related, "KHHDCLQuan") or ""),
                                value(related, "SHDCLQuan")) if related is not None else None
        lines = []
        for line in hd.iter():
            if line.tag.rsplit("}", 1)[-1] != "HHDVu":
                continue
            lines.append({"name": value(line, "THHDVu"), "unit": value(line, "DVTinh"),
                          "qty": amount(value(line, "SLuong")), "price": amount(value(line, "DGia")),
                          "goods": amount(value(line, "ThTien")), "rate": value(line, "TSuat"),
                          "nature": value(line, "TChat")})
        xmls[key].append({"key": key, "date": invoice_day(value(tt, "NLap")), "related": reference,
                          "money": [amount(value(pay, f)) for f in ("TgTCThue", "TgTThue", "TgTTTBSo")],
                          "lines": lines, "file": path.relative_to(ROOT).as_posix(), "sha256": sha(path)})
        source_hashes[path.relative_to(ROOT).as_posix()] = sha(path)

    mismatches, blank_zero_differences, pending, review = [], [], [], []
    for rn, v in pending_rows:
        root = tuple(v[:2])
        main_rn, m = main_rows[root]
        if [amount(x) for x in v[3:6]] != [amount(x) for x in m[24:27]]:
            mismatches.append({"type": "pending_saved_balance", "root": root, "row": rn})
        observations = []
        for i in (0, 6, 12, 18):
            if not m[i]:
                continue
            key = tuple(m[i:i + 2])
            h = headers.get(key)
            saved = [amount(x) for x in m[i + 3:i + 6]]
            if h is None or h["money"] != saved or h["date"] != invoice_day(m[i + 2]):
                mismatches.append({"type": "pending_header", "key": key, "main_row": main_rn})
            if i and (root, key) not in edges:
                mismatches.append({"type": "pending_link", "root": root, "related": key})
            xchecks = [{"file": x["file"], "sha256": x["sha256"], "money": x["money"],
                        "header_arithmetic_difference": arithmetic(x["money"]),
                        "line_goods_sum": sum((l["goods"] for l in x["lines"] if l["goods"] is not None), Decimal(0)),
                        "header_matches_current": h is not None and x["money"] == h["money"] and x["date"] == h["date"],
                        "money_comparison": compare_xml_money(x["money"], h["money"]) if h else "missing_export",
                        "date_matches_current": h is not None and x["date"] == h["date"],
                        "reference_matches": i == 0 or x["related"] == root} for x in xmls.get(key, [])]
            for x in xchecks:
                if x["money_comparison"] == "xml_zero_export_blank":
                    blank_zero_differences.append({"key": key, "root": root, "xml": x["file"],
                                                   "xml_money": x["money"], "export_money": h["money"]})
                if x["money_comparison"] not in ("exact", "xml_zero_export_blank") or not x["date_matches_current"] or not x["reference_matches"]:
                    mismatches.append({"type": "pending_xml", "key": key, "xml": x["file"]})
            observations.append({"key": key, "role": "original" if i == 0 else "adjustment",
                                 "header": h, "header_arithmetic_difference": arithmetic(h["money"]) if h else None,
                                 "detail_line_count": len(details[key]), "detail_totals": detail_totals(details[key]),
                                 "tax_rates": sorted({str(l["rate"]) for l in details[key]}), "xml": xchecks})
        current_balance = []
        for column in range(3):
            values = [x["header"]["money"][column] if x["header"] else None for x in observations]
            # This one existing user-confirmed exception is diagnostic only.
            if root == ("1C24TUV", "00000574") and column == 2 and values[0] is None:
                values[0] = Decimal(0)
            current_balance.append(sum(values, Decimal(0)) if all(x is not None for x in values) else None)
        if current_balance != [amount(x) for x in v[3:6]]:
            mismatches.append({"type": "pending_current_source_balance", "root": root, "calculated": current_balance})
        zero_qty = [l for l in details[root] if l["qty"] in (None, Decimal(0))]
        if len(details[root]) != v[9] or len(zero_qty) != v[10]:
            mismatches.append({"type": "pending_original_detail_counts", "root": root})
        pending.append({"root": root, "pending_sheet_row": rn, "main_sheet_row": main_rn,
                        "saved_status": v[6], "saved_reason": v[7], "saved_cumulative_money": v[3:6],
                        "current_source_cumulative_money": current_balance,
                        "zero_or_blank_qty_lines": zero_qty, "invoices": observations})

    for rn, v in review_rows:
        root, key = tuple(v[:2]), tuple(v[3:5])
        original, adjustment = headers.get(root), headers.get(key)
        original_detail = details[root]
        if len(original_detail) != 1:
            mismatches.append({"type": "review_original_line_count", "root": root, "count": len(original_detail)})
        else:
            line = original_detail[0]
            expected = (norm(v[6]), norm(v[8]), amount(v[10]), amount(v[12]))
            actual = (norm(line["name"]), norm(line["unit"]), line["qty"], line["price"])
            if expected != actual:
                mismatches.append({"type": "review_original_detail", "root": root})
        if original is None or original["money"] != [amount(v[i]) for i in (14, 16, 18)] or original["date"] != invoice_day(v[2]):
            mismatches.append({"type": "review_original_header", "root": root})
        if adjustment is None or adjustment["money"] != [amount(v[i]) for i in (15, 17, 19)] or adjustment["date"] != invoice_day(v[5]):
            mismatches.append({"type": "review_adjustment_header", "key": key})
        if (root, key) not in edges:
            mismatches.append({"type": "review_link", "root": root, "key": key})
        xs = xmls.get(key, [])
        if not xs:
            mismatches.append({"type": "review_missing_xml", "key": key})
        for x in xs:
            if len(x["lines"]) != 1:
                mismatches.append({"type": "review_xml_line_count", "key": key})
                continue
            line = x["lines"][0]
            expected = (norm(v[7]), norm(v[9]), amount(v[11]), amount(v[13]), amount(v[15]))
            actual = (norm(line["name"]), norm(line["unit"]), line["qty"], line["price"], line["goods"])
            rate_matches = amount(str(line["rate"] or "").removesuffix("%")) == amount(v[20])
            if expected != actual or not rate_matches or x["related"] != root or x["date"] != invoice_day(v[5]) or x["money"] != [amount(v[i]) for i in (15, 17, 19)]:
                mismatches.append({"type": "review_xml", "key": key, "file": x["file"]})
        balance = [a + b for a, b in zip(original["money"], adjustment["money"])] if original and adjustment else None
        if balance != [Decimal(0)] * 3 or list(v[22:25]) != [0, 0, 0]:
            mismatches.append({"type": "review_balance", "root": root})
        review.append({"root": root, "adjustment": key, "sheet_row": rn, "category": v[21],
                       "saved_action": v[25], "original_header": original, "adjustment_header": adjustment,
                       "original_detail": original_detail, "xml": xs, "cumulative_money": balance})

    all_pending = [x for chain in pending for x in chain["invoices"]]
    adjustments = [x for x in all_pending if x["role"] == "adjustment"]
    summary = {"pending_chains": len(pending), "pending_invoices": len(all_pending),
               "pending_adjustments": len(adjustments), "pending_originals_with_xml": sum(bool(x["xml"]) for x in all_pending if x["role"] == "original"),
               "pending_adjustments_with_xml": sum(bool(x["xml"]) for x in adjustments),
               "pending_invoices_with_blank_header_money": sum(any(v is None for v in x["header"]["money"]) for x in all_pending if x["header"]),
               "pending_invoices_with_arithmetic_difference": sum(x["header_arithmetic_difference"] not in (None, 0) for x in all_pending),
               "pending_roots_with_zero_or_blank_qty": sum(bool(x["zero_or_blank_qty_lines"]) for x in pending),
               "review_chains": len(review), "review_categories": dict(Counter(x["category"] for x in review)),
               "review_adjustments_with_xml": sum(bool(x["xml"]) for x in review),
               "review_all_zero": all(x["cumulative_money"] == [Decimal(0)] * 3 for x in review),
               "all_30_roots_excluded_from_import": True, "new_mismatch_count": len(mismatches),
               "xml_zero_export_blank_invoices": len({tuple(x["key"]) for x in blank_zero_differences}),
               "source_conflict_count": len(conflicts), "workbook_unchanged": before == sha(WORKBOOK)}
    unchanged = all(sha(ROOT / path) == expected for path, expected in source_hashes.items())
    receipts = json.loads((ROOT / "scripts/data/check-review/zip-237-xml-review-20261009.json").read_text(encoding="utf-8"))
    result = {"review_date": "2026-10-11", "read_only": True, "summary": summary,
              "xml_availability_scope": "Extracted XML under data/Bo sung/meInvoice only; this does not establish whether an invoice was supplied in the original bulk archives.",
              "historical_bulk_xml_receipts": receipts["inventory"],
              "passed": not mismatches and not conflicts and summary["workbook_unchanged"] and unchanged,
              "workbook_sha256": before, "source_files_unchanged": unchanged, "source_hashes": source_hashes,
              "pending": pending, "review_11": review, "new_mismatches": mismatches, "source_conflicts": conflicts,
              "xml_zero_export_blank_differences": blank_zero_differences,
              "pending_adjustment_xml_not_in_repo": [x["key"] for x in adjustments if not x["xml"]],
              "limits": ["PASS checks consistency with current files, not resolution of existing pending issues.",
                         "Two historical bulk ZIPs contained 1,436 XML; only the 237 target XML were extracted by audit_downloaded_xml.py. Recheck those original ZIPs before requesting invoices again.",
                         "Blank money and zero remain distinct; XML-zero/export-blank differences are explicitly listed, not treated as exact matches; source and workbook values are not changed.",
                         "No accounting decision, cryptographic signature verification, live CQT check, import or issuance."]}
    REPORT.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8")
    print(json.dumps({"passed": result["passed"], **summary, "new_mismatches": mismatches}, ensure_ascii=False, indent=2))
    if not result["passed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
