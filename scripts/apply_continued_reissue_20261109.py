"""Stage paired drafts from the current delivered workbooks and verified sources."""

import argparse
import csv
import hashlib
import json
import shutil
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.datetime import to_excel

from import_2025_buyer_fields import apply_buyer_fields
from prepare_invoice_reports import invoice_day, invoice_key, rows

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs/import-2025-20261109"
LEDGER = "DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv"
LINKS = "LIEN_KET_IMPORT_MOI.csv"
NEW, REVIEW, KEEP = "Tạo import mới", "Chờ xác minh", "Giữ nguyên, không import"
SHARED = {"BHHT2025027442": "BHHT2025028267", "BHHT2025027439": "BHHT2025028252"}
PSC2 = {"BHHT2025034835": 66000, "BHHT2025032283": 1160000, "BHHS2025018941": 237000}


def read_json(name):
    return json.loads((OUT / name).read_text(encoding="utf-8"))


def read_csv(name):
    with (OUT / name).open(encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file)
        return next(reader), list(reader)


def write_csv(path, headers, data):
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        writer.writerows(data)


def write_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def number(value):
    return Decimal(str(value or 0))


def money_round(value):
    return value.quantize(Decimal(1), rounding=ROUND_HALF_UP)


def excel_cell(value):
    if isinstance(value, (date, datetime)):
        return int(to_excel(value))
    return value


def load_delivered_payload():
    payload = read_json("authoring.json")
    payload["summary"] = read_json("final-summary.json")
    proof = read_json("verification.json")
    assert proof["passed"]
    checks = {entry["file"]: entry for entry in proof["exported_files"]}
    payload["manifest"] = proof["source_hashes"]
    for output in payload["outputs"]:
        path = OUT / "deliverables" / output["file"]
        assert digest(path) == checks[path.name]["sha256"], path.name
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            tables = output["sheets"] if output["kind"] == "report" else [output]
            for table in tables:
                name = table.get("name")
                sheet = workbook[name] if name else workbook.active
                start = table.get("start", 1) if name else 9
                limit = 13 if name == "Tong hop" else None
                data = sheet.iter_rows(min_row=start, max_row=limit,
                                       max_col=len(table["headers"]), values_only=True)
                assert list(next(data)) == table["headers"], name
                table["rows"] = [[excel_cell(cell) for cell in row] for row in data if row[0] is not None]
            if output["kind"] != "report":
                output["invoices"] = checks[path.name]["invoices"]
                assert len(output["rows"]) == checks[path.name]["rows"]
                output["notes"] = [row[0] for row in workbook.active.iter_rows(
                    min_row=2, max_row=5, max_col=1, values_only=True) if row[0]]
        finally:
            workbook.close()
    return payload


def read_replacement_sources(keys, ledger):
    wanted = defaultdict(dict)
    for key in keys:
        source = ledger[key]["source"]
        wanted[source["file"]][source["row"]] = key
    found = {}
    for file, row_keys in wanted.items():
        for row_number, row in rows(ROOT / file):
            if row_number not in row_keys:
                continue
            key = row_keys[row_number]
            assert "/".join(invoice_key(row.get("B"), row.get("C"))) == key
            found[key] = row
    assert found.keys() == set(keys)
    return found


def apply(stage):
    prepared = read_json("prepared.json")
    source_ledger = read_json("source-index.json")["ledger"]
    payload = load_delivered_payload()
    outputs = {item["kind"]: item for item in payload["outputs"]}
    summary = payload["summary"]
    sales, repl = outputs["new"], outputs["replacement"]
    sheets = {tab["name"]: tab for tab in outputs["report"]["sheets"]}
    headers, csv_rows = read_csv(LEDGER)
    decisions = {row[0]: row for row in csv_rows}
    candidate = [row for row in csv_rows if row[2] == REVIEW and "Chuỗi thay thế phức tạp" in row[10]]
    assert len(candidate) == 13
    selected = [row for row in candidate if row[0] not in SHARED]
    active = {row[0]: prepared["chains"][row[8]]["active"][0] for row in selected}
    sources = read_replacement_sources(active.values(), source_ledger)
    date_serial = sales["rows"][0][1]
    old_sales_rows, old_repl_rows = len(sales["rows"]), len(repl["rows"])
    old_sales_count, old_repl_count = sales["invoices"], repl["invoices"]
    assert (old_sales_rows, old_repl_rows, old_sales_count) == (35024, 1980, 30248)
    doc_roots = {code: [r for r in csv_rows if root in r[8].split("; ")]
                 for code, root in [(r[0], r[8]) for r in selected]}
    excluded = dict(summary.get("psc2_exchanges", {}))
    records, new_rows, replacement_rows, link_additions = [], [], [], []
    sequences = {row[18]: row[0] for row in sales["rows"]}
    returns = {row[25] for row in outputs["return_adjustment"]["rows"]}
    company_roots = {row["key"] for row in prepared["company_rows"]}
    replacement_roots = {row[25] for row in repl["rows"]}
    for decision in selected:
        code, root = decision[0], decision[8]
        reference = active[code]
        chain = prepared["chains"][root]
        docs = [code]
        assert {r[0] for r in doc_roots[code]} == set(docs)
        assert root not in replacement_roots and not set(chain["keys"]) & company_roots
        assert len(chain["keys"]) == 2 and chain["active"] == [reference] and not chain["issues"]
        assert source_ledger[reference]["status"] == "Hoá đơn thay thế"
        identity = sources[reference]
        before = [int(number(identity.get(field))) for field in "PQR"]
        assert before == chain["balance"] and before[0] + before[1] == before[2]
        assert int(decision[5]) - before[2] == int(decision[7])
        assert abs(money_round(number(before[0]) * Decimal("0.08")) - before[1]) <= 2
        day = invoice_day(identity.get("D"))
        assert day and identity.get("E") and (identity.get("H") or identity.get("L"))
        order = str(prepared["docs"][code]["order"] or "")
        known_order = source_ledger[reference]["order"]
        assert known_order == order or (len(order) == len(known_order) == 18
            and known_order.endswith("000") and known_order[:15] == order[:15])
        sig, invoice_no = reference.split("/")
        zero_seq = old_repl_count + len(replacement_rows) + 1
        reason = f"Thay thế tiếp HĐ {sig}/{invoice_no} ngày {day:%d/%m/%Y} về 0. Lập lại theo {', '.join(docs)}."
        replacement_rows.append([
            zero_seq, date_serial, identity.get("H") or None, identity.get("J") or None,
            identity.get("I") or None, identity.get("L") or None, None, None, None,
            "TM/CK", None, sig, invoice_no, int(to_excel(day)), identity["E"],
            reason, None, 0, 0, 0, 8, 0, reason, 1, ", ".join(docs), root,
        ])
        new_sequences, new_totals = {}, {}
        for doc_code in docs:
            assert doc_code not in sequences and doc_code not in returns and doc_code != "BHHT2025001488"
            doc = prepared["docs"][doc_code]
            lines = doc["lines"]
            excluded_line = [line for line in lines if line["v"][15] == "PSC2"]
            goods, vat, pay = [int(doc[k]) for k in ("goods", "vat", "pay")]
            if doc_code in PSC2:
                assert len(excluded_line) == 1
                fee = excluded_line[0]["v"]
                assert (fee[24], fee[41], fee[43]) == (PSC2[doc_code], 0, PSC2[doc_code])
                goods -= fee[24]
                lines = [line for line in lines if line not in excluded_line]
                source_vat = vat
                assert sum(line["v"][41] for line in lines) == source_vat
                vat = int(money_round(number(goods) * Decimal("0.08")))
                pay = goods + vat
                excluded[doc_code] = {
                    "excluded_goods": fee[24], "goods": goods, "vat": vat, "payment": pay,
                    "rate": 8, "product_sources": [line["source"] for line in lines],
                    "psc2_source": excluded_line[0]["source"], "old_root": root,
                    "rate_policy": "recalculate", "source_vat": source_vat,
                }
            else:
                assert not excluded_line and abs(money_round(number(goods) * Decimal("0.08")) - vat) <= 2
            assert goods + vat == pay and sum(line["v"][24] for line in lines) == goods
            assert doc_code in PSC2 or sum(line["v"][41] for line in lines) == vat
            seq = len(sequences) + 1
            sequences[doc_code] = seq
            new_sequences[doc_code] = seq
            new_totals[doc_code] = [goods, vat, pay]
            first = 10 + old_sales_rows + len(new_rows)
            for i, line in enumerate(lines):
                item = line["v"]
                assert abs(money_round(number(item[22]) * number(item[23])) - number(item[24])) <= 2
                new_rows.append([
                    seq, date_serial, doc["buyer"], item[13] or None, item[12] or None,
                    item[14] or None, None, None, None, "TM/CK", 8, vat if i == 0 else None,
                    str(item[7] or item[16]).strip(), item[19], item[22], item[23],
                    item[24], doc["order"], doc_code, item[15], doc["date"], line["source"], None,
                ])
            link_additions.append([sales["file"], seq, doc_code, first, len(lines), root])
            decisions[doc_code][2] = NEW
            decisions[doc_code][-1] += f". Thay thế tiếp HĐ cuối {reference} về 0, dự thảo thay thế STT {zero_seq}, bán mới STT {seq}."
            if doc_code in PSC2:
                decisions[doc_code][-1] += f" Trừ riêng PSC2 {PSC2[doc_code]}đ, tính lại VAT 8% là {vat}đ."
        records.append({
            "root": root, "active": reference, "keys": chain["keys"], "docs": docs,
            "before": before, "zero_seq": zero_seq, "new_sequences": new_sequences,
            "new_totals": new_totals, "identity_source": source_ledger[reference]["source"],
        })
    apply_buyer_fields({"outputs": [{**sales, "rows": new_rows}, {**repl, "rows": replacement_rows}],
                        "manifest": []}, OUT, ROOT)
    sales["rows"].extend(new_rows)
    repl["rows"].extend(replacement_rows)
    sales["invoices"] = len(sequences)
    repl["invoices"] = len(repl["rows"])
    sales["notes"][0] = f"Gộp toàn bộ bán mới năm 2025. STT liên tục 1–{sales['invoices']:,}.".replace(",", ".")
    sales["notes"][2] = "Ba chứng từ bổ sung đã trừ PSC2 và tính lại VAT 8%. Hai chứng từ PSC2 đã chốt trước giữ VAT AMIS. Phiếu trả hàng lập riêng."
    sales["notes"].append("BHHT2025027502 chờ KT chốt chênh thanh toán 723.000đ trước khi dùng.")
    old_overrides = summary["rounding_reissue_20261109"]["records"]
    repl["notes"][0] = f"Tham chiếu HĐ còn hiệu lực. Có {len(records) + sum(x['root'] != x['active'] for x in old_overrides)} HĐ thay thế cũ cần thay thế tiếp về 0."

    new_codes = {code for record in records for code in record["docs"]}
    sheets["Cho xac minh"]["rows"] = [row for row in sheets["Cho xac minh"]["rows"] if row[0] not in new_codes]
    report_rows = {row[0]: row for row in sheets["Xu ly chuoi cu"]["rows"]}
    for record in records:
        row = report_rows[record["root"]]
        row[1:3] = ["Thay thế tiếp về 0", ", ".join(record["docs"])]
        row[3:6] = record["before"]
        row[7] = f"Đã tạo dự thảo thay thế STT {record['zero_seq']} cho HĐ cuối {record['active']}."

    counts = Counter((int(row[1][5:7]), row[2]) for row in csv_rows)
    monthly_pay = defaultdict(int)
    for row in sales["rows"]:
        monthly_pay[int(row[20][5:7])] += int(number(row[16]) + number(row[11]))
    for row in sheets["Tong hop"]["rows"]:
        month = row[0]
        row[2:6] = [counts[month, label] for label in (NEW, KEEP, "Loại theo HD CTY 2025", REVIEW)]
        row[6] = monthly_pay[month]
        assert sum(row[2:6]) == row[1]
    actions = Counter(row[2] for row in csv_rows)
    summary["actions"].update(keep=actions[KEEP], new=actions[NEW], review=actions[REVIEW])
    summary["new_invoices"], summary["new_lines"] = sales["invoices"], len(sales["rows"])
    summary["replacement"] = repl["invoices"]
    summary["new_goods"] = sum(int(number(row[16])) for row in sales["rows"])
    summary["new_vat"] = sum(int(number(row[11])) for row in sales["rows"])
    summary["new_pay"] = summary["new_goods"] + summary["new_vat"]
    summary["psc2_exchanges"] = excluded
    summary["continued_reissue_20261109"] = {
        "records": records, "remaining": [row[0] for row in candidate if row not in selected],
        "base_commit": "9d11749", "psc2_excluded": True, "shared_reissued": False,
    }
    assert len(sheets["Cho xac minh"]["rows"]) == summary["actions"]["review"]
    assert len(set(sequences.values())) == summary["new_invoices"]

    # The ignored authoring cache predates the pulled files. Rebuild its chain
    # plan from the delivered references and their recorded source evidence.
    source_audit = {z["root"]: z for z in payload["zero_audit"] + payload["unmapped_zero_audit"]}
    overrides = {z["root"]: z for z in old_overrides + records}
    old_unmapped = {z["root"] for z in payload["unmapped_zero_audit"]}
    adjustments = [z for z in source_audit.values() if z["kind"] == "adjustment"]
    payload["zero_audit"] = [z for z in adjustments if z["root"] not in old_unmapped]
    payload["unmapped_zero_audit"] = [z for z in adjustments if z["root"] in old_unmapped]
    for row in repl["rows"]:
        root = row[25]
        assert root in source_audit or root in overrides, root
        audit = dict(source_audit.get(root) or {})
        override = overrides.get(root)
        if override:
            assert f"{row[11]}/{row[12]}" == override["active"]
            audit.update(before=override["before"], keys=override["keys"],
                         docs=override["docs"] if "docs" in override else [override["doc"]])
        elif any(decisions[code][2] != NEW for code in audit.get("docs", [])):
            current_docs = [code.strip() for code in row[24].replace(";", ",").split(",")]
            assert set(current_docs) <= set(audit["docs"])
            assert all(decisions[code][2] == NEW for code in current_docs)
            audit["docs"] = current_docs
        audit.update(root=root, kind="replacement", sequence=row[0],
                     change=[-v for v in audit["before"]], after=[0, 0, 0])
        target = "unmapped_zero_audit" if root in old_unmapped and not override else "zero_audit"
        payload[target].append(audit)
    assert len(payload["zero_audit"] + payload["unmapped_zero_audit"]) == 18 + repl["invoices"]

    stage.mkdir(parents=True, exist_ok=True)
    (stage / "deliverables").mkdir(exist_ok=True)
    (stage / "qa").mkdir(exist_ok=True)
    for output in (outputs["adjustment"], outputs["return_adjustment"]):
        shutil.copy2(OUT / "deliverables" / output["file"], stage / "deliverables" / output["file"])
    for name in ("prepared.json", "return-verification.json", "rounding-confirmation.json",
                 "service-quantity-confirmations.json", "unmapped-import-verification.json",
                 "adjustment-rule-2024-verification.json", "unmapped-gift-confirmation.json",
                 "LIEN_KET_BTL_BHH_2025.csv"):
        if not (stage / name).exists():
            shutil.copy2(OUT / name, stage / name)
    write_csv(stage / LEDGER, headers, csv_rows)
    link_headers, links = read_csv(LINKS)
    assert {row[2] for row in links}.isdisjoint(new_codes)
    write_csv(stage / LINKS, link_headers, links + link_additions)
    write_json(stage / "authoring.json", payload)
    write_json(stage / "final-summary.json", summary)
    write_json(stage / "invoice-sequences.json", sequences)
    write_json(stage / "native-additions.json", {
        "sale": new_rows, "replacement": replacement_rows, "records": records,
        "counts": {"new": sales["invoices"], "lines": len(sales["rows"]),
                   "replacement": repl["invoices"], "review": actions[REVIEW]},
        "pending_goods": sum(int(row[3]) for row in csv_rows if row[2] == REVIEW),
        "pending_vat": sum(int(row[4]) for row in csv_rows if row[2] == REVIEW),
        "new_goods": summary["new_goods"], "new_vat": summary["new_vat"],
        "psc2": sum(value["excluded_goods"] for value in excluded.values()),
        "psc2_vat_reduction": sum(value.get("source_vat", value["vat"]) - value["vat"]
                                   for value in excluded.values()),
    })
    print(json.dumps({"new": sales["invoices"], "new_lines": len(sales["rows"]),
                      "replacement": repl["invoices"], "review": actions[REVIEW],
                      "paired_roots": len(records), "added_new": len(new_codes)}))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    args = parser.parse_args()
    stage = args.stage.resolve()
    assert stage.is_relative_to((OUT / "_build").resolve())
    apply(stage)
