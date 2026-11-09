"""Refresh the AMIS reporting baseline from the 09/11 exports."""

import argparse
import shutil
from collections import defaultdict
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from apply_continued_reissue_20261109 import (
    LEDGER, LINKS, NEW, OUT, ROOT, digest, load_delivered_payload,
    read_csv, read_json, write_csv, write_json,
)


LATEST = ROOT / "data/0911"
SERVICES = {"BHDV20250104", "BHDV20250115"}
DELETED = "BHDV20250045"


def load_sales(path):
    book = load_workbook(path, read_only=True, data_only=True)
    data = {}
    try:
        for sheet in book:
            for i, row in enumerate(sheet.values, 1):
                if len(row) < 8 or not isinstance(row[1], date) or not isinstance(row[2], str) or not row[2]:
                    continue
                code, vat, payment = row[2], row[6], row[7]
                assert code not in data and type(vat) is int and type(payment) is int, (path, code)
                data[code] = {"money": [payment - vat, vat, payment], "row": i, "sheet": sheet.title}
    finally:
        book.close()
    return data


def validate_detail(path, sales, exchanges):
    book = load_workbook(path, read_only=True, data_only=True)
    sums = defaultdict(lambda: [0, 0, 0])
    service_rows = set()
    try:
        for sheet in book:
            for row in sheet.values:
                if len(row) < 44 or not isinstance(row[0], date) or not isinstance(row[2], str):
                    continue
                assert row[15] != "PSC2" and row[2] != DELETED
                if row[2] in exchanges:
                    for j, k in enumerate((24, 41, 43)):
                        assert type(row[k]) is int
                        sums[row[2]][j] += row[k]
                if row[2] in SERVICES:
                    assert row[15] == "PSC1" and row[20] == row[22] == 2
                    service_rows.add(row[2])
    finally:
        book.close()
    assert service_rows == SERVICES and sums.keys() == exchanges.keys()
    assert all(v == sales[code]["money"] for code, v in sums.items())


def stage_update(stage):
    baseline = read_json(stage / "baseline.json")
    payload = load_delivered_payload()
    summary = payload["summary"]
    assert "amis_reporting_0911" not in summary
    exchanges = summary["psc2_exchanges"]
    new_path = LATEST / "Ban_hang.xlsx"
    old = load_sales(ROOT / "data/0811/Ban_hang.xlsx")
    new = load_sales(new_path)
    assert old.keys() - new.keys() == {DELETED} and not new.keys() - old.keys()
    assert old[DELETED]["money"] == [0, 0, 0]
    assert {code for code in new if old[code]["money"] != new[code]["money"]} == exchanges.keys()
    detail_path = LATEST / "So_chi_tiet_ban_hang.xlsx"
    validate_detail(detail_path, new, exchanges)

    headers, ledger = read_csv(LEDGER)
    assert {r[0] for r in ledger if r[2] != "Đã xóa trên AMIS"} == new.keys()
    assert len(new) == summary["total"] == 65492
    seqs = read_json("invoice-sequences.json")
    zero_by_doc = {code: p for p in payload["zero_audit"] for code in p["docs"]}
    adjustment_by_root = {p["root"]: p for p in payload["adjustment_rule_2024"]["chains"]}
    changed = {}
    for row in ledger:
        code = row[0]
        if code not in exchanges:
            continue
        before, after = old[code]["money"], new[code]["money"]
        e = exchanges[code]
        assert list(map(int, row[3:6])) == before and row[2] == NEW
        assert before[0] - after[0] == before[2] - after[2] == e["excluded_goods"] and before[1] == after[1]
        assert after[0] == e["goods"]
        if e.get("rate_policy") == "recalculate":
            assert e["vat"] < after[1]
        else:
            assert e["vat"] == after[1]
        row[3:6] = list(map(str, after))
        audit = zero_by_doc.get(code)
        action = ""
        if audit and audit["kind"] == "replacement":
            action = f" Thay thế về 0 STT {audit['sequence']}."
        if audit and audit["kind"] == "adjustment":
            seq = adjustment_by_root[audit["root"]]["sequences"]
            action = f" Điều chỉnh về 0 STT {', '.join(map(str, seq))}."
        row[-1] = (f"AMIS 09/11 đã bỏ dòng PSC2. STT bán mới {seqs[code]}. "
                   f"VAT AMIS {after[1]}đ, VAT dự thảo {e['vat']}đ.{action}")
        changed[code] = {**new[code], "previous": before, "draft": [e["goods"], e["vat"], e["payment"]]}
    assert changed.keys() == exchanges.keys()
    goods_delta = sum(rec["money"][0] - rec["previous"][0] for rec in changed.values())
    vat_delta = sum(rec["money"][1] - rec["previous"][1] for rec in changed.values())
    assert (goods_delta, vat_delta) == (-2351000, 0)
    scope = [v + d for v, d in zip(baseline["scope_before"], (goods_delta, vat_delta))]
    projection = baseline["projection"]
    gap = [a - b for a, b in zip(projection, scope)]
    assert gap == [-4871571, -515483]
    vat_policy_delta = sum(rec["draft"][1] - rec["money"][1] for rec in changed.values())
    assert vat_policy_delta == -125759

    sources = [{"file": p.relative_to(ROOT).as_posix(), "sha256": digest(p)} for p in (new_path, detail_path)]
    payload["manifest"].extend(sources)
    summary["amis_reporting_0911"] = {
        "source_date": "09/11/2026", "source_files": sources, "documents": len(new),
        "source_gross": [sum(r["money"][i] for r in new.values()) for i in range(3)],
        "scope_previous": baseline["scope_before"], "scope_current": scope,
        "projection": projection, "gap": gap, "psc2_documents": changed,
        "vat_policy_delta": vat_policy_delta,
    }

    (stage / "deliverables").mkdir(parents=True, exist_ok=True)
    for entry in payload["outputs"]:
        if entry["kind"] != "report":
            name = entry["file"]
            shutil.copy2(OUT / "deliverables" / name, stage / "deliverables" / name)
    for name in ("prepared.json", "return-verification.json", "rounding-confirmation.json",
                 "service-quantity-confirmations.json", "unmapped-import-verification.json",
                 "adjustment-rule-2024-verification.json", "unmapped-gift-confirmation.json",
                 "LIEN_KET_BTL_BHH_2025.csv", LINKS):
        shutil.copy2(OUT / name, stage / name)
    write_csv(stage / LEDGER, headers, ledger)
    write_json(stage / "authoring.json", payload)
    write_json(stage / "final-summary.json", summary)
    write_json(stage / "report-refresh.json", summary["amis_reporting_0911"])
    print(f"AMIS reporting refreshed from {len(new)} vouchers; scope {scope}, gap {gap}.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    args = parser.parse_args()
    stage = args.stage.resolve()
    assert stage.is_relative_to((OUT / "_build").resolve())
    stage_update(stage)
