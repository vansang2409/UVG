"""Stage the confirmed BHDV20250074 invoice link and refresh the report."""

import argparse
import shutil
from pathlib import Path

from apply_continued_reissue_20261109 import (
    KEEP, LEDGER, LINKS, OUT, load_delivered_payload, read_csv, write_csv, write_json,
)


DOC = "BHDV20250074"
INVOICE = "1C25MUV/00020784"
OTHER_DOC = "BHDV20250066"
OTHER_INVOICE = "1C25MUV/00024410"
MONEY = [138889, 11111, 150000]
NOTE = (f"Kế toán chốt giữ HĐ {INVOICE} cho {DOC}. Khớp ngày, người mua và số tiền. "
        "Mã đơn AMIS HD009935, mã đơn trên HĐ meInvoice HD009798.")


def stage_update(stage):
    payload = load_delivered_payload()
    summary = payload["summary"]
    outputs = {entry["kind"]: entry for entry in payload["outputs"]}
    tabs = {tab["name"]: tab["rows"] for tab in outputs["report"]["sheets"]}
    keys = {INVOICE, OTHER_INVOICE}
    assert not any(keys & set(audit.get("keys", [audit["root"]]))
                   for audit in payload["zero_audit"] + payload["unmapped_zero_audit"])
    assert not keys & {row[25] for row in outputs["replacement"]["rows"]}
    assert not {DOC, OTHER_DOC} & {row[18] for row in outputs["new"]["rows"]}
    assert not {DOC, OTHER_DOC} & {row[25] for row in outputs["return_adjustment"]["rows"]}

    headers, ledger = read_csv(LEDGER)
    records = {row[0]: row for row in ledger if row[0] in {DOC, OTHER_DOC}}
    for doc, row in records.items():
        assert row[2] == KEEP and row[8] == OTHER_INVOICE
        assert list(map(int, row[3:6])) == MONEY and int(row[7]) == 0
    assert records[DOC][6] == "13420" and records[OTHER_DOC][6] == "9865"
    records[DOC][8], records[DOC][-1] = INVOICE, NOTE
    records[OTHER_DOC][-1] = f"Giữ HĐ {OTHER_INVOICE} cho {OTHER_DOC}, khớp mã đơn HD009935."
    row, = [row for row in tabs["MI2025 chua map"] if row[0] == INVOICE]
    assert row[1:3] == ["Chờ KT — tạm rút dự thảo", None] and row[4:7] == MONEY and row[10] is None
    row[1], row[9] = "Đã nối AMIS và giữ HĐ", NOTE
    assert not keys & {row[0] for row in tabs["Xu ly chuoi cu"]}
    tabs["Xu ly chuoi cu"].extend([
        [INVOICE, "Giữ HĐ", DOC, *MONEY, INVOICE, NOTE],
        [OTHER_INVOICE, "Giữ HĐ", OTHER_DOC, *MONEY, OTHER_INVOICE, records[OTHER_DOC][-1]],
    ])

    reporting = summary["amis_reporting_0911"]
    assert reporting["gap"] == [-138889, -136870]
    reporting["projection"] = [a + b for a, b in zip(reporting["projection"], MONEY[:2])]
    reporting["gap"] = [a - b for a, b in zip(reporting["projection"], reporting["scope_current"])]
    assert reporting["gap"] == [0, reporting["vat_policy_delta"]] == [0, -125759]
    summary["confirmed_service_keep_20261109"] = {
        "doc": DOC, "invoice": INVOICE, "amis_row": 13420, "money": MONEY,
        "amis_order": "HD009935", "invoice_order": "HD009798", "note": NOTE,
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


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    stage = parser.parse_args().stage.resolve()
    assert stage.is_relative_to((OUT / "_build").resolve())
    stage_update(stage)
