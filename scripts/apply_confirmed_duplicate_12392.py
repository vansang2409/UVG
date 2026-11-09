"""Stage replacement of the duplicate 00012392 while keeping 00012263."""

import argparse
import shutil
from datetime import date
from pathlib import Path

from openpyxl.utils.datetime import to_excel

from apply_continued_reissue_20261109 import (
    KEEP, LEDGER, LINKS, OUT, ROOT, invoice_day, load_delivered_payload,
    read_csv, read_json, read_replacement_sources, write_csv, write_json,
)
from import_2025_buyer_fields import apply_buyer_fields


DUPLICATE = "1C25MUV/00012392"
KEPT = "1C25MUV/00012263"
DOC = "BHHT2025027643"
ORDER = "580325498726155292"
MONEY = [1616713, 129337, 1746050]
REASON = f"Thay thế HĐ {DUPLICATE} về 0 do xuất trùng. Giữ HĐ {KEPT} cho {DOC}."


def stage_update(stage):
    payload = load_delivered_payload()
    summary = payload["summary"]
    prior = summary["confirmed_keep_12515_20261109"]
    assert prior["detached_invoice"] == DUPLICATE
    prepared = read_json("prepared.json")
    source = read_json("source-index.json")
    keys = {DUPLICATE, KEPT}
    identity = read_replacement_sources(keys, source["ledger"])
    original, kept = identity[DUPLICATE], identity[KEPT]
    assert original["L"] and original["L"] == kept["L"]
    assert original["H"] == kept["H"] and invoice_day(original["D"]) == invoice_day(kept["D"])
    for key, invoice in identity.items():
        assert source["ledger"][key]["status"] == "Hoá đơn mới" and invoice["E"]
        assert [int(invoice[f]) for f in "PQR"] == MONEY
        assert not any(source[k].get(key) for k in ("children", "parents", "replacement"))
    a, = prepared["invoice_detail"][DUPLICATE]
    b, = prepared["invoice_detail"][KEPT]
    assert {k: v for k, v in a.items() if k != "source"} == {k: v for k, v in b.items() if k != "source"}
    assert a["qty"] == "1" and "KA350 Light" in a["name"] and a["rate"] == "8%"
    assert prepared["docs"][DOC]["order"] == ORDER
    assert not keys & {r["key"] for r in prepared["company_rows"]}
    assert not any(keys & set(z.get("keys", [z["root"]]))
                   for z in payload["zero_audit"] + payload["unmapped_zero_audit"])

    headers, ledger = read_csv(LEDGER)
    decision, = [r for r in ledger if r[0] == DOC]
    assert decision[2] == KEEP and decision[8] == KEPT and list(map(int, decision[3:6])) == MONEY
    assert DOC not in read_json("invoice-sequences.json")
    outputs = {o["kind"]: o for o in payload["outputs"]}
    replacement = outputs["replacement"]
    seq = replacement["invoices"] + 1
    assert seq == 1990 and len(replacement["rows"]) == seq - 1
    symbol, number = DUPLICATE.split("/")
    draft = [
        seq, int(to_excel(date(2026, 11, 9))), original.get("H") or None, original.get("J") or None,
        original.get("I") or None, original.get("L") or None, None, None, None, "TM/CK", None,
        symbol, number, int(to_excel(invoice_day(original["D"]))), original["E"], REASON,
        None, 0, 0, 0, 8, 0, REASON, 1, DOC, DUPLICATE,
    ]
    apply_buyer_fields({"outputs": [{**replacement, "rows": [draft]}], "manifest": []}, OUT, ROOT)
    replacement["rows"].append(draft)
    replacement["invoices"] = summary["replacement"] = seq
    note = f"Giữ HĐ {KEPT} cho {DOC}. Thay thế HĐ trùng {DUPLICATE} về 0, STT {seq}."
    decision[-1] += f". {note}"
    tabs = {tab["name"]: tab["rows"] for tab in outputs["report"]["sheets"]}
    old, = [r for r in tabs["Xu ly chuoi cu"] if r[0] == DUPLICATE]
    assert not old[2] and old[3:6] == MONEY
    old[1:3] = ["Thay thế về 0 do xuất trùng", DOC]
    old[7] = note
    payload["zero_audit"].append({
        "root": DUPLICATE, "kind": "replacement", "sequence": seq, "docs": [DOC],
        "keys": [DUPLICATE], "before": MONEY, "change": [-v for v in MONEY], "after": [0, 0, 0],
        "identity_source": source["ledger"][DUPLICATE]["source"],
    })
    record = {
        "root": DUPLICATE, "kept_invoice": KEPT, "doc": DOC, "order": ORDER, "money": MONEY,
        "zero_seq": seq, "note": note,
    }
    summary["confirmed_duplicate_12392"] = record

    (stage / "deliverables").mkdir(parents=True, exist_ok=True)
    for kind in ("new", "adjustment", "return_adjustment"):
        name = outputs[kind]["file"]
        shutil.copy2(OUT / "deliverables" / name, stage / "deliverables" / name)
    for name in ("prepared.json", "return-verification.json", "rounding-confirmation.json",
                 "service-quantity-confirmations.json", "unmapped-import-verification.json",
                 "adjustment-rule-2024-verification.json", "unmapped-gift-confirmation.json",
                 "LIEN_KET_BTL_BHH_2025.csv", LINKS):
        shutil.copy2(OUT / name, stage / name)
    write_csv(stage / LEDGER, headers, ledger)
    write_json(stage / "authoring.json", payload)
    write_json(stage / "final-summary.json", summary)
    write_json(stage / "duplicate-12392.json", {**record, "row": draft})
    print("Staged replacement 1990 for duplicate 00012392. Existing sale keeps invoice 00012263.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    args = parser.parse_args()
    stage = args.stage.resolve()
    assert stage.is_relative_to((OUT / "_build").resolve())
    stage_update(stage)
