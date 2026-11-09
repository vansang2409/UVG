"""Stage the verified links and remove the two valid invoices from replacement drafts."""

import argparse
import shutil
from pathlib import Path

from apply_continued_reissue_20261109 import (
    OUT, LEDGER, NEW, REVIEW, KEEP,
    load_delivered_payload, read_csv, read_json, write_csv, write_json,
)


CASES = (
    ("BHHT2025027442", "580356480920290819", "1C25MUV/00012483", 1216,
     "BHHT2025028267", "580299168658785596", "1C25MUV/00010850", "1C25MUV/00012360"),
    ("BHHT2025027439", "580357260801967142", "1C25MUV/00012486", 1219,
     "BHHT2025028252", "580300943056209030", "1C25MUV/00010855", "1C25MUV/00012363"),
)


def stage_fix(stage):
    payload = load_delivered_payload()
    prepared = read_json("prepared.json")
    headers, ledger = read_csv(LEDGER)
    decisions = {row[0]: row for row in ledger}
    outputs = {item["kind"]: item for item in payload["outputs"]}
    tabs = {tab["name"]: tab for tab in outputs["report"]["sheets"]}
    repl = outputs["replacement"]
    summary = payload["summary"]
    matched_audit = {row["root"]: row for row in payload["zero_audit"]}
    orphan_audit = {row["root"]: row for row in payload["unmapped_zero_audit"]}
    historical_sequences = read_json("unmapped-import-sequences.json")
    continued = summary["continued_reissue_20261109"]
    continued_records = {record["root"]: record for record in continued["records"]}
    assert summary["new_invoices"] == 30259 and summary["actions"]["review"] == 18
    assert len(repl["rows"]) == summary["replacement"] == 1991
    records, moved = [], []
    for offset, (code, order, actual, dropped_seq, kept, kept_order, root, active) in enumerate(CASES):
        doc = prepared["docs"][code]
        kept_doc = prepared["docs"][kept]
        old = prepared["chains"][root]
        money = [doc[key] for key in ("goods", "vat", "pay")]
        assert str(doc["order"]) == order and str(kept_doc["order"]) == kept_order
        assert decisions[code][2] == REVIEW and decisions[code][8] == root
        assert decisions[kept][2] == KEEP and decisions[kept][8] == root
        assert old["active"] == [active] and old["balance"] == [kept_doc[k] for k in ("goods", "vat", "pay")]
        assert orphan_audit[actual]["before"] == money and orphan_audit[actual]["keys"] == [actual]
        assert repl["rows"][dropped_seq - 1][25] == actual
        assert decisions[code][0] not in {r[18] for r in outputs["new"]["rows"]}
        tail_seq = 1990 + offset
        tail = repl["rows"][tail_seq - 1].copy()
        moved_root = tail[25]
        assert tail[0] == continued_records[moved_root]["zero_seq"] == matched_audit[moved_root]["sequence"] == tail_seq
        tail[0] = dropped_seq
        repl["rows"][dropped_seq - 1] = tail
        matched_audit[moved_root]["sequence"] = continued_records[moved_root]["zero_seq"] = dropped_seq
        moved.append({"root": moved_root, "active": f"{tail[11]}/{tail[12]}", "doc": tail[24],
                      "from": tail_seq, "to": dropped_seq, "row": tail})
        assert historical_sequences[actual]["sequence"] == dropped_seq
        del historical_sequences[actual]
        del orphan_audit[actual]
        decisions[code][2], decisions[code][7], decisions[code][8] = KEEP, "0", actual
        decisions[code][9] = "Khớp hàng và số lượng"
        decisions[code][-1] = f"Giữ HĐ riêng {actual}. Đối chiếu mã đơn {order} qua NB, người mua, hàng và đủ 3 khoản. Gỡ map nhầm {active}."
        decisions[kept][-1] = f"Giữ HĐ cuối {active} theo mã đơn {kept_order}. Khớp đủ 3 khoản."
        for row in tabs["Xu ly chuoi cu"]["rows"]:
            if row[0] == root:
                row[1:3] = ["Giữ HĐ cuối chuỗi", kept]
                row[7] = f"Giữ {active}. Đã gỡ liên kết nhầm của {code}."
        for row in tabs["MI2025 chua map"]["rows"]:
            if row[0] == actual:
                row[1:3] = ["Đã nối AMIS, giữ HĐ", None]
                row[9:11] = [f"{code}, mã đơn {order}. Đã rút khỏi dự thảo thay thế.", None]
        records.append({"doc": code, "order": order, "actual": actual, "money": money,
                        "removed_replacement_seq": dropped_seq, "kept_doc": kept,
                        "kept_order": kept_order, "root": root, "kept_invoice": active,
                        "kept_money": old["balance"]})

    del repl["rows"][-2:]
    repl["invoices"] = len(repl["rows"])
    assert [row[0] for row in repl["rows"]] == list(range(1, 1990))
    payload["unmapped_zero_audit"] = list(orphan_audit.values())
    confirmed = {row["doc"] for row in records}
    tabs["Cho xac minh"]["rows"] = [row for row in tabs["Cho xac minh"]["rows"] if row[0] not in confirmed]
    month = tabs["Tong hop"]["rows"][8]
    assert month[0] == 9
    month[3] += 2
    month[5] -= 2
    summary["replacement"] = repl["invoices"]
    summary["actions"]["keep"] += 2
    summary["actions"]["review"] -= 2
    assert len(tabs["Cho xac minh"]["rows"]) == summary["actions"]["review"] == 16
    summary["unmapped_meinvoice"]["added"]["replacement"] -= 2
    summary["unmapped_meinvoice"]["zeroed_payment"] -= sum(row["money"][2] for row in records)
    continued["remaining"] = [code for code in continued["remaining"] if code not in confirmed]
    summary["shared_invoice_map_fix_20261109"] = {
        "records": records, "moved": [{k: v for k, v in x.items() if k != "row"} for x in moved],
        "evidence": "Mã đơn từ NB. Mỗi HĐ riêng khớp ngày, người mua sau chuẩn hóa, mặt hàng, số lượng và đủ 3 khoản. Cột mã đơn gốc meInvoice để trống.",
    }

    (stage / "deliverables").mkdir(parents=True, exist_ok=True)
    for kind in ("new", "adjustment", "return_adjustment"):
        name = outputs[kind]["file"]
        shutil.copy2(OUT / "deliverables" / name, stage / "deliverables" / name)
    for name in ("prepared.json", "return-verification.json", "rounding-confirmation.json",
                 "service-quantity-confirmations.json", "unmapped-import-verification.json",
                 "adjustment-rule-2024-verification.json", "unmapped-gift-confirmation.json",
                 "LIEN_KET_BTL_BHH_2025.csv", "LIEN_KET_IMPORT_MOI.csv"):
        shutil.copy2(OUT / name, stage / name)
    write_csv(stage / LEDGER, headers, ledger)
    write_json(stage / "authoring.json", payload)
    write_json(stage / "unmapped-import-sequences.json", historical_sequences)
    write_json(stage / "final-summary.json", summary)
    write_json(stage / "native-keep.json", {"records": records, "moved": moved, "replacement": repl["invoices"]})
    print("Staged 4 kept invoices. Replacement count 1989. New sales count 30259.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    args = parser.parse_args()
    assert args.stage.resolve().is_relative_to((OUT / "_build").resolve())
    stage_fix(args.stage.resolve())
