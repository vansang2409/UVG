"""Stage the confirmed KOL/KOC keeps and the AMIS deletion for delivery."""

import argparse
import shutil
from pathlib import Path

from apply_continued_reissue_20261109 import (
    LEDGER, LINKS, OUT, REVIEW, load_delivered_payload, read_csv, read_json,
    write_csv, write_json,
)


GIFT_CASES = (
    ("1C25TUV/00013178", 257, "2025-12-31", [24810677, 1984854, 26795531]),
    ("1C25TUV/00013158", 706, "2025-10-31", [20838358, 1667069, 22505427]),
    ("1C25TUV/00013177", 707, "2025-12-20", [25252874, 2020230, 27273104]),
)
DELETED_DOC = "BHDV20250045"
DELETED_LABEL = "Đã xóa trên AMIS"
DELETED_NOTE = "Người dùng xác nhận đã xóa trên AMIS. Lưu dấu vết bản trích cũ."
GIFT_LABEL = "Giữ HĐ biếu tặng KOL/KOC"
GIFT_NOTE = "Danh sách KOL/KOC bổ sung, người dùng xác nhận giữ. Theo dõi ngoài doanh thu bán hàng."


def stage_update(stage):
    payload = load_delivered_payload()
    summary = payload["summary"]
    outputs = {item["kind"]: item for item in payload["outputs"]}
    tabs = {tab["name"]: tab["rows"] for tab in outputs["report"]["sheets"]}
    replacement = outputs["replacement"]
    replacement_rows = replacement["rows"]
    matched = {entry["root"]: entry for entry in payload["zero_audit"]}
    unmapped = {entry["root"]: entry for entry in payload["unmapped_zero_audit"]}
    historical_sequences = read_json("unmapped-import-sequences.json")
    continued = {entry["root"]: entry for entry in summary["continued_reissue_20261109"]["records"]}
    unmapped_rows = {row[0]: row for row in tabs["MI2025 chua map"]}
    chain_rows = {row[0]: row for row in tabs["Xu ly chuoi cu"]}
    assert len(replacement_rows) == replacement["invoices"] == summary["replacement"] == 1989

    records, moves = [], []
    for tail_seq, (root, seq, invoice_date, amounts) in enumerate(GIFT_CASES, 1987):
        plan = unmapped.pop(root)
        assert plan["keys"] == [root] and plan["before"] == amounts
        assert plan["sequence"] == seq == replacement_rows[seq - 1][0]
        assert replacement_rows[seq - 1][25] == root and not replacement_rows[seq - 1][24]
        assert historical_sequences.pop(root) == {"kind": "replacement", "sequence": seq}
        report_row = unmapped_rows[root]
        assert report_row[4:7] == amounts and report_row[10] == seq
        report_row[1:3] = [GIFT_LABEL, None]
        report_row[9:11] = [GIFT_NOTE, None]
        records.append({"root": root, "date": invoice_date, "money": amounts, "removed_seq": seq})

        tail = replacement_rows[tail_seq - 1].copy()
        moved_root = tail[25]
        assert tail[0] == matched[moved_root]["sequence"] == continued[moved_root]["zero_seq"] == tail_seq
        tail[0] = seq
        replacement_rows[seq - 1] = tail
        matched[moved_root]["sequence"] = continued[moved_root]["zero_seq"] = seq
        marker = f"STT {tail_seq}"
        assert chain_rows[moved_root][7].count(marker) == 1
        chain_rows[moved_root][7] = chain_rows[moved_root][7].replace(marker, f"STT {seq}")
        moves.append({"root": moved_root, "from": tail_seq, "to": seq, "row": tail})

    del replacement_rows[-len(GIFT_CASES):]
    replacement["invoices"] = summary["replacement"] = len(replacement_rows)
    assert [row[0] for row in replacement_rows] == list(range(1, 1987))
    payload["unmapped_zero_audit"] = list(unmapped.values())
    summary["unmapped_meinvoice"]["added"]["replacement"] -= len(GIFT_CASES)
    summary["unmapped_meinvoice"]["zeroed_payment"] -= sum(entry["money"][2] for entry in records)

    headers, ledger = read_csv(LEDGER)
    matches = [row for row in ledger if row[0] == DELETED_DOC]
    assert len(matches) == 1
    deleted = matches[0]
    assert deleted[1:6] == ["2025-07-28", REVIEW, "0", "0", "0"] and not deleted[8]
    assert not any(row[18] == DELETED_DOC for row in outputs["new"]["rows"])
    assert not any(row[24] == DELETED_DOC for row in outputs["return_adjustment"]["rows"])
    assert not any(DELETED_DOC in row for row in tabs["Lien ket BTL"])
    deleted[2], deleted[-1] = DELETED_LABEL, DELETED_NOTE
    pending = tabs["Cho xac minh"]
    assert sum(row[0] == DELETED_DOC for row in pending) == 1
    pending[:] = [row for row in pending if row[0] != DELETED_DOC]
    july = tabs["Tong hop"][6]
    assert july[0] == 7 and july[5] == 1
    july[1] -= 1
    july[5] -= 1
    summary["total"] -= 1
    summary["actions"]["review"] -= 1
    assert summary["actions"]["review"] == len(pending) == 6
    assert summary["total"] == sum(summary["actions"].values()) == len(ledger) - 1
    assert all(row[1] == sum(row[2:6]) for row in tabs["Tong hop"])
    summary["gift_keeps_and_amis_deletion_20261109"] = {
        "records": records,
        "moves": [{key: value for key, value in move.items() if key != "row"} for move in moves],
        "deleted_amis": [DELETED_DOC],
        "source_workbook": "Đơn hàng KOC ^M từ thiện 2025 bổ sung.xlsx",
        "reason": DELETED_NOTE,
    }

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
    write_json(stage / "unmapped-import-sequences.json", historical_sequences)
    write_json(stage / "final-summary.json", summary)
    write_json(stage / "native-update.json", {
        "records": records, "moves": moves, "replacement": replacement["invoices"],
        "remaining": [row[0] for row in pending],
    })
    print("Staged 3 KOL/KOC keeps, 1 AMIS deletion, 1,986 replacements and 6 pending.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    args = parser.parse_args()
    stage = args.stage.resolve()
    assert stage.is_relative_to((OUT / "_build").resolve())
    stage_update(stage)
