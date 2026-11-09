"""Stage the confirmed AMIS remap while preserving all four import workbooks."""

import argparse
import shutil
from pathlib import Path

from apply_continued_reissue_20261109 import (
    KEEP, LEDGER, LINKS, OUT, REVIEW, load_delivered_payload, read_csv, read_json, write_csv, write_json,
)
from apply_final_pending_reissues_20261109 import AMIS_ORDER, CANDIDATE, PENDING, WRONG


KEEP_NOTE = (
    f"Giữ {CANDIDATE} cho {PENDING} theo xác nhận người dùng. "
    f"Mã đơn {AMIS_ORDER}, khớp người mua, 8 KH200 và đủ ba khoản. Đã gỡ liên kết {WRONG}."
)
WRONG_NOTE = f"Đã gỡ liên kết nhầm với {PENDING}. Rà riêng chứng từ thực tế của HĐ này."


def stage_keep(stage):
    payload = load_delivered_payload()
    summary = payload["summary"]
    prior = summary["final_pending_reissues_20261109"]["pending"]
    assert prior["doc"] == PENDING and prior["candidate"] == CANDIDATE and prior["sheet_link"] == WRONG
    assert prior["order"] == AMIS_ORDER and summary["actions"]["review"] == 1
    report = next(entry for entry in payload["outputs"] if entry["kind"] == "report")
    tabs = {tab["name"]: tab["rows"] for tab in report["sheets"]}
    headers, ledger = read_csv(LEDGER)
    decisions = [row for row in ledger if row[0] == PENDING]
    assert len(decisions) == 1
    row = decisions[0]
    assert row[2] == REVIEW and row[8] == WRONG
    assert [int(v) for v in row[3:6]] == prior["candidate_money"] == [4659259, 372741, 5032000]
    assert all(r[8] != CANDIDATE for r in ledger)
    assert not any({CANDIDATE, WRONG} & set(z.get("keys", [z["root"]]))
                   for z in payload["zero_audit"] + payload["unmapped_zero_audit"])
    assert PENDING not in read_json("invoice-sequences.json")
    assert {CANDIDATE, WRONG}.isdisjoint(read_json("unmapped-import-sequences.json"))
    row[2], row[7], row[8], row[9], row[-1] = KEEP, "0", CANDIDATE, "Khớp hàng và số lượng", KEEP_NOTE

    pending_rows = tabs["Cho xac minh"]
    assert len(pending_rows) == 1 and pending_rows[0][0] == PENDING
    pending_rows.clear()
    month = tabs["Tong hop"][8]
    assert month[0] == 9 and month[5] == 1
    month[3] += 1
    month[5] -= 1
    summary["actions"]["keep"] += 1
    summary["actions"]["review"] -= 1
    assert summary["actions"]["keep"] == 34832 and summary["actions"]["review"] == 0

    old_chain = next(r for r in tabs["Xu ly chuoi cu"] if r[0] == WRONG)
    assert old_chain[2] == PENDING
    old_chain[1:3] = ["Đã gỡ liên kết, rà riêng", None]
    old_chain[7] = WRONG_NOTE
    tabs["Xu ly chuoi cu"].append([CANDIDATE, "Giữ HĐ", PENDING, *prior["candidate_money"], CANDIDATE, KEEP_NOTE])
    unmapped = next(r for r in tabs["MI2025 chua map"] if r[0] == CANDIDATE)
    assert unmapped[1] == "Tạm giữ, chờ xác minh liên kết AMIS" and not unmapped[10]
    unmapped[1:3] = ["Đã nối AMIS, giữ HĐ", None]
    unmapped[9] = KEEP_NOTE
    summary["confirmed_keep_12515_20261109"] = {
        "doc": PENDING, "invoice": CANDIDATE, "detached_invoice": WRONG, "order": AMIS_ORDER,
        "money": prior["candidate_money"], "source": prior["source"],
        "note": KEEP_NOTE, "detached_note": WRONG_NOTE,
    }

    (stage / "deliverables").mkdir(parents=True, exist_ok=True)
    for output in payload["outputs"]:
        if output["kind"] != "report":
            name = output["file"]
            shutil.copy2(OUT / "deliverables" / name, stage / "deliverables" / name)
    for name in ("prepared.json", "return-verification.json", "rounding-confirmation.json",
                 "service-quantity-confirmations.json", "unmapped-import-verification.json",
                 "adjustment-rule-2024-verification.json", "unmapped-gift-confirmation.json",
                 "LIEN_KET_BTL_BHH_2025.csv", LINKS):
        shutil.copy2(OUT / name, stage / name)
    write_csv(stage / LEDGER, headers, ledger)
    write_json(stage / "authoring.json", payload)
    write_json(stage / "final-summary.json", summary)
    print("Staged keep 00012515. No AMIS vouchers pending; four import workbooks unchanged.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    args = parser.parse_args()
    stage = args.stage.resolve()
    assert stage.is_relative_to((OUT / "_build").resolve())
    stage_keep(stage)
