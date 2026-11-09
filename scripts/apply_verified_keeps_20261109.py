"""Stage the five verified invoice keeps without changing the sales import."""

import argparse
import shutil
from pathlib import Path

from apply_continued_reissue_20261109 import (
    KEEP, LEDGER, LINKS, OUT, load_delivered_payload, read_csv, read_json,
    write_csv, write_json,
)


SPLIT_FROM = "1C25MUV/00002367"
SPLIT_TO = "1C25MUV/00002368"
TAIL = "1C25MUV/00012392"
KEEPS = (
    (SPLIT_TO, "BHHT2025019975", 30799, [2442731, 195418, 2638149]),
    ("1C25MUV/00012243", "BHHT2025027676", 21456, [498147, 39852, 537999]),
    ("1C25MUV/00012247", "BHHT2025027665", 21463, [597268, 47781, 645049]),
    ("1C25MUV/00012251", "BHHT2025027554", 21353, [597268, 47781, 645049]),
    ("1C25MUV/00012273", "BHHT2025027627", 21334, [597268, 47781, 645049]),
)


def stage_update(stage):
    payload = load_delivered_payload()
    summary = payload["summary"]
    outputs = {entry["kind"]: entry for entry in payload["outputs"]}
    tabs = {tab["name"]: tab["rows"] for tab in outputs["report"]["sheets"]}
    chain_rows = {row[0]: row for row in tabs["Xu ly chuoi cu"]}
    replacement = outputs["replacement"]
    replacement_rows = replacement["rows"]
    removed_seq, tail_seq = 608, 1990
    assert len(replacement_rows) == replacement["invoices"] == summary["replacement"] == tail_seq
    assert replacement_rows[removed_seq - 1][25] == SPLIT_TO
    assert replacement_rows[-1][25] == TAIL
    assert len({row[25] for row in replacement_rows}) == tail_seq
    assert not {row[25] for row in replacement_rows} & {record[0] for record in KEEPS[1:]}

    unmapped = payload["unmapped_zero_audit"]
    zero, = [row for row in unmapped if row["root"] == SPLIT_TO]
    assert zero["sequence"] == removed_seq and zero["keys"] == [SPLIT_TO]
    assert zero["before"] == KEEPS[0][3]
    unmapped.remove(zero)
    sequences = read_json("unmapped-import-sequences.json")
    assert sequences.pop(SPLIT_TO) == {"kind": "replacement", "sequence": removed_seq}
    tail_zero, = [row for row in payload["zero_audit"] if row["root"] == TAIL]
    duplicate = summary["confirmed_duplicate_12392"]
    assert tail_zero["sequence"] == duplicate["zero_seq"] == tail_seq
    tail_zero["sequence"] = duplicate["zero_seq"] = removed_seq
    old_marker, new_marker = f"STT {tail_seq}", f"STT {removed_seq}"
    assert old_marker in duplicate["note"] and old_marker in chain_rows[TAIL][7]
    duplicate["note"] = duplicate["note"].replace(old_marker, new_marker)
    chain_rows[TAIL][7] = chain_rows[TAIL][7].replace(old_marker, new_marker)
    tail = replacement_rows.pop()
    tail[0] = removed_seq
    replacement_rows[removed_seq - 1] = tail
    replacement["invoices"] = summary["replacement"] = len(replacement_rows)
    assert [row[0] for row in replacement_rows] == list(range(1, tail_seq))
    summary["unmapped_meinvoice"]["added"]["replacement"] -= 1
    summary["unmapped_meinvoice"]["zeroed_payment"] -= KEEPS[0][3][2]

    headers, ledger = read_csv(LEDGER)
    decisions = {row[0]: row for row in ledger}
    original = decisions["BHHT2025019976"]
    assert original[2] == KEEP and original[8] == SPLIT_FROM
    assert list(map(int, original[3:6])) == KEEPS[0][3]
    original[-1] = f"Giữ HĐ {SPLIT_FROM} theo mã đơn và người mua. Đã tách chứng từ gắn nhầm."
    assert old_marker in decisions[duplicate["doc"]][-1]
    decisions[duplicate["doc"]][-1] = decisions[duplicate["doc"]][-1].replace(old_marker, new_marker)
    sale_docs = {row[18] for row in outputs["new"]["rows"]}
    records = []
    for root, doc, source_row, money in KEEPS:
        decision = decisions[doc]
        assert decision[2] == KEEP and list(map(int, decision[3:6])) == money
        assert int(decision[6]) == source_row and int(decision[7]) == 0 and doc not in sale_docs
        assert decision[8] == (SPLIT_FROM if root == SPLIT_TO else root)
        note = (f"Giữ HĐ {root} cho {doc}. Đối chiếu AMIS dòng {source_row}, khớp đủ ba khoản. "
                + ("Mã đơn và tên người mua khớp. Đã rút khỏi lô thay thế về 0."
                   if root == SPLIT_TO else "Đã tách chứng từ gắn nhầm và cộng lại vào doanh thu giữ lại."))
        decision[8], decision[-1] = root, note
        if root == SPLIT_TO:
            row, = [row for row in tabs["MI2025 chua map"] if row[0] == root]
            assert row[4:7] == money and row[10] == removed_seq
            row[1:3] = ["Đã nối AMIS và giữ HĐ", None]
            row[9:11] = [note, None]
            tabs["Xu ly chuoi cu"].append([root, "Giữ HĐ", doc, *money, root, note])
        else:
            row = chain_rows[root]
            assert doc in row[2].split("; ") and row[3:6] == money
            row[1:3], row[7] = ["Giữ HĐ", doc], note
        records.append({"root": root, "doc": doc, "amis_row": source_row, "money": money})
    assert not {r["root"] for r in records} & {
        key for audit in payload["zero_audit"] + unmapped for key in audit.get("keys", [audit["root"]])
    }

    reporting = summary["amis_reporting_0911"]
    assert reporting["gap"] == [-4871571, -515483]
    delta = [sum(record["money"][i] for record in records) for i in range(2)]
    reporting["projection"] = [a + b for a, b in zip(reporting["projection"], delta)]
    reporting["gap"] = [a - b for a, b in zip(reporting["projection"], reporting["scope_current"])]
    assert reporting["gap"] == [-138889, -136870]
    summary["verified_keeps_20261109"] = {
        "records": records, "removed_stt": removed_seq, "moved_root": TAIL,
        "moved_from_stt": tail_seq, "moved_to_stt": removed_seq, "added_to_projection": delta,
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
    write_json(stage / "unmapped-import-sequences.json", sequences)
    write_json(stage / "final-summary.json", summary)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    stage = parser.parse_args().stage.resolve()
    assert stage.is_relative_to((OUT / "_build").resolve())
    stage_update(stage)
