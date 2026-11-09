"""Link the existing replacement and new-sale drafts for BHHT2025027502."""

import argparse
import shutil
from pathlib import Path

from apply_continued_reissue_20261109 import (
    KEEP, LEDGER, LINKS, NEW, OUT, load_delivered_payload, read_csv, read_json, write_csv, write_json,
)


DOC = "BHHT2025027502"
ROOT = "1C25MUV/00012497"
OTHER_ROOT = "1C25MUV/00012374"
ORDER = "580358394021512445"
REPLACEMENT_SEQ = 1230
SALE_SEQ = 29640
OLD_TOTALS = [607322, 48586, 655908]
NEW_TOTALS = [1276767, 102141, 1378908]
REASON = f"Thay thế HĐ {ROOT} về 0 do lệch mặt hàng và số tiền; lập lại theo chứng từ {DOC}."
NOTE = (
    f"Đã chốt thay thế {ROOT} về 0, STT {REPLACEMENT_SEQ}, và bán mới {DOC}, STT {SALE_SEQ}, "
    "theo AMIS 09/11: 1 KA150P, thanh toán 1.378.908đ."
)


def stage_update(stage):
    payload = load_delivered_payload()
    summary = payload["summary"]
    previous = summary["kt_final_8_20261109"]["mismatch"]
    assert previous["doc"] == DOC and previous["key"] == ROOT and previous["old"] == OTHER_ROOT
    assert previous["amis"] == NEW_TOTALS and previous["mi"] == OLD_TOTALS
    assert summary["actions"]["review"] == 0 and summary["replacement"] == 1989
    outputs = {o["kind"]: o for o in payload["outputs"]}
    sale = [r for r in outputs["new"]["rows"] if r[18] == DOC]
    assert len(sale) == 1
    r = sale[0]
    assert (r[0], r[17], r[19], r[14], r[16], r[11]) == (SALE_SEQ, ORDER, "KA150PFULL", 1, *NEW_TOTALS[:2])
    replacement = outputs["replacement"]["rows"][REPLACEMENT_SEQ - 1]
    assert replacement[0] == REPLACEMENT_SEQ and replacement[25] == ROOT and not replacement[24]
    assert "/".join(replacement[11:13]) == ROOT and replacement[14]
    assert all(replacement[i] == 0 for i in (17, 18, 19, 21))
    replacement[15] = replacement[22] = REASON
    replacement[24] = DOC

    previous_plans = [row for row in payload["unmapped_zero_audit"] if row["root"] == ROOT]
    assert len(previous_plans) == 1
    zero = previous_plans[0]
    assert zero["sequence"] == REPLACEMENT_SEQ and zero["before"] == OLD_TOTALS and zero["keys"] == [ROOT]
    assert not any(z["root"] == ROOT for z in payload["zero_audit"])
    payload["unmapped_zero_audit"].remove(zero)
    payload["zero_audit"].append({**zero, "docs": [DOC]})
    unmapped_seqs = read_json("unmapped-import-sequences.json")
    assert unmapped_seqs.pop(ROOT) == {"kind": "replacement", "sequence": REPLACEMENT_SEQ}
    summary["unmapped_meinvoice"]["added"]["replacement"] -= 1
    summary["unmapped_meinvoice"]["zeroed_payment"] -= OLD_TOTALS[2]
    hold = "BHHT2025027502 chờ KT chốt chênh thanh toán 723.000đ trước khi dùng."
    assert hold in summary["limitations"]
    summary["limitations"].remove(hold)

    ledger_headers, ledger = read_csv(LEDGER)
    decisions = {r[0]: r for r in ledger}
    decision = decisions[DOC]
    assert decision[2] == NEW and decision[8] == OTHER_ROOT and list(map(int, decision[3:6])) == NEW_TOTALS
    assert decisions["BHHT2025027561"][2] == KEEP and decisions["BHHT2025027561"][8] == OTHER_ROOT
    decision[7], decision[8], decision[9], decision[-1] = "723000", ROOT, "Khác mặt hàng", NOTE

    link_headers, links = read_csv(LINKS)
    link, = [r for r in links if r[2] == DOC]
    assert link[1:] == [str(SALE_SEQ), DOC, "34284", "1", OTHER_ROOT]
    link[-1] = ROOT
    tabs = {tab["name"]: tab["rows"] for tab in outputs["report"]["sheets"]}
    unmatched, = [r for r in tabs["MI2025 chua map"] if r[0] == ROOT]
    assert unmatched[2] == "replacement" and unmatched[10] == REPLACEMENT_SEQ
    unmatched[1], unmatched[9] = "Đã nối AMIS, thay thế về 0", NOTE
    tabs["Xu ly chuoi cu"].append([ROOT, "Thay thế về 0", DOC, *OLD_TOTALS, ROOT, NOTE])
    summary["confirmed_reissue_7502"] = {
        "root": ROOT, "active": ROOT, "keys": [ROOT], "docs": [DOC], "before": OLD_TOTALS,
        "zero_seq": REPLACEMENT_SEQ, "new_sequences": {DOC: SALE_SEQ}, "new_totals": {DOC: NEW_TOTALS},
        "order": ORDER, "detached_root": OTHER_ROOT, "reason": REASON, "note": NOTE,
    }

    (stage / "deliverables").mkdir(parents=True, exist_ok=True)
    for kind in ("new", "adjustment", "return_adjustment"):
        filename = outputs[kind]["file"]
        shutil.copy2(OUT / "deliverables" / filename, stage / "deliverables" / filename)
    for name in ("prepared.json", "return-verification.json", "rounding-confirmation.json",
                 "service-quantity-confirmations.json", "unmapped-import-verification.json",
                 "adjustment-rule-2024-verification.json", "unmapped-gift-confirmation.json",
                 "LIEN_KET_BTL_BHH_2025.csv"):
        shutil.copy2(OUT / name, stage / name)
    write_csv(stage / LEDGER, ledger_headers, ledger)
    write_csv(stage / LINKS, link_headers, links)
    write_json(stage / "authoring.json", payload)
    write_json(stage / "final-summary.json", summary)
    write_json(stage / "unmapped-import-sequences.json", unmapped_seqs)
    write_json(stage / "reissue-7502.json", summary["confirmed_reissue_7502"])
    print("Confirmed replacement 1230 and new sale 29640 for BHHT2025027502.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    args = parser.parse_args()
    stage = args.stage.resolve()
    assert stage.is_relative_to((OUT / "_build").resolve())
    stage_update(stage)
