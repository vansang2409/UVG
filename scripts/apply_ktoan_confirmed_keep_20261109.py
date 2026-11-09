"""Stage seven verified accounting keep decisions and refresh the Excel report."""

import argparse
import json
import shutil
from pathlib import Path

from apply_continued_reissue_20261109 import (
    OUT, KEEP, REVIEW, LEDGER, LINKS, load_delivered_payload,
    read_csv, read_json, write_csv, write_json,
)


def stage_keep(stage):
    plan = json.loads((stage / "keep-plan.json").read_text(encoding="utf-8"))
    payload = load_delivered_payload()
    prepared = read_json("prepared.json")
    summary = payload["summary"]
    headers, ledger = read_csv(LEDGER)
    decisions = {row[0]: row for row in ledger}
    report = next(item for item in payload["outputs"] if item["kind"] == "report")
    tabs = {tab["name"]: tab for tab in report["sheets"]}
    already_planned = {row["root"] for row in payload["zero_audit"] + payload["unmapped_zero_audit"]}
    seen = set()
    for entry in plan["records"]:
        doc = entry["doc"]
        key = entry["invoice"]
        row = decisions[doc]
        chain = prepared["chains"][key]
        assert row[2] == REVIEW and key in row[8].split("; ")
        assert chain["active"] == [key] and not chain["issues"] and key not in already_planned
        assert chain["balance"] == entry["money"] == [int(value) for value in row[3:6]]
        assert str(prepared["docs"][doc]["order"]) == entry["order"]
        assert key not in seen
        seen.add(key)
        row[2], row[7], row[8] = KEEP, "0", key
        row[9] = "Hàng ghi theo combo" if doc == "BHHT2025038170" else "Khớp hàng và số lượng"
        other = ", ".join(item["key"] for item in entry["other"])
        row[10] = f"Giữ {key}, đúng mã đơn và ba khoản. KT ghi xuất đúng. Chuỗi còn lại theo dõi riêng: {other}."
        if doc == "BHHT2025038170":
            row[10] += " " + plan["product_note"]
        month = int(row[1][5:7])
        monthly = tabs["Tong hop"]["rows"][month - 1]
        assert monthly[0] == month and monthly[5] > 0
        monthly[3] += 1
        monthly[5] -= 1

    codes = {row["doc"] for row in plan["records"]}
    pending = tabs["Cho xac minh"]["rows"]
    tabs["Cho xac minh"]["rows"] = [row for row in pending if row[0] not in codes]
    summary["actions"]["keep"] += len(codes)
    summary["actions"]["review"] -= len(codes)
    assert len(codes) == 7 and summary["actions"]["review"] == len(tabs["Cho xac minh"]["rows"]) == 9
    assert summary["replacement"] == 1989 and summary["new_invoices"] == 30259
    summary["ktoan_confirmed_keep_20261109"] = plan

    (stage / "deliverables").mkdir(exist_ok=True)
    for item in payload["outputs"]:
        if item["kind"] != "report":
            name = item["file"]
            shutil.copy2(OUT / "deliverables" / name, stage / "deliverables" / name)
    for name in ("prepared.json", "return-verification.json", "rounding-confirmation.json",
                 "service-quantity-confirmations.json", "unmapped-import-verification.json",
                 "adjustment-rule-2024-verification.json", "unmapped-gift-confirmation.json",
                 "LIEN_KET_BTL_BHH_2025.csv", LINKS):
        shutil.copy2(OUT / name, stage / name)
    write_csv(stage / LEDGER, headers, ledger)
    write_json(stage / "authoring.json", payload)
    write_json(stage / "final-summary.json", summary)
    print("Staged seven confirmed keeps. Nine AMIS vouchers remain pending.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    args = parser.parse_args()
    assert args.stage.resolve().is_relative_to((OUT / "_build").resolve())
    stage_keep(args.stage.resolve())
