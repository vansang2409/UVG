"""Use the approved AMIS 09/11 VAT amounts on the four remaining PSC2 drafts."""

import argparse
import shutil
from collections import Counter, defaultdict
from pathlib import Path

from apply_amis_0911_reporting import LATEST, load_sales
from apply_continued_reissue_20261109 import (
    LEDGER, LINKS, OUT, digest, load_delivered_payload, load_workbook, number,
    read_csv, write_csv, write_json,
)


IMPORT_NOTE = ("Sáu chứng từ đã bỏ PSC2. Tiền hàng và tiền thuế GTGT lấy theo AMIS 09/11. "
               "Phiếu trả hàng lập riêng.")


def detail_by_doc(path, docs):
    result = defaultdict(list)
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in workbook:
            for row_number, row in enumerate(sheet.values, 1):
                if len(row) <= 43 or row[2] not in docs:
                    continue
                sku, qty, price, goods, vat, pay = (row[i] for i in (15, 22, 23, 24, 41, 43))
                assert sku != "PSC2" and goods + vat == pay
                result[row[2]].append({"source": f"{sheet.title}!{row_number}",
                                       "item": [sku, qty, price, goods], "money": [goods, vat, pay]})
    finally:
        workbook.close()
    assert result.keys() == docs
    return result


def stage_update(stage):
    payload = load_delivered_payload()
    summary = payload["summary"]
    reporting = summary["amis_reporting_0911"]
    exchanges = summary["psc2_exchanges"]
    docs = {code for code, e in exchanges.items() if e.get("rate_policy") == "recalculate"}
    assert len(docs) == 4 and reporting["gap"] == [0, -125759]
    for source in reporting["source_files"]:
        assert digest(LATEST / Path(source["file"]).name) == source["sha256"]
    latest_sales = load_sales(LATEST / "Ban_hang.xlsx")
    latest_detail = detail_by_doc(LATEST / "So_chi_tiet_ban_hang.xlsx", exchanges.keys())

    outputs = {item["kind"]: item for item in payload["outputs"]}
    sales = outputs["new"]
    sales["notes"][2] = IMPORT_NOTE
    tabs = {tab["name"]: tab["rows"] for tab in outputs["report"]["sheets"]}
    ledger_header, ledger = read_csv(LEDGER)
    ledger_rows = {row[0]: row for row in ledger}
    old_chains = {row[0]: row for row in tabs["Xu ly chuoi cu"]}
    zero_plans = {audit["root"]: audit for audit in payload["zero_audit"]}
    import_by_doc = defaultdict(list)
    for excel_row, row in enumerate(sales["rows"], 10):
        if row[18] in exchanges:
            import_by_doc[row[18]].append((excel_row, row))

    changes = {}
    for code, exchange in exchanges.items():
        money = latest_sales[code]["money"]
        detail = latest_detail[code]
        assert money == [sum(line["money"][i] for line in detail) for i in range(3)]
        assert money == reporting["psc2_documents"][code]["money"] == list(map(int, ledger_rows[code][3:6]))
        rows = import_by_doc[code]
        assert len(rows) == len(detail)
        for (_, row), line in zip(rows, detail):
            assert [row[i] for i in (19, 14, 15, 16)] == line["item"]
        excel_row, first = rows[0]
        assert all(row[11] is None for _, row in rows[1:])
        if code not in docs:
            assert first[11] == money[1] == exchange["vat"]
            continue
        assert first[11] == exchange["vat"] and exchange["goods"] == money[0]
        changes[code] = {"row": excel_row, "stt": first[0], "old_vat": first[11],
                         "money": money, "source_lines": detail}
        first[11] = money[1]
        exchange.update(vat=money[1], payment=money[2], rate_policy="amis_0911")
        reporting["psc2_documents"][code]["draft"] = money

        root = exchange["old_root"]
        plan = zero_plans[root]
        if plan["kind"] == "replacement":
            rec, = [rec for rec in summary["continued_reissue_20261109"]["records"] if rec["root"] == root]
            rec["new_totals"][code] = money
            zero_note = f"Thay thế HĐ cuối {rec['active']} về 0, STT {plan['sequence']}."
        else:
            rec, = [rec for rec in summary["final_pending_reissues_20261109"]["records"] if rec["code"] == code]
            rec["new"] = money
            reversal, reduction = rec["zero_seqs"]
            zero_note = f"Đảo F1 STT {reversal}, giảm F0 STT {reduction}."
        vat_text = f"{money[1]:,}".replace(",", ".")
        note = f"{zero_note} Bán mới STT {first[0]}. VAT {vat_text}đ theo AMIS 09/11."
        ledger_rows[code][-1] = note
        old_chains[root][7] = note
        if plan["kind"] == "adjustment":
            rec["note"] = note

    delta = sum(change["money"][1] - change["old_vat"] for change in changes.values())
    assert delta == -reporting["vat_policy_delta"] == 125759
    reporting["projection"][1] += delta
    assert reporting["projection"] == reporting["scope_current"]
    reporting["gap"] = [0, 0]
    reporting["vat_policy_delta"] = 0
    summary["new_vat"] += delta
    summary["new_pay"] += delta
    summary["amis_0911_import_vat"] = changes

    monthly = Counter()
    for row in sales["rows"]:
        monthly[int(row[20][5:7])] += number(row[16]) + number(row[11])
    assert sum(monthly.values()) == summary["new_pay"]
    for row in tabs["Tong hop"]:
        row[6] = int(monthly[row[0]])

    (stage / "deliverables").mkdir(parents=True, exist_ok=True)
    for kind in ("replacement", "adjustment", "return_adjustment"):
        name = outputs[kind]["file"]
        shutil.copy2(OUT / "deliverables" / name, stage / "deliverables" / name)
    for name in ("prepared.json", "return-verification.json", "rounding-confirmation.json",
                 "service-quantity-confirmations.json", "unmapped-import-verification.json",
                 "adjustment-rule-2024-verification.json", "unmapped-gift-confirmation.json",
                 "LIEN_KET_BTL_BHH_2025.csv", LINKS):
        shutil.copy2(OUT / name, stage / name)
    write_csv(stage / LEDGER, ledger_header, ledger)
    write_json(stage / "authoring.json", payload)
    write_json(stage / "final-summary.json", summary)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    stage = parser.parse_args().stage.resolve()
    assert stage.is_relative_to((OUT / "_build").resolve())
    stage_update(stage)
