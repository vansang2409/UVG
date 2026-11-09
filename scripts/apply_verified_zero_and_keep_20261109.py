"""Stage one new sale after a verified zero chain and one verified invoice keep."""

import argparse
import shutil
from decimal import Decimal
from pathlib import Path

from apply_continued_reissue_20261109 import (
    OUT, ROOT, LEDGER, LINKS, NEW, KEEP, REVIEW,
    load_delivered_payload, money_round, read_csv, read_json, write_csv, write_json,
)
from import_2025_buyer_fields import apply_buyer_fields


NEW_DOC = "BHHT2025035946"
KEPT_DOC = "BHHT2025040425"
ZERO_CHAIN = ["1C25MUV/00031054", "1C25MUV/00034893", "1C26MUV/00002639"]
KEPT_INVOICE = "1C26MUV/00005776"
OTHER_ZERO_CHAIN = ["1C26MUV/00038350", "1C26MUV/00058635"]
PORTAL_DETAIL = {
    "invoice": KEPT_INVOICE, "invoice_date": "2026-01-27", "order": "580221132778210811",
    "model": "KA150P", "quantity": 2, "unit": "Cái", "goods": 2461111, "vat": 196889,
    "pay": 2658000, "source": "Bản thể hiện hóa đơn tại trang tra cứu chính thức MISA meInvoice",
}


def stage_update(stage):
    payload = load_delivered_payload()
    prepared = read_json("prepared.json")
    ledger_index = read_json("source-index.json")["ledger"]
    summary = payload["summary"]
    outputs = {item["kind"]: item for item in payload["outputs"]}
    tabs = {tab["name"]: tab for tab in outputs["report"]["sheets"]}
    sales = outputs["new"]
    headers, ledger = read_csv(LEDGER)
    by_doc = {row[0]: row for row in ledger}
    assert len(sales["rows"]) == 35036 and sales["invoices"] == 30259
    assert summary["actions"]["review"] == 9
    assert by_doc[NEW_DOC][2] == by_doc[KEPT_DOC][2] == REVIEW
    sequences = {row[18]: row[0] for row in sales["rows"]}
    assert NEW_DOC not in sequences and KEPT_DOC not in sequences
    assert NEW_DOC not in {row[25] for row in outputs["return_adjustment"]["rows"]}
    assert not {ZERO_CHAIN[0], ZERO_CHAIN[1], KEPT_INVOICE, OTHER_ZERO_CHAIN[0]} & {
        row["root"] for row in payload["zero_audit"] + payload["unmapped_zero_audit"]
    }

    new_doc = prepared["docs"][NEW_DOC]
    kept_doc = prepared["docs"][KEPT_DOC]
    line, = new_doc["lines"]
    item = line["v"]
    goods, vat, pay = [int(new_doc[k]) for k in ("goods", "vat", "pay")]
    assert item[15] == "KA350STANDARDFULL" and item[22] == 1
    assert item[22] * item[23] == item[24] == goods and item[41] == vat
    assert goods + vat == pay and money_round(Decimal(goods) * Decimal("0.08")) == vat
    assert all(str(ledger_index[k]["order"]) == str(new_doc["order"]) for k in ZERO_CHAIN)
    assert prepared["chains"][ZERO_CHAIN[1]]["active"] == [ZERO_CHAIN[-1]]
    assert ledger_index[ZERO_CHAIN[-1]]["money"] in (["0", "0", None], ["0", "0", "0"])
    assert prepared["chains"][OTHER_ZERO_CHAIN[0]]["balance"] == [0, 0, 0]

    kept_line, = kept_doc["lines"]
    assert kept_line["v"][15] == "KA150PFULL" and kept_line["v"][22] == PORTAL_DETAIL["quantity"]
    assert str(kept_doc["order"]) == PORTAL_DETAIL["order"]
    assert ledger_index[KEPT_INVOICE]["order"] == PORTAL_DETAIL["order"]
    assert ledger_index[KEPT_INVOICE]["date"] == PORTAL_DETAIL["invoice_date"]
    assert prepared["chains"][KEPT_INVOICE]["active"] == [KEPT_INVOICE]
    assert prepared["chains"][KEPT_INVOICE]["balance"] == [PORTAL_DETAIL[k] for k in ("goods", "vat", "pay")]
    assert [kept_doc[k] for k in ("goods", "vat", "pay")] == [PORTAL_DETAIL[k] for k in ("goods", "vat", "pay")]

    seq = len(sequences) + 1
    excel_row = 10 + len(sales["rows"])
    draft = [
        seq, sales["rows"][0][1], new_doc["buyer"], item[13] or None, item[12] or None,
        item[14] or None, None, None, None, "TM/CK", 8, vat, str(item[7] or item[16]).strip(),
        item[19], item[22], item[23], goods, new_doc["order"], NEW_DOC, item[15],
        new_doc["date"], line["source"], None,
    ]
    apply_buyer_fields({"outputs": [{**sales, "rows": [draft]}], "manifest": []}, OUT, ROOT)
    sales["rows"].append(draft)
    sequences[NEW_DOC] = seq
    sales["invoices"] = seq
    assert seq == 30260 and excel_row == 35046
    sales["notes"][0] = "Gộp toàn bộ bán mới năm 2025. STT liên tục 1–30.260. Cột Ngày chứng từ AMIS giữ thông tin kỳ gốc."
    sales["notes"].append("BHHT2025027502 chờ KT chốt chênh thanh toán 723.000đ trước khi dùng.")

    by_doc[NEW_DOC][2], by_doc[NEW_DOC][8] = NEW, ZERO_CHAIN[0]
    by_doc[NEW_DOC][10] = f"Chuỗi {' → '.join(ZERO_CHAIN)} đã về 0 theo tab KT. Bán mới dự thảo STT {seq}."
    by_doc[KEPT_DOC][2], by_doc[KEPT_DOC][8], by_doc[KEPT_DOC][9] = KEEP, KEPT_INVOICE, "Khớp hàng và số lượng"
    by_doc[KEPT_DOC][10] = "Giữ HĐ 1C26MUV/00005776. Tra cứu MISA xác nhận đúng mã đơn, 2 máy KA150P và đủ ba khoản."

    tabs["Cho xac minh"]["rows"] = [row for row in tabs["Cho xac minh"]["rows"] if row[0] not in {NEW_DOC, KEPT_DOC}]
    dec, sep = tabs["Tong hop"]["rows"][11], tabs["Tong hop"]["rows"][8]
    assert dec[0] == 12 and sep[0] == 9
    dec[2] += 1
    dec[5] -= 1
    dec[6] += pay
    sep[3] += 1
    sep[5] -= 1
    summary["actions"]["review"] -= 2
    summary["actions"]["new"] += 1
    summary["actions"]["keep"] += 1
    summary.update(new_invoices=seq, new_lines=len(sales["rows"]),
                   new_goods=summary["new_goods"] + goods, new_vat=summary["new_vat"] + vat,
                   new_pay=summary["new_pay"] + pay)
    assert summary["actions"]["review"] == len(tabs["Cho xac minh"]["rows"]) == 7
    decision = {
        "new": {"doc": NEW_DOC, "seq": seq, "row": excel_row, "money": [goods, vat, pay],
                "zero_chain": ZERO_CHAIN, "source": "Ktoan, dòng 30482, AD:AF = 0/0/0"},
        "keep": {"doc": KEPT_DOC, **PORTAL_DETAIL, "other_zero_chain": OTHER_ZERO_CHAIN},
    }
    summary["verified_zero_and_keep_20261109"] = decision
    prior = summary.get("ktoan_confirmed_keep_20261109", {})
    if prior:
        prior["resolved_later"] = [NEW_DOC, KEPT_DOC]

    (stage / "deliverables").mkdir(exist_ok=True)
    for kind in ("adjustment", "replacement", "return_adjustment"):
        filename = outputs[kind]["file"]
        shutil.copy2(OUT / "deliverables" / filename, stage / "deliverables" / filename)
    for name in ("prepared.json", "return-verification.json", "rounding-confirmation.json",
                 "service-quantity-confirmations.json", "unmapped-import-verification.json",
                 "adjustment-rule-2024-verification.json", "unmapped-gift-confirmation.json",
                 "LIEN_KET_BTL_BHH_2025.csv"):
        shutil.copy2(OUT / name, stage / name)
    link_headers, links = read_csv(LINKS)
    assert all(row[2] != NEW_DOC for row in links)
    write_csv(stage / LINKS, link_headers, links + [[sales["file"], seq, NEW_DOC, excel_row, 1, ZERO_CHAIN[0]]])
    write_csv(stage / LEDGER, headers, ledger)
    write_json(stage / "authoring.json", payload)
    write_json(stage / "invoice-sequences.json", sequences)
    write_json(stage / "final-summary.json", summary)
    write_json(stage / "native-additions.json", {"draft": draft, "decision": decision,
               "new_goods": summary["new_goods"], "new_vat": summary["new_vat"]})
    print("Staged sale 30260 and verified keep 00005776. Seven AMIS vouchers remain pending.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    args = parser.parse_args()
    stage = args.stage.resolve()
    assert stage.is_relative_to((OUT / "_build").resolve())
    stage_update(stage)
