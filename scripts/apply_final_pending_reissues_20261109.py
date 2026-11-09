"""Stage the five authorized reissues and protect the sixth order's candidate."""

import argparse
import shutil
from collections import Counter
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl.utils.datetime import to_excel

from apply_continued_reissue_20261109 import (
    KEEP, LEDGER, LINKS, NEW, OUT, REVIEW, ROOT, invoice_day, load_delivered_payload,
    money_round, number, read_csv, read_json, read_replacement_sources, write_csv, write_json,
)
from import_2025_buyer_fields import apply_buyer_fields


REPLACEMENTS = {
    "BHHT2025038444": "1C25MUV/00027821",
    "BHHT2025038427": "1C25MUV/00027825",
    "BHHT2025038199": "1C25MUV/00027826",
    "BHHT2025029713": "1C25MUV/00017875",
}
EXCHANGE, F0, F1 = "BHHT2025027227", "1C25MUV/00013149", "1C25MUV/00017218"
PENDING, WRONG, CANDIDATE = "BHHT2025027490", "1C25MUV/00012392", "1C25MUV/00012515"
CANDIDATE_SEQ = 1248
AMIS_ORDER = "580369342577804306"
CANDIDATE_NOTE = (
    f"Mã đơn {AMIS_ORDER}, NB dòng 15512. {CANDIDATE} khớp người mua, 8 KH200 và "
    f"4.659.259đ tiền hàng + 372.741đ VAT = 5.032.000đ. Mã đơn trống trên HĐ gốc. "
    f"Tạm rút {CANDIDATE} khỏi lô thay thế; chờ chốt liên kết."
)


def stage_update(stage):
    payload = load_delivered_payload()
    prepared = read_json("prepared.json")
    ledger = read_json("source-index.json")["ledger"]
    outputs = {entry["kind"]: entry for entry in payload["outputs"]}
    summary = payload["summary"]
    tabs = {tab["name"]: tab["rows"] for tab in outputs["report"]["sheets"]}
    sale, replacement, adjustment = (outputs[k] for k in ("new", "replacement", "adjustment"))
    headers, decisions = read_csv(LEDGER)
    by_doc = {row[0]: row for row in decisions}
    by_root = {row[0]: row for row in tabs["Xu ly chuoi cu"]}
    sale_seq = read_json("invoice-sequences.json")
    unmapped_seq = read_json("unmapped-import-sequences.json")
    selected = {**REPLACEMENTS, EXCHANGE: F0}
    keys = set(selected.values()) | {F1, CANDIDATE}
    source = read_replacement_sources(keys, ledger)
    company = {entry["key"] for entry in prepared["company_rows"]}
    assert not keys & company
    assert not (keys - {CANDIDATE}) & {x["root"] for x in payload["zero_audit"] + payload["unmapped_zero_audit"]}
    assert set(selected).isdisjoint(row[25] for row in outputs["return_adjustment"]["rows"])
    assert all(code not in sale_seq and by_doc[code][2] == REVIEW and by_doc[code][8] == root
               for code, root in selected.items())
    assert by_doc[PENDING][2] == REVIEW and by_doc[PENDING][8] == WRONG
    assert prepared["docs"][PENDING]["order"] == AMIS_ORDER
    assert ledger[CANDIDATE]["order"] == "" and ledger[CANDIDATE]["date"] == "2025-09-15"
    assert not [row for row in decisions if row[8] == CANDIDATE]
    assert len(replacement["rows"]) == 1986 and len(adjustment["rows"]) == 33
    assert replacement["rows"][CANDIDATE_SEQ - 1][25] == CANDIDATE
    assert unmapped_seq.pop(CANDIDATE) == {"kind": "replacement", "sequence": CANDIDATE_SEQ}
    pending_audit = [plan for plan in payload["unmapped_zero_audit"] if plan["root"] == CANDIDATE]
    assert len(pending_audit) == 1 and pending_audit[0]["keys"] == [CANDIDATE]
    payload["unmapped_zero_audit"].remove(pending_audit[0])
    candidate_money = [int(number(source[CANDIDATE][f])) for f in "PQR"]
    assert candidate_money == [4659259, 372741, 5032000] == pending_audit[0]["before"]
    candidate = next(row for row in tabs["MI2025 chua map"] if row[0] == CANDIDATE)
    assert candidate[2] == "replacement" and candidate[10] == CANDIDATE_SEQ
    candidate[1:3] = ["Tạm giữ, chờ xác minh liên kết AMIS", None]
    candidate[9:11] = [CANDIDATE_NOTE, None]
    by_doc[PENDING][-1] = CANDIDATE_NOTE
    by_root[WRONG][7] = CANDIDATE_NOTE
    summary["unmapped_meinvoice"]["added"]["replacement"] -= 1
    summary["unmapped_meinvoice"]["zeroed_payment"] -= candidate_money[2]

    current = int(to_excel(date(2026, 11, 9)))
    sale_rows, replacement_rows, adjustment_rows, records, links = [], [], [], [], []
    for code, root in selected.items():
        doc = prepared["docs"][code]
        chain = prepared["chains"][root]
        identity = source[root]
        assert not chain["issues"] and identity["E"] and (identity.get("H") or identity.get("L"))
        day = invoice_day(identity["D"])
        symbol, invoice = root.split("/")
        before = [int(v) for v in chain["balance"]]
        f0_money = [int(number(identity[field])) for field in "PQR"]
        assert before[0] + before[1] == before[2] and ledger[root]["cqt_code_present"]
        goods, vat, pay = [doc[k] for k in ("goods", "vat", "pay")]
        lines = doc["lines"]
        new_seq = len(sale_seq) + 1
        sale_seq[code] = new_seq
        first_row = 10 + len(sale["rows"]) + len(sale_rows)
        if code == EXCHANGE:
            assert chain["keys"] == chain["active"] == [F0, F1] and ledger[F1]["order"] == doc["order"]
            assert ledger[F1]["status"] == "Hoá đơn điều chỉnh"
            assert f0_money == [600411, 48033, 648444]
            f1_money = [int(number(source[F1][field])) for field in "PQR"]
            assert f1_money == [-100926, -8074, -109000]
            assert [a + b for a, b in zip(f0_money, f1_money)] == before
            product, fee = lines
            assert product["v"][15] == "KH200FULL" and product["v"][24] == 491411
            assert (fee["v"][15], fee["v"][24], fee["v"][41]) == ("PSC2", 109000, 0)
            goods, vat = product["v"][24], int(money_round(number(product["v"][24]) * Decimal("0.08")))
            pay, lines = goods + vat, [product]
            summary["psc2_exchanges"][code] = {
                "excluded_goods": 109000, "goods": goods, "vat": vat, "payment": pay, "rate": 8,
                "product_sources": [product["source"]], "psc2_source": fee["source"],
                "old_root": F0, "rate_policy": "recalculate", "source_vat": doc["vat"],
            }
            common = [None, current, identity.get("H") or None, identity.get("I") or None,
                      identity.get("J") or None, identity.get("L") or None, None,
                      identity.get("M") or "TM/CK", None, symbol, invoice, int(to_excel(day)), identity["E"]]
            item, = prepared["invoice_detail"][root]
            assert [int(number(item[f])) for f in ("qty", "price", "goods", "vat")] == [1, 600411, 600411, 48033]
            undo_reason = f"Điều chỉnh tăng để đảo khoản đã điều chỉnh trên HĐ {F1}; tham chiếu HĐ gốc {F0}"
            reverse_reason = f"Điều chỉnh giảm toàn bộ hàng bán trên HĐ số {invoice} ngày {day:%d/%m/%Y} để đưa chuỗi về 0"
            adjustment_rows = [
                [*common, undo_reason, undo_reason, None, None, None, 100926, 100926, 8, 8074, 109000,
                 None, code, root, *f1_money],
                [*common, reverse_reason, item["name"], item["unit"], -1, 600411, -600411, -600411, 8,
                 -48033, -648444, None, code, root, *f0_money],
            ]
            adj_seqs = [adjustment["invoices"] + i for i in (1, 2)]
            audit_plan = payload["adjustment_rule_2024"]
            for row, seq, role, reference, src in zip(
                adjustment_rows, adj_seqs, ("Đảo khoản điều chỉnh cũ", "Giảm toàn bộ chi tiết F0"),
                (F1, F0), (ledger[F1]["source"], item["source"]),
            ):
                row[0] = seq
                audit_plan["rows"].append({"seq": seq, "root": root, "source_invoice": reference,
                                           "role": role, "first": True, "source": src})
                tabs["Quy cach DC2024"].append([root, None, seq, role, reference, row[2], row[5],
                                               row[19], row[21], row[22], str(src), "Tham chiếu F0; tên chép riêng từ F0"])
            audit_plan["chains"].append({"root": root, "old_sequence": None, "sequences": adj_seqs,
                                         "original_reversal_sequence": adj_seqs[-1], "before": before, "origin": "AMIS"})
            audit_plan["invoices"] += len(adjustment_rows)
            audit_plan["line_count"] += len(adjustment_rows)
            seq = adj_seqs[-1]
            kind, action = "adjustment", "Điều chỉnh về 0"
            note = f"Đảo F1 STT {adj_seqs[0]}, giảm F0 STT {adj_seqs[1]}; bán mới STT {new_seq}. Trừ PSC2 109.000đ; VAT 8% là 39.313đ."
        else:
            assert chain["keys"] == chain["active"] == [root] and before == f0_money
            assert ledger[root]["status"] == "Hoá đơn mới" and ledger[root]["order"] == doc["order"]
            assert {item["rate"] for item in prepared["invoice_detail"][root]} == {"8%"}
            seq = CANDIDATE_SEQ if not replacement_rows else len(replacement["rows"]) + len(replacement_rows)
            reason = f"Thay thế HĐ {root} ngày {day:%d/%m/%Y} về 0. Lập lại theo {code}."
            replacement_rows.append([
                seq, current, identity.get("H") or None, identity.get("J") or None, identity.get("I") or None,
                identity.get("L") or None, None, None, None, "TM/CK", None, symbol, invoice, int(to_excel(day)),
                identity["E"], reason, None, 0, 0, 0, 8, 0, reason, 1, code, root,
            ])
            kind, action = "replacement", "Thay thế về 0"
            note = f"Thay thế F0 về 0 STT {seq}; bán mới STT {new_seq}. Đã xác nhận xử lý chênh VAT trên HĐ cũ."

        assert [sum(line["v"][24] for line in lines), goods + vat] == [goods, pay]
        for i, line in enumerate(lines):
            item = line["v"]
            assert abs(number(item[22]) * number(item[23]) - number(item[24])) <= 2
            sale_rows.append([
                new_seq, current, doc["buyer"], item[13] or None, item[12] or None, item[14] or None, None,
                None, None, "TM/CK", 8, vat if i == 0 else None, str(item[7] or item[16]).strip(), item[19],
                item[22], item[23], item[24], doc["order"], code, item[15], doc["date"], line["source"], None,
            ])
        links.append([sale["file"], new_seq, code, first_row, len(lines), root])
        by_doc[code][2], by_doc[code][-1] = NEW, note
        by_root[root][1], by_root[root][7] = action, note
        payload["zero_audit"].append({"root": root, "kind": kind, "sequence": seq, "docs": [code],
                                      "keys": chain["keys"], "before": before, "change": [-v for v in before],
                                      "after": [0, 0, 0], "identity_source": ledger[root]["source"]})
        records.append({"root": root, "code": code, "kind": kind, "zero_seqs": adj_seqs if code == EXCHANGE else [seq],
                        "before": before, "new_seq": new_seq, "new": [goods, vat, pay], "note": note})

    apply_buyer_fields({"outputs": [{**sale, "rows": sale_rows}, {**replacement, "rows": replacement_rows}],
                        "manifest": []}, OUT, ROOT)
    sale["rows"].extend(sale_rows)
    sale["invoices"] = len(sale_seq)
    sale["notes"][0] = "Gộp toàn bộ bán mới năm 2025. STT liên tục 1–30.265. Cột Ngày chứng từ AMIS giữ thông tin kỳ gốc."
    sale["notes"][2] = "Bốn chứng từ trừ PSC2 rồi tính lại VAT 8%. Hai chứng từ PSC2 đã chốt trước giữ VAT AMIS. Phiếu trả hàng lập riêng."
    replacement["rows"][CANDIDATE_SEQ - 1] = replacement_rows[0]
    replacement["rows"].extend(replacement_rows[1:])
    replacement["invoices"] = len(replacement["rows"])
    adjustment["rows"].extend(adjustment_rows)
    adjustment["invoices"] += len(adjustment_rows)
    adjustment["notes"][1] = "19 chuỗi /27 HĐ /35 dòng: 8 HĐ đảo khoản cũ và 19 HĐ giảm chi tiết F0; 13 khoản cũ 0đ không tạo thêm."
    adjustment["notes"].append("Ngày dự thảo09/11/2026; dữ liệu kiểm đến01/11/2026. CộtY:AC chỉ tra cứu; không ghép vào trường import.")

    tabs["Cho xac minh"][:] = [row for row in tabs["Cho xac minh"] if row[0] not in selected]
    assert len(tabs["Cho xac minh"]) == 1 and tabs["Cho xac minh"][0][0] == PENDING
    tabs["Cho xac minh"][0][-1] = CANDIDATE_NOTE
    counts = Counter((int(row[1][5:7]), row[2]) for row in decisions)
    monthly_pay = Counter()
    for row in sale["rows"]:
        monthly_pay[int(row[20][5:7])] += int(number(row[16]) + number(row[11]))
    for row in tabs["Tong hop"]:
        row[2:6] = [counts[row[0], label] for label in (NEW, KEEP, "Loại theo HD CTY 2025", REVIEW)]
        assert sum(row[2:6]) == row[1]
        row[6] = monthly_pay[row[0]]
    summary["actions"].update(new=sale["invoices"], review=len(tabs["Cho xac minh"]))
    summary.update(new_invoices=sale["invoices"], new_lines=len(sale["rows"]), adjustment=adjustment["invoices"],
                   adjustment_chains=len(payload["adjustment_rule_2024"]["chains"]), replacement=replacement["invoices"])
    summary["new_goods"] = sum(int(number(row[16])) for row in sale["rows"])
    summary["new_vat"] = sum(int(number(row[11])) for row in sale["rows"])
    summary["new_pay"] = summary["new_goods"] + summary["new_vat"]
    summary["final_pending_reissues_20261109"] = {
        "records": records,
        "pending": {"doc": PENDING, "order": AMIS_ORDER, "sheet_link": WRONG, "candidate": CANDIDATE,
                    "removed_replacement_seq": CANDIDATE_SEQ, "candidate_money": candidate_money,
                    "source": ledger[CANDIDATE]["source"], "note": CANDIDATE_NOTE},
    }
    assert [row[0] for row in replacement["rows"]] == list(range(1, 1990))
    assert (sale["invoices"], len(sale["rows"]), adjustment["invoices"], len(adjustment["rows"])) == (30265, 35047, 27, 35)

    (stage / "deliverables").mkdir(parents=True, exist_ok=True)
    name = outputs["return_adjustment"]["file"]
    shutil.copy2(OUT / "deliverables" / name, stage / "deliverables" / name)
    for name in ("prepared.json", "return-verification.json", "rounding-confirmation.json",
                 "service-quantity-confirmations.json", "unmapped-import-verification.json",
                 "adjustment-rule-2024-verification.json", "unmapped-gift-confirmation.json", "LIEN_KET_BTL_BHH_2025.csv"):
        shutil.copy2(OUT / name, stage / name)
    link_headers, old_links = read_csv(LINKS)
    write_csv(stage / LINKS, link_headers, old_links + links)
    write_csv(stage / LEDGER, headers, decisions)
    write_json(stage / "authoring.json", payload)
    write_json(stage / "final-summary.json", summary)
    write_json(stage / "invoice-sequences.json", sale_seq)
    write_json(stage / "unmapped-import-sequences.json", unmapped_seq)
    write_json(stage / "adjustment-rule-2024-plan.json", payload["adjustment_rule_2024"])
    write_json(stage / "native-update.json", {
        **summary["final_pending_reissues_20261109"], "sale": sale_rows,
        "replacement": replacement_rows, "adjustment": adjustment_rows,
        "sale_notes": sale["notes"], "adjustment_notes": adjustment["notes"],
        "new_goods": summary["new_goods"], "new_vat": summary["new_vat"],
    })
    print("Staged 5 new invoices, 4 replacements, 2 adjustments; protected 1 candidate.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    args = parser.parse_args()
    stage = args.stage.resolve()
    assert stage.is_relative_to((OUT / "_build").resolve())
    stage_update(stage)
