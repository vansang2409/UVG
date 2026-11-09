"""Áp dụng xác nhận đổi sản phẩm cho hai chứng từ trong bộ dự thảo 2025."""

import csv
import hashlib
import json
from decimal import Decimal
from pathlib import Path

from import_2025_buyer_fields import apply_buyer_fields


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs/import-2025-20261109"
NEW = "Tạo import mới"
REVIEW = "Chờ xác minh"
CONFIRMED = {
    "BHHT2025004408": ("KA350LIGHTFULL", 540000, 1662778, 176222, "1C25TUV/00000386"),
    "BHHT2025011220": ("KM300FULL", 239000, 203593, 35407, "1C25TUV/00000459"),
}


def read_json(name):
    return json.loads((OUT / name).read_text(encoding="utf-8"))


def save_json(name, data):
    target = OUT / name
    temporary = target.with_suffix(".tmp")
    temporary.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    temporary.replace(target)


def save_csv(name, headers, rows):
    target = OUT / name
    temporary = target.with_suffix(".tmp")
    with temporary.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        writer.writerows(rows)
    temporary.replace(target)


def main():
    prepared = read_json("prepared.json")
    payload = read_json("authoring.json")
    sales = next(output for output in payload["outputs"] if output["kind"] == "new")
    report = next(output for output in payload["outputs"] if output["kind"] == "report")
    tabs = {sheet["name"]: sheet for sheet in report["sheets"]}
    decisions = {row["doc"]: row for row in prepared["decisions"]}
    sequences = read_json("invoice-sequences.json")
    audit = {}

    for code, (sku, excluded, goods, vat, root) in CONFIRMED.items():
        doc = prepared["docs"][code]
        product, exchange = doc["lines"]
        item, fee = product["v"], exchange["v"]
        assert (item[15], item[22], item[23], item[24], item[41], item[43]) == (
            sku, 1, goods, goods, vat, goods + vat
        )
        assert (fee[15], fee[24], fee[41], fee[43]) == ("PSC2", excluded, 0, excluded)
        assert doc["goods"] == goods + excluded and doc["vat"] == vat
        assert doc["pay"] == goods + vat + excluded
        chain = prepared["chains"][root]
        assert decisions[code]["roots"] == [root] and doc["date"] < "2025-07-01"
        assert not chain["issues"] and chain["balance"] == [0, 0, 0]
        assert all(line["rate"] == "8%" for line in prepared["invoice_detail"][root])
        audit[code] = {
            "sku": sku, "excluded_goods": excluded, "goods": goods, "vat": vat,
            "payment": goods + vat, "rate": 8, "product_source": product["source"],
            "psc2_source": exchange["source"], "old_root": root,
        }
        if code in sequences and any(row[18] == code for row in sales["rows"]):
            continue
        sequence = sequences.setdefault(code, len(sequences) + 1)
        sales["rows"].append([
            sequence, sales["rows"][0][1], doc["buyer"], item[13] or None,
            item[12] or None, item[14] or None, None, None, None, "TM/CK", 8, vat,
            str(item[7] or item[16]).strip(), item[19], 1, goods, goods,
            doc["order"], code, sku, doc["date"], product["source"], None,
        ])
        month = int(doc["date"][5:7])
        totals = tabs["Tong hop"]["rows"][month - 1]
        assert totals[0] == month and totals[5] > 0
        totals[2] += 1
        totals[5] -= 1
        totals[6] += goods + vat

    sale_count = len({row[0] for row in sales["rows"]})
    assert sale_count == len(sequences) and len(sequences) == len({row[18] for row in sales["rows"]})
    sales["invoices"] = sale_count
    count_label = f"{sale_count:,}".replace(",", ".")
    sales["notes"][0] = f"Gộp toàn bộ bán mới năm 2025 trong một sheet. STT liên tục 1–{count_label}. Cột Ngày chứng từ AMIS giữ thông tin kỳ gốc."
    sales["notes"][2] = "Hai chứng từ đổi sản phẩm bỏ dòng PSC2, giữ VAT dòng sản phẩm. Phiếu trả lại lập riêng sau khi có số HĐ thực tế."
    pending = tabs["Cho xac minh"]
    pending["rows"] = [row for row in pending["rows"] if row[0] not in audit]

    summary = payload["summary"]
    summary["actions"].update(new=sale_count, review=len(pending["rows"]))
    new_goods = int(sum(Decimal(str(row[16])) for row in sales["rows"]))
    new_vat = int(sum(Decimal(str(row[11] or 0)) for row in sales["rows"]))
    summary.update(new_invoices=sale_count, new_lines=len(sales["rows"]),
                   new_goods=new_goods, new_vat=new_vat, new_pay=new_goods + new_vat,
                   psc2_exchanges=audit)
    apply_buyer_fields(
        {"outputs": [{**sales, "rows": [row for row in sales["rows"] if row[18] in audit]}], "manifest": []},
        OUT, ROOT,
    )
    source = Path(__file__).resolve()
    source_name = str(source.relative_to(ROOT))
    payload["manifest"] = [entry for entry in payload["manifest"] if entry["file"] != source_name]
    payload["manifest"].append({"file": source_name, "sha256": hashlib.sha256(source.read_bytes()).hexdigest()})

    ledger_name = "DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv"
    with (OUT / ledger_name).open(encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file)
        headers = next(reader)
        ledger = list(reader)
    for row in ledger:
        if row[0] in audit:
            entry = audit[row[0]]
            assert row[2] in (NEW, REVIEW)
            row[2] = NEW
            row[-1] = f"Đổi sản phẩm. Trừ PSC2 {entry['excluded_goods']}đ, giữ VAT AMIS {entry['vat']}đ theo xác nhận người dùng."
    links_name = "LIEN_KET_IMPORT_MOI.csv"
    with (OUT / links_name).open(encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file)
        link_headers = next(reader)
        links = [row for row in reader if row[2] not in audit]
    for index, row in enumerate(sales["rows"], 10):
        if row[18] in audit:
            links.append([sales["file"], row[0], row[18], index, 1, audit[row[18]]["old_root"]])

    save_csv(ledger_name, headers, ledger)
    save_csv(links_name, link_headers, links)
    save_json("invoice-sequences.json", sequences)
    save_json("final-summary.json", summary)
    save_json("authoring.json", payload)
    print(json.dumps({"new_invoices": sale_count, "review": len(pending["rows"]), "psc2_exchanges": audit}))


if __name__ == "__main__":
    main()
