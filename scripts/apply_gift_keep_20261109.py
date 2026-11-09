"""Cập nhật quyết định giữ hóa đơn biếu tặng trong báo cáo dự thảo."""

import hashlib
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs/import-2025-20261109"
POLICY = OUT / "unmapped-gift-confirmation.json"
KEEP_LABEL = "Kế toán chốt giữ lại — biếu tặng"


def read(path):
    return json.loads(path.read_text(encoding="utf-8"))


def save(name, data):
    path = OUT / name
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


def main():
    policy = read(POLICY)
    keys = set(policy["invoice_roots"])
    assert policy["accountant_decision"] == "keep" and policy["keep_out_of_import"]
    assert len(keys) == len(policy["invoice_roots"]) == 16

    payload = read(OUT / "authoring.json")
    plans = payload["zero_audit"] + payload["unmapped_zero_audit"]
    assert not any(keys.intersection(plan.get("keys", [plan["root"]])) for plan in plans)
    report = next(output for output in payload["outputs"] if output["kind"] == "report")
    tabs = {sheet["name"]: sheet for sheet in report["sheets"]}
    gifts = tabs["Bieu tang - KT check"]
    assert {row[0] for row in gifts["rows"]} == keys
    for row in gifts["rows"]:
        row[9] = "Kế toán chốt giữ lại. Theo dõi riêng nhóm biếu tặng."

    matched = set()
    for row in tabs["MI2025 chua map"]["rows"]:
        if row[0] in keys:
            assert row[10] is None
            row[1] = KEEP_LABEL
            row[2] = None
            row[9] = policy["accountant_decision_source"]
            matched.add(row[0])
    assert matched == keys

    summary = payload["summary"]
    unmapped = summary["unmapped_meinvoice"]
    unmapped["gift_kt_kept"] = len(keys)
    unmapped["classification"].pop("kt_review", None)
    unmapped["classification"]["kt_keep"] = len(keys)
    policy_path = str(POLICY.relative_to(ROOT))
    manifest = [entry for entry in payload["manifest"] if entry["file"] == policy_path]
    assert len(manifest) == 1
    manifest[0]["sha256"] = hashlib.sha256(POLICY.read_bytes()).hexdigest()

    save("gift-kt-check.json", {"headers": gifts["headers"], "rows": gifts["rows"]})
    save("final-summary.json", summary)
    save("authoring.json", payload)
    print(json.dumps({"gift_invoices_kept": len(keys), "overlap_with_zero_plans": 0}))


if __name__ == "__main__":
    main()
