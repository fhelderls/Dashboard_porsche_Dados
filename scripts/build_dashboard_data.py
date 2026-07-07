#!/usr/bin/env python3
"""Regenerate the `const DATA = [...]` array embedded in index.html from the
canonical sanitized dataset (data/porsche_sales_sanitized.csv).

The dashboard only needs the sanitized subset of columns; the raw columns
required by SchemaPorshe.md's "no original columns removed" rule live in the
CSV, not in the HTML.
"""
import csv
import json
import re

CSV_PATH = "data/porsche_sales_sanitized.csv"
HTML_PATH = "index.html"


def main():
    with open(CSV_PATH, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    records = []
    for r in rows:
        price = float(r["SalesPriceSanitized"]) if r["SalesPriceSanitized"] != "INVALID" else 0.0
        mileage = int(r["VehicleMileageSanitized"]) if r["VehicleMileageSanitized"] != "INVALID" else 0
        records.append({
            "id": r["sale_id"],
            "date": r["SaleDateSanitized"],
            "model": r["PorscheModelSanitized"],
            "year": r["ModelYearSanitized"],
            "price": price,
            "mileage": mileage,
            "pay": r["PayMethodSanitized"],
            "city": r["CitySanitized"],
            "state": r["StateSanitized"],
            "status": r["DeliveryStatusSanitized"],
        })

    data_js = json.dumps(records, ensure_ascii=False, separators=(",", ":"))

    html = open(HTML_PATH, encoding="utf-8").read()
    new_html, n = re.subn(
        r"const DATA = \[.*?\];",
        "const DATA = " + data_js + ";",
        html,
        count=1,
        flags=re.DOTALL,
    )
    if n != 1:
        raise SystemExit("Could not find `const DATA = [...]` in index.html")
    open(HTML_PATH, "w", encoding="utf-8").write(new_html)
    print(f"Injected {len(records)} records into {HTML_PATH}")


if __name__ == "__main__":
    main()
