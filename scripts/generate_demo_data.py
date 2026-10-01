"""Generate a deterministic, synthetic sales dataset with known quality issues.

The dataset contains no private or university data. Run this script from the
project root to rebuild the CSV, the workbook-source JSON, and the expectations
used by automated tests.
"""

from __future__ import annotations

import csv
import json
import random
from copy import deepcopy
from datetime import date, timedelta
from pathlib import Path
from typing import Any


SEED = 20260921
BASE_ROW_COUNT = 120
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "sample_data"

HEADERS = [
    " Order ID ",
    "Customer Name",
    "EMAIL Address",
    "Order Date",
    "Product Category",
    "Product Name",
    "Quantity",
    "Unit Price EUR",
    "Discount Percent",
    "City ",
    "Sales Channel",
    "Order Status",
    "Internal Notes",
]

FIRST_NAMES = ["Maria", "Nikos", "Eleni", "Giorgos", "Sofia", "Dimitris"]
LAST_NAMES = ["Papadopoulou", "Nikolaou", "Georgiou", "Ioannou", "Pappas"]
CITIES = ["Athens", "Thessaloniki", "Patras", "Heraklion", "Larissa"]
PRODUCTS = [
    ("Electronics", "USB-C Hub", 34.90),
    ("Electronics", "Wireless Mouse", 24.50),
    ("Office", "Notebook Set", 12.80),
    ("Office", "Desk Lamp", 29.90),
    ("Accessories", "Laptop Sleeve", 21.40),
    ("Accessories", "Cable Organizer", 8.60),
]
CHANNELS = ["Online", "Retail", "Marketplace"]
STATUSES = ["Completed", "Processing", "Cancelled"]


def _formatted_date(order_date: date, index: int) -> str:
    formats = (
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d %b %Y",
        "%d.%m.%Y",
    )
    return order_date.strftime(formats[index % len(formats)])


def build_rows() -> list[dict[str, Any]]:
    """Create rows whose defects are deliberate and reproducible."""
    rng = random.Random(SEED)
    rows: list[dict[str, Any]] = []
    start_date = date(2026, 1, 5)

    for index in range(BASE_ROW_COUNT):
        first = rng.choice(FIRST_NAMES)
        last = rng.choice(LAST_NAMES)
        category, product, price = rng.choice(PRODUCTS)
        order_date = start_date + timedelta(days=rng.randint(0, 170))
        customer_name = f"{first} {last}"
        rows.append(
            {
                " Order ID ": f"ORD-{1001 + index}",
                "Customer Name": customer_name,
                "EMAIL Address": f"{first}.{last}{index}@example.com".lower(),
                "Order Date": _formatted_date(order_date, index),
                "Product Category": category,
                "Product Name": product,
                "Quantity": rng.randint(1, 6),
                "Unit Price EUR": price,
                "Discount Percent": rng.choice([0, 0, 5, 10, 15]),
                "City ": rng.choice(CITIES),
                "Sales Channel": rng.choice(CHANNELS),
                "Order Status": rng.choice(STATUSES),
                "Internal Notes": "",
            }
        )

    # Missing values. Duplicate source indices below deliberately avoid these.
    for index in (4, 44, 84):
        rows[index]["Customer Name"] = ""
    for index in (5, 25, 45, 65, 85):
        rows[index]["EMAIL Address"] = ""
    for index in (6, 36, 66, 96):
        rows[index]["City "] = ""
    for index in (7, 47, 87):
        rows[index]["Order Status"] = ""

    # Outer whitespace in cells.
    for index in (1, 11, 21, 31, 41):
        rows[index]["Customer Name"] = f"  {rows[index]['Customer Name']} "
    for index in (2, 12, 22, 32, 42, 52, 62, 72):
        rows[index]["EMAIL Address"] = f" {rows[index]['EMAIL Address']}  "
    for index in (3, 13, 23, 33, 43, 53, 63, 73, 83, 93):
        rows[index]["City "] = f" {rows[index]['City ']} "

    # Case variants that should be proposed, not silently merged.
    rows[8]["City "] = "ATHENS"
    rows[18]["City "] = "athens"
    rows[28]["Product Category"] = "electronics"
    rows[38]["Product Category"] = "ELECTRONICS"
    rows[48]["Sales Channel"] = "online"
    rows[58]["Order Status"] = "COMPLETED"

    # Numeric-looking text, invalid text, negative value, and a valid outlier.
    rows[9]["Quantity"] = " 4 "
    rows[19]["Quantity"] = "3"
    rows[29]["Quantity"] = "unknown"
    rows[39]["Quantity"] = -2
    rows[49]["Quantity"] = 75
    rows[59]["Unit Price EUR"] = " 34.90 "
    rows[69]["Unit Price EUR"] = "29,90"
    rows[79]["Unit Price EUR"] = "not available"
    rows[89]["Discount Percent"] = "10%"

    # Invalid and ambiguous dates remain strings for later flagging.
    rows[10]["Order Date"] = "31/02/2026"
    rows[20]["Order Date"] = "not recorded"
    rows[30]["Order Date"] = "03/04/2026"
    rows[40]["Order Date"] = "04/05/2026"
    rows[50]["Order Date"] = "05/06/2026"

    duplicate_source_indices = (0, 14, 24, 34, 54, 64, 74, 94)
    rows.extend(deepcopy(rows[index]) for index in duplicate_source_indices)

    blank_row = {header: "" for header in HEADERS}
    for position in (15, 57, 103):
        rows.insert(position, deepcopy(blank_row))

    return rows


def build_expectations(rows: list[dict[str, Any]]) -> dict[str, Any]:
    """Describe the intentional defects independently of future app logic."""
    return {
        "seed": SEED,
        "base_business_rows": BASE_ROW_COUNT,
        "total_rows_including_duplicates_and_blanks": len(rows),
        "total_columns": len(HEADERS),
        "exact_duplicate_business_rows": 8,
        "fully_blank_rows": 3,
        "fully_blank_columns": ["Internal Notes"],
        "missing_values_in_base_rows": {
            "Customer Name": 3,
            "EMAIL Address": 5,
            "City ": 4,
            "Order Status": 3,
        },
        "outer_whitespace_cells_in_base_rows": 25,
        "known_case_variant_columns": [
            "City ",
            "Product Category",
            "Sales Channel",
            "Order Status",
        ],
        "numeric_review_values": [" 4 ", "3", "unknown", -2, 75, " 34.90 ", "29,90", "not available", "10%"],
        "date_review_values": [
            "31/02/2026",
            "not recorded",
            "03/04/2026",
            "04/05/2026",
            "05/06/2026",
        ],
        "notes": [
            "Duplicate count excludes fully blank rows.",
            "Missing-value counts refer to the 120 base business rows before duplicate insertion.",
            "Ambiguous dates and suspicious numeric values must be flagged rather than silently changed.",
        ],
    }


def write_csv(rows: list[dict[str, Any]], path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS)
        writer.writeheader()
        writer.writerows(rows)


def write_json(data: Any, path: Path) -> None:
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    rows = build_rows()
    write_csv(rows, DATA_DIR / "messy_sales_data.csv")
    write_json({"headers": HEADERS, "rows": rows}, DATA_DIR / "messy_sales_data_source.json")
    write_json(build_expectations(rows), DATA_DIR / "demo_data_expectations.json")
    print(f"Generated {len(rows)} rows and {len(HEADERS)} columns in {DATA_DIR}")


if __name__ == "__main__":
    main()
