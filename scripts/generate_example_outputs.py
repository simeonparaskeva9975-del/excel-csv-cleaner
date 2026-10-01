"""Generate portfolio-ready example outputs from the demo dataset."""

from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


from cleaner_app.analysis import analyze_dataset
from cleaner_app.cleaning import FOUNDATION_ACTIONS, clean_dataset
from cleaner_app.file_io import LoadedDataset, load_dataset
from cleaner_app.models import CleaningAction
from cleaner_app.reporting import (
    export_cleaned_csv,
    export_excel_report,
)


DEMO_PATH = (
    PROJECT_ROOT
    / "sample_data"
    / "messy_sales_data.xlsx"
)

OUTPUT_DIRECTORY = PROJECT_ROOT / "output"

EXCEL_OUTPUT_PATH = (
    OUTPUT_DIRECTORY
    / "messy_sales_data_cleaned_report.xlsx"
)

CSV_OUTPUT_PATH = (
    OUTPUT_DIRECTORY
    / "messy_sales_data_cleaned.csv"
)


def main() -> None:
    """Run the complete demo cleaning and reporting workflow."""

    loaded_dataset = load_dataset(DEMO_PATH)
    before_profile = analyze_dataset(loaded_dataset)

    all_actions = FOUNDATION_ACTIONS | {
        CleaningAction.STANDARDIZE_TEXT_CASE,
        CleaningAction.CONVERT_NUMERIC,
        CleaningAction.CONVERT_DATES,
    }

    cleaning_result = clean_dataset(
        loaded_dataset,
        actions=all_actions,
        text_case_columns=(
            "Product Category",
            "City ",
            "Sales Channel",
            "Order Status",
        ),
        numeric_columns=(
            "Quantity",
            "Unit Price EUR",
            "Discount Percent",
        ),
        date_columns=("Order Date",),
    )

    cleaned_dataset = LoadedDataset(
        data=cleaning_result.data,
        source_path=EXCEL_OUTPUT_PATH,
        file_type="Excel",
        file_size_bytes=0,
    )

    after_profile = analyze_dataset(cleaned_dataset)

    exported_excel = export_excel_report(
        EXCEL_OUTPUT_PATH,
        cleaning_result,
        before_profile,
        after_profile,
    )

    exported_csv = export_cleaned_csv(
        CSV_OUTPUT_PATH,
        cleaning_result,
    )

    print("Example outputs generated successfully.")
    print(
        f"Rows: {before_profile.row_count} "
        f"-> {after_profile.row_count}"
    )
    print(
        f"Columns: {before_profile.column_count} "
        f"-> {after_profile.column_count}"
    )
    print(
        f"Cleaning log entries: "
        f"{len(cleaning_result.log_entries)}"
    )
    print(f"Excel report: {exported_excel}")
    print(f"Cleaned CSV: {exported_csv}")


if __name__ == "__main__":
    main()