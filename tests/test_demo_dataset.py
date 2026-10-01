"""Contract tests for the deliberately messy demonstration dataset."""

from __future__ import annotations

import json
import unittest
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "sample_data"
CSV_PATH = DATA_DIR / "messy_sales_data.csv"
XLSX_PATH = DATA_DIR / "messy_sales_data.xlsx"
EXPECTATIONS_PATH = DATA_DIR / "demo_data_expectations.json"


def _expected() -> dict:
    return json.loads(EXPECTATIONS_PATH.read_text(encoding="utf-8"))


def _read_csv() -> pd.DataFrame:
    return pd.read_csv(CSV_PATH, dtype=object, keep_default_na=False, skip_blank_lines=False)


class DemoDatasetTests(unittest.TestCase):
    def test_demo_files_exist(self) -> None:
        self.assertTrue(CSV_PATH.exists())
        self.assertTrue(XLSX_PATH.exists())
        self.assertTrue(EXPECTATIONS_PATH.exists())

    def test_dimensions_and_blank_structures_are_exact(self) -> None:
        frame = _read_csv()
        expected = _expected()

        self.assertEqual(
            frame.shape,
            (
                expected["total_rows_including_duplicates_and_blanks"],
                expected["total_columns"],
            ),
        )
        self.assertEqual(int(frame.eq("").all(axis=1).sum()), expected["fully_blank_rows"])
        blank_columns = frame.columns[frame.eq("").all(axis=0)].tolist()
        self.assertEqual(blank_columns, expected["fully_blank_columns"])

    def test_exact_duplicate_business_rows_are_present(self) -> None:
        frame = _read_csv()
        expected = _expected()
        nonblank = frame.loc[~frame.eq("").all(axis=1)]
        self.assertEqual(
            int(nonblank.duplicated(keep="first").sum()),
            expected["exact_duplicate_business_rows"],
        )

    def test_expected_missing_values_exist_in_base_business_rows(self) -> None:
        frame = _read_csv()
        nonblank = frame.loc[~frame.eq("").all(axis=1)]
        unique_orders = nonblank.drop_duplicates(subset=[" Order ID "], keep="first")

        for column, count in _expected()["missing_values_in_base_rows"].items():
            self.assertEqual(int(unique_orders[column].eq("").sum()), count)

    def test_case_whitespace_numeric_and_date_examples_are_present(self) -> None:
        frame = _read_csv()
        values = set(frame.astype(str).to_numpy().ravel())

        self.assertTrue({"ATHENS", "athens"}.issubset(values))
        self.assertTrue(any(value != value.strip() for value in values))
        self.assertTrue({"unknown", "-2", "75", "not available", "10%"}.issubset(values))
        self.assertTrue(set(_expected()["date_review_values"]).issubset(values))

    def test_excel_and_csv_contain_the_same_raw_values(self) -> None:
        csv_frame = _read_csv().fillna("").astype(str)
        excel_frame = (
            pd.read_excel(XLSX_PATH, dtype=object, keep_default_na=False)
            .fillna("")
            .astype(str)
        )
        pd.testing.assert_frame_equal(csv_frame, excel_frame, check_dtype=False)


if __name__ == "__main__":
    unittest.main()
