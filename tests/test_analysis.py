"""Tests for non-destructive data-quality analysis."""

from __future__ import annotations

import unittest
from pathlib import Path

import pandas as pd

from cleaner_app.analysis import analyze_dataset, is_blank_value
from cleaner_app.file_io import LoadedDataset, load_dataset


PROJECT_ROOT = Path(__file__).resolve().parents[1]
XLSX_PATH = PROJECT_ROOT / "sample_data" / "messy_sales_data.xlsx"


class AnalysisTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.loaded_dataset = load_dataset(XLSX_PATH)

    def test_numeric_columns_are_detected_conservatively(self) -> None:
        profile = analyze_dataset(self.loaded_dataset)

        self.assertEqual(
            set(profile.numeric_findings),
            {
                "Quantity",
                "Unit Price EUR",
                "Discount Percent",
            },
        )

        quantity = profile.numeric_findings["Quantity"]
        self.assertEqual(quantity.nonblank_count, 128)
        self.assertEqual(quantity.numeric_value_count, 125)
        self.assertEqual(quantity.numeric_text_count, 2)
        self.assertEqual(quantity.invalid_count, 1)
        self.assertEqual(quantity.invalid_examples, ("unknown",))
        self.assertAlmostEqual(
            quantity.conversion_ratio,
            127 / 128,
        )

        unit_price = profile.numeric_findings["Unit Price EUR"]
        self.assertEqual(unit_price.nonblank_count, 128)
        self.assertEqual(unit_price.numeric_value_count, 125)
        self.assertEqual(unit_price.numeric_text_count, 1)
        self.assertEqual(unit_price.invalid_count, 2)
        self.assertEqual(
            unit_price.invalid_examples,
            ("29,90", "not available"),
        )
        self.assertAlmostEqual(
            unit_price.conversion_ratio,
            126 / 128,
        )

        discount = profile.numeric_findings["Discount Percent"]
        self.assertEqual(discount.nonblank_count, 128)
        self.assertEqual(discount.numeric_value_count, 127)
        self.assertEqual(discount.numeric_text_count, 0)
        self.assertEqual(discount.invalid_count, 1)
        self.assertEqual(discount.invalid_examples, ("10%",))

        numeric_issues = [
            issue
            for issue in profile.issues
            if issue.issue_type in {
                "mixed_numeric_values",
                "numeric_text",
            }
        ]

        self.assertTrue(numeric_issues)
        self.assertTrue(
            all(not issue.safe_to_fix for issue in numeric_issues)
        )

    def test_date_column_is_detected_conservatively(self) -> None:
        profile = analyze_dataset(self.loaded_dataset)

        self.assertEqual(
            set(profile.date_findings),
            {"Order Date"},
        )

        finding = profile.date_findings["Order Date"]

        self.assertEqual(finding.nonblank_count, 128)
        self.assertEqual(finding.date_value_count, 0)
        self.assertEqual(finding.recognized_text_count, 111)
        self.assertEqual(finding.ambiguous_text_count, 15)
        self.assertEqual(finding.invalid_count, 2)

        self.assertEqual(
            finding.detected_formats,
            (
                "dd/mm/yyyy",
                "yyyy-mm-dd",
                "dd Mon yyyy",
                "dd.mm.yyyy",
                "dd/mm/yyyy or mm/dd/yyyy",
            ),
        )

        self.assertEqual(
            finding.ambiguous_examples,
            (
                "07/01/2026",
                "12/02/2026",
                "04/02/2026",
            ),
        )

        self.assertEqual(
            finding.invalid_examples,
            (
                "31/02/2026",
                "not recorded",
            ),
        )

        date_issues = [
            issue
            for issue in profile.issues
            if issue.issue_type in {
                "mixed_date_formats",
                "ambiguous_dates",
                "invalid_date_values",
            }
        ]

        self.assertEqual(len(date_issues), 3)
        self.assertTrue(
            all(not issue.safe_to_fix for issue in date_issues)
        )
    def test_blank_value_detection(self) -> None:
        self.assertTrue(is_blank_value(None))
        self.assertTrue(is_blank_value(float("nan")))
        self.assertTrue(is_blank_value(""))
        self.assertTrue(is_blank_value("   "))

        self.assertFalse(is_blank_value(0))
        self.assertFalse(is_blank_value(False))
        self.assertFalse(is_blank_value("0"))
        self.assertFalse(is_blank_value("Athens"))

    def test_case_variants_are_detected_and_not_marked_safe(self) -> None:
        profile = analyze_dataset(self.loaded_dataset)

        self.assertEqual(
            set(profile.case_variants_by_column),
            {
                "Product Category",
                "City ",
                "Sales Channel",
                "Order Status",
            },
        )

        self.assertEqual(
            profile.case_variants_by_column["Product Category"]["electronics"],
            ("ELECTRONICS", "Electronics", "electronics"),
        )
        self.assertEqual(
            profile.case_variants_by_column["City "]["athens"],
            ("ATHENS", "Athens", "athens"),
        )
        self.assertEqual(
            profile.case_variants_by_column["Sales Channel"]["online"],
            ("Online", "online"),
        )
        self.assertEqual(
            profile.case_variants_by_column["Order Status"]["completed"],
            ("COMPLETED", "Completed"),
        )

        case_issues = [
            issue
            for issue in profile.issues
            if issue.issue_type == "case_variants"
        ]

        self.assertEqual(len(case_issues), 4)
        self.assertTrue(
            all(not issue.safe_to_fix for issue in case_issues)
        )
    
    def test_known_demo_profile(self) -> None:
        profile = analyze_dataset(self.loaded_dataset)

        self.assertEqual(profile.source_name, "messy_sales_data.xlsx")
        self.assertEqual(profile.row_count, 131)
        self.assertEqual(profile.column_count, 13)
        self.assertEqual(profile.duplicate_rows, 8)
        self.assertEqual(profile.blank_rows, 3)
        self.assertEqual(profile.blank_columns, ("Internal Notes",))

    def test_whitespace_and_column_name_findings(self) -> None:
        profile = analyze_dataset(self.loaded_dataset)

        self.assertEqual(
            profile.whitespace_cells_by_column,
            {
                "Customer Name": 5,
                "EMAIL Address": 8,
                "Quantity": 1,
                "Unit Price EUR": 1,
                "City ": 10,
            },
        )
        self.assertEqual(
            sum(profile.whitespace_cells_by_column.values()),
            25,
        )

        self.assertEqual(
            profile.column_name_suggestions[" Order ID "],
            "order_id",
        )
        self.assertEqual(
            profile.column_name_suggestions["EMAIL Address"],
            "email_address",
        )
        self.assertEqual(
            profile.column_name_suggestions["City "],
            "city",
        )
        self.assertEqual(
            len(profile.column_name_suggestions),
            13,
        )

    def test_missing_counts_exclude_fully_blank_rows(self) -> None:
        profile = analyze_dataset(self.loaded_dataset)

        self.assertEqual(profile.missing_by_column["Customer Name"], 3)
        self.assertEqual(profile.missing_by_column["EMAIL Address"], 5)
        self.assertEqual(profile.missing_by_column["City "], 4)
        self.assertEqual(profile.missing_by_column["Order Status"], 3)

        # There are 128 nonblank business rows. The three fully blank
        # rows are reported separately and do not inflate this count.
        self.assertEqual(profile.missing_by_column["Internal Notes"], 128)

    def test_safe_and_uncertain_issues_are_distinguished(self) -> None:
        profile = analyze_dataset(self.loaded_dataset)

        safe_issue_types = {
            issue.issue_type
            for issue in profile.issues
            if issue.safe_to_fix
        }
        unsafe_issue_types = {
            issue.issue_type
            for issue in profile.issues
            if not issue.safe_to_fix
        }

        self.assertIn("duplicate_rows", safe_issue_types)
        self.assertIn("blank_rows", safe_issue_types)
        self.assertIn("blank_column", safe_issue_types)
        self.assertIn("missing_values", unsafe_issue_types)

    def test_analysis_does_not_modify_dataframe(self) -> None:
        original = self.loaded_dataset.data.copy(deep=True)

        analyze_dataset(self.loaded_dataset)

        pd.testing.assert_frame_equal(
            self.loaded_dataset.data,
            original,
            check_dtype=True,
        )

    def test_whitespace_only_row_is_treated_as_blank(self) -> None:
        data = pd.DataFrame(
            {
                "Name": ["Alice", "   ", "Alice"],
                "Empty Column": ["", None, ""],
            }
        )
        loaded = LoadedDataset(
            data=data,
            source_path=Path("temporary.csv"),
            file_type="CSV",
            file_size_bytes=100,
        )

        profile = analyze_dataset(loaded)

        self.assertEqual(profile.blank_rows, 1)
        self.assertEqual(profile.duplicate_rows, 1)
        self.assertEqual(profile.blank_columns, ("Empty Column",))


if __name__ == "__main__":
    unittest.main()