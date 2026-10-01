"""Tests for conservative cleaning operations."""

from __future__ import annotations

import unittest
from pathlib import Path

import pandas as pd
import tempfile

from openpyxl import load_workbook

from cleaner_app.cleaning import FOUNDATION_ACTIONS, clean_dataset
from cleaner_app.file_io import LoadedDataset, load_dataset
from cleaner_app.analysis import analyze_dataset

from cleaner_app.models import CleaningAction


PROJECT_ROOT = Path(__file__).resolve().parents[1]
XLSX_PATH = PROJECT_ROOT / "sample_data" / "messy_sales_data.xlsx"


class CleaningTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.loaded_dataset = load_dataset(XLSX_PATH)

    def test_complete_cleaning_pipeline_can_be_reanalyzed(
        self,
    ) -> None:
        all_actions = FOUNDATION_ACTIONS | {
            CleaningAction.STANDARDIZE_TEXT_CASE,
            CleaningAction.CONVERT_NUMERIC,
            CleaningAction.CONVERT_DATES,
        }

        result = clean_dataset(
            self.loaded_dataset,
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

        self.assertEqual(result.cleaned_row_count, 120)
        self.assertEqual(result.cleaned_column_count, 12)
        self.assertEqual(len(result.log_entries), 33)

        cleaned_dataset = LoadedDataset(
            data=result.data,
            source_path=Path("cleaned_sales_data.xlsx"),
            file_type="Excel",
            file_size_bytes=0,
        )

        cleaned_profile = analyze_dataset(cleaned_dataset)

        self.assertEqual(cleaned_profile.row_count, 120)
        self.assertEqual(cleaned_profile.column_count, 12)
        self.assertEqual(cleaned_profile.blank_rows, 0)
        self.assertEqual(cleaned_profile.duplicate_rows, 0)
        self.assertEqual(cleaned_profile.blank_columns, ())
        self.assertEqual(
            cleaned_profile.whitespace_cells_by_column,
            {},
        )
        self.assertEqual(
            cleaned_profile.column_name_suggestions,
            {},
        )
        self.assertEqual(
            cleaned_profile.case_variants_by_column,
            {},
        )

        self.assertEqual(
            cleaned_profile.numeric_findings[
                "quantity"
            ].numeric_text_count,
            0,
        )
        self.assertEqual(
            cleaned_profile.numeric_findings[
                "unit_price_eur"
            ].numeric_text_count,
            0,
        )

        date_finding = cleaned_profile.date_findings["order_date"]

        self.assertEqual(date_finding.date_value_count, 104)
        self.assertEqual(date_finding.recognized_text_count, 0)
        self.assertEqual(date_finding.ambiguous_text_count, 14)
        self.assertEqual(date_finding.invalid_count, 2)

        self.assertEqual(
            set(cleaned_profile.suspicious_numeric_findings),
            {"quantity"},
        )

    def test_selected_numeric_text_is_converted(self) -> None:
        result = clean_dataset(
            self.loaded_dataset,
            actions=(CleaningAction.CONVERT_NUMERIC,),
            numeric_columns=(
                "Quantity",
                "Unit Price EUR",
                "Discount Percent",
            ),
        )

        numeric_entries = [
            entry
            for entry in result.log_entries
            if entry.action == CleaningAction.CONVERT_NUMERIC
        ]

        self.assertEqual(len(numeric_entries), 3)
        self.assertEqual(
            sum(entry.affected_cells for entry in numeric_entries),
            3,
        )

        converted_values = {
            entry.after_value
            for entry in numeric_entries
        }

        self.assertEqual(converted_values, {3, 4, 34.9})

        quantity_text = {
            value.strip()
            for value in result.data["Quantity"]
            if isinstance(value, str)
        }
        price_text = {
            value.strip()
            for value in result.data["Unit Price EUR"]
            if isinstance(value, str)
        }
        discount_text = {
            value.strip()
            for value in result.data["Discount Percent"]
            if isinstance(value, str)
        }

        self.assertIn("unknown", quantity_text)
        self.assertIn("29,90", price_text)
        self.assertIn("not available", price_text)
        self.assertIn("10%", discount_text)

    def test_numeric_conversion_requires_selected_columns(self) -> None:
        with self.assertRaises(ValueError):
            clean_dataset(
                self.loaded_dataset,
                actions=(CleaningAction.CONVERT_NUMERIC,),
            )

    def test_numeric_conversion_only_changes_selected_columns(
        self,
    ) -> None:
        result = clean_dataset(
            self.loaded_dataset,
            actions=(CleaningAction.CONVERT_NUMERIC,),
            numeric_columns=("Quantity",),
        )

        numeric_entries = [
            entry
            for entry in result.log_entries
            if entry.action == CleaningAction.CONVERT_NUMERIC
        ]

        self.assertEqual(
            sum(entry.affected_cells for entry in numeric_entries),
            2,
        )

        unit_price_text = [
            value
            for value in result.data["Unit Price EUR"]
            if isinstance(value, str)
        ]

        self.assertIn(" 34.90 ", unit_price_text)

    def test_selected_text_case_columns_are_standardized(self) -> None:
        selected_columns = (
            "Product Category",
            "City ",
            "Sales Channel",
            "Order Status",
        )

        result = clean_dataset(
            self.loaded_dataset,
            actions=(CleaningAction.STANDARDIZE_TEXT_CASE,),
            text_case_columns=selected_columns,
        )

        expected_values = {
            "Product Category": ("electronics", "Electronics"),
            "City ": ("athens", "Athens"),
            "Sales Channel": ("online", "Online"),
            "Order Status": ("completed", "Completed"),
        }

        for column, (comparison_key, expected_value) in (
            expected_values.items()
        ):
            matching_values = {
                value.strip()
                for value in result.data[column]
                if (
                    isinstance(value, str)
                    and value.strip().casefold() == comparison_key
                )
            }

            self.assertEqual(matching_values, {expected_value})

        case_entries = [
            entry
            for entry in result.log_entries
            if entry.action == CleaningAction.STANDARDIZE_TEXT_CASE
        ]

        self.assertEqual(len(case_entries), 6)
        self.assertEqual(
            sum(entry.affected_cells for entry in case_entries),
            6,
        )

    def test_text_case_action_requires_selected_columns(self) -> None:
        with self.assertRaises(ValueError):
            clean_dataset(
                self.loaded_dataset,
                actions=(CleaningAction.STANDARDIZE_TEXT_CASE,),
            )

    def test_tied_case_variants_are_not_changed(self) -> None:
        data = pd.DataFrame(
            {
                "Status": ["Open", "OPEN", "Closed"],
            }
        )

        loaded = LoadedDataset(
            data=data,
            source_path=Path("temporary.csv"),
            file_type="CSV",
            file_size_bytes=100,
        )

        result = clean_dataset(
            loaded,
            actions=(CleaningAction.STANDARDIZE_TEXT_CASE,),
            text_case_columns=("Status",),
        )

        self.assertEqual(
            result.data["Status"].tolist(),
            ["Open", "OPEN", "Closed"],
        )
        self.assertEqual(len(result.log_entries), 0)

    def test_column_names_are_standardized_and_logged(self) -> None:
        result = clean_dataset(
            self.loaded_dataset,
            actions=(CleaningAction.STANDARDIZE_COLUMN_NAMES,),
        )

        self.assertEqual(result.cleaned_row_count, 131)
        self.assertEqual(result.cleaned_column_count, 13)

        self.assertIn("order_id", result.data.columns)
        self.assertIn("email_address", result.data.columns)
        self.assertIn("city", result.data.columns)
        self.assertIn("internal_notes", result.data.columns)

        rename_entries = [
            entry
            for entry in result.log_entries
            if entry.action == CleaningAction.STANDARDIZE_COLUMN_NAMES
        ]

        self.assertEqual(len(rename_entries), 13)
    
    def test_outer_whitespace_is_trimmed_and_logged(self) -> None:
        result = clean_dataset(
            self.loaded_dataset,
            actions=(CleaningAction.TRIM_WHITESPACE,),
        )

        whitespace_entries = [
            entry
            for entry in result.log_entries
            if entry.action == CleaningAction.TRIM_WHITESPACE
        ]

        self.assertEqual(len(whitespace_entries), 5)
        self.assertEqual(
            sum(entry.affected_cells for entry in whitespace_entries),
            25,
        )

        for column in result.data.columns:
            for value in result.data[column]:
                if isinstance(value, str):
                    self.assertEqual(value, value.strip())

    def test_fully_blank_column_is_removed(self) -> None:
        result = clean_dataset(
            self.loaded_dataset,
            actions=(CleaningAction.REMOVE_BLANK_COLUMNS,),
        )

        self.assertNotIn("Internal Notes", result.data.columns)
        self.assertEqual(result.cleaned_row_count, 131)
        self.assertEqual(result.cleaned_column_count, 12)

        self.assertEqual(len(result.log_entries), 1)
        self.assertEqual(
            result.log_entries[0].action,
        CleaningAction.REMOVE_BLANK_COLUMNS,
    )
    def test_default_cleaning_removes_blank_and_duplicate_rows(self) -> None:
        result = clean_dataset(self.loaded_dataset)

        self.assertEqual(result.original_row_count, 131)
        self.assertEqual(result.cleaned_row_count, 120)
        self.assertEqual(result.removed_row_count, 11)
        self.assertEqual(result.cleaned_column_count, 12)


        self.assertEqual(len(result.log_entries), 20)
        rename_entries = [
            entry
            for entry in result.log_entries
            if entry.action == CleaningAction.STANDARDIZE_COLUMN_NAMES
        ]

        self.assertEqual(len(rename_entries), 12)

        self.assertEqual(
            tuple(result.data.columns),
            (
                "order_id",
                "customer_name",
                "email_address",
                "order_date",
                "product_category",
                "product_name",
                "quantity",
                "unit_price_eur",
                "discount_percent",
                "city",
                "sales_channel",
                "order_status",
            ),
        )
        self.assertEqual(result.log_entries[0].affected_rows, 3)
        self.assertEqual(result.log_entries[1].affected_rows, 8)
        whitespace_entries = [
            entry
            for entry in result.log_entries
            if entry.action == CleaningAction.TRIM_WHITESPACE
        ]

        self.assertEqual(len(whitespace_entries), 5)
        self.assertEqual(
            sum(entry.affected_cells for entry in whitespace_entries),
            25,
        )
        self.assertEqual(
            result.log_entries[2].action,
            CleaningAction.REMOVE_BLANK_COLUMNS,
        )
        self.assertEqual(result.log_entries[2].column, "Internal Notes")
        self.assertEqual(result.log_entries[2].affected_cells, 120)

    def test_cleaning_does_not_modify_source_dataframe(self) -> None:
        original = self.loaded_dataset.data.copy(deep=True)

        clean_dataset(self.loaded_dataset)

        pd.testing.assert_frame_equal(
            self.loaded_dataset.data,
            original,
            check_dtype=True,
        )

    def test_actions_can_be_selected_individually(self) -> None:
        result = clean_dataset(
            self.loaded_dataset,
            actions=(CleaningAction.REMOVE_BLANK_ROWS,),
        )

        self.assertEqual(result.cleaned_row_count, 128)
        self.assertEqual(len(result.log_entries), 1)
        self.assertEqual(
            result.log_entries[0].action,
            CleaningAction.REMOVE_BLANK_ROWS,
        )

    def test_selected_date_text_is_converted_conservatively(
        self,
    ) -> None:
        result = clean_dataset(
            self.loaded_dataset,
            actions=(CleaningAction.CONVERT_DATES,),
            date_columns=("Order Date",),
        )

        date_entries = [
            entry
            for entry in result.log_entries
            if entry.action == CleaningAction.CONVERT_DATES
        ]

        self.assertEqual(len(date_entries), 4)
        self.assertEqual(
            sum(entry.affected_cells for entry in date_entries),
            111,
        )

        converted_count = sum(
            isinstance(value, pd.Timestamp)
            for value in result.data["Order Date"]
        )
        self.assertEqual(converted_count, 111)

        remaining_text = {
            value
            for value in result.data["Order Date"]
            if isinstance(value, str)
        }

        self.assertIn("07/01/2026", remaining_text)
        self.assertIn("12/02/2026", remaining_text)
        self.assertIn("31/02/2026", remaining_text)
        self.assertIn("not recorded", remaining_text)

    def test_date_conversion_requires_selected_columns(self) -> None:
        with self.assertRaises(ValueError):
            clean_dataset(
                self.loaded_dataset,
                actions=(CleaningAction.CONVERT_DATES,),
            )

    def test_missing_date_conversion_column_is_rejected(self) -> None:
        with self.assertRaises(ValueError):
            clean_dataset(
                self.loaded_dataset,
                actions=(CleaningAction.CONVERT_DATES,),
                date_columns=("Missing Date Column",),
            )
            
    def test_conflicting_column_rename_is_skipped(self) -> None:
        data = pd.DataFrame(
            {
                "Customer Name": ["Alice"],
                "customer_name": ["Bob"],
            }
        )

        loaded = LoadedDataset(
            data=data,
            source_path=Path("temporary.csv"),
            file_type="CSV",
            file_size_bytes=100,
        )

        result = clean_dataset(
            loaded,
            actions=(CleaningAction.STANDARDIZE_COLUMN_NAMES,),
        )

        self.assertEqual(
            tuple(result.data.columns),
            ("Customer Name", "customer_name"),
        )
        self.assertEqual(len(result.log_entries), 0)


if __name__ == "__main__":
    unittest.main()