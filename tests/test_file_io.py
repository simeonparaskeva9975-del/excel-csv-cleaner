"""Tests for safe CSV and Excel loading."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from cleaner_app.file_io import (
    DatasetLoadError,
    EmptyDatasetError,
    UnsupportedFileError,
    list_excel_sheets,
    load_dataset,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "sample_data"
CSV_PATH = DATA_DIR / "messy_sales_data.csv"
XLSX_PATH = DATA_DIR / "messy_sales_data.xlsx"


class FileIOTests(unittest.TestCase):
    def test_load_demo_csv(self) -> None:
        result = load_dataset(CSV_PATH)

        self.assertEqual(result.file_type, "CSV")
        self.assertEqual(result.row_count, 131)
        self.assertEqual(result.column_count, 13)
        self.assertEqual(result.csv_encoding, "utf-8-sig")
        self.assertEqual(result.csv_delimiter, ",")
        self.assertIsNone(result.selected_sheet)

    def test_load_demo_excel(self) -> None:
        result = load_dataset(XLSX_PATH)

        self.assertEqual(result.file_type, "Excel")
        self.assertEqual(result.row_count, 131)
        self.assertEqual(result.column_count, 13)
        self.assertEqual(result.selected_sheet, "Messy Sales Data")
        self.assertEqual(result.available_sheets, ("Messy Sales Data",))

    def test_excel_sheet_discovery(self) -> None:
        sheets = list_excel_sheets(XLSX_PATH)

        self.assertEqual(sheets, ("Messy Sales Data",))

    def test_unknown_excel_sheet_is_rejected(self) -> None:
        with self.assertRaises(DatasetLoadError):
            load_dataset(XLSX_PATH, sheet_name="Does Not Exist")

    def test_csv_rejects_worksheet_name(self) -> None:
        with self.assertRaises(DatasetLoadError):
            load_dataset(CSV_PATH, sheet_name="Sheet1")

    def test_missing_file_is_rejected(self) -> None:
        missing_path = DATA_DIR / "missing_file.csv"

        with self.assertRaises(DatasetLoadError):
            load_dataset(missing_path)

    def test_unsupported_file_type_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            text_path = Path(temporary_directory) / "notes.txt"
            text_path.write_text("not a dataset", encoding="utf-8")

            with self.assertRaises(UnsupportedFileError):
                load_dataset(text_path)

    def test_empty_file_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            empty_path = Path(temporary_directory) / "empty.csv"
            empty_path.write_bytes(b"")

            with self.assertRaises(EmptyDatasetError):
                load_dataset(empty_path)

    def test_semicolon_delimiter_is_detected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            csv_path = Path(temporary_directory) / "semicolon_data.csv"
            csv_path.write_text(
                "name;amount\nAlice;10\nBob;20\n",
                encoding="utf-8",
            )

            result = load_dataset(csv_path)

            self.assertEqual(result.csv_delimiter, ";")
            self.assertEqual(result.row_count, 2)
            self.assertEqual(result.column_count, 2)
            self.assertEqual(result.data.iloc[0]["amount"], "10")

    def test_blank_csv_rows_are_preserved(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            csv_path = Path(temporary_directory) / "blank_rows.csv"
            csv_path.write_text(
                "name,amount\nAlice,10\n,\nBob,20\n",
                encoding="utf-8",
            )

            result = load_dataset(csv_path)
            blank_rows = result.data.eq("").all(axis=1)

            self.assertEqual(result.row_count, 3)
            self.assertEqual(int(blank_rows.sum()), 1)

    def test_loading_does_not_modify_source_files(self) -> None:
        csv_before = CSV_PATH.read_bytes()
        excel_before = XLSX_PATH.read_bytes()

        load_dataset(CSV_PATH)
        load_dataset(XLSX_PATH)

        self.assertEqual(CSV_PATH.read_bytes(), csv_before)
        self.assertEqual(XLSX_PATH.read_bytes(), excel_before)


if __name__ == "__main__":
    unittest.main()