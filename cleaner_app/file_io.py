"""Safe loading and validation of CSV and Excel datasets."""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


SUPPORTED_EXTENSIONS = {".csv", ".xlsx"}
CSV_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252")


class DatasetLoadError(Exception):
    """Base exception for dataset-loading problems."""


class UnsupportedFileError(DatasetLoadError):
    """Raised when the selected file type is unsupported."""


class EmptyDatasetError(DatasetLoadError):
    """Raised when the selected file contains no usable table."""


@dataclass(slots=True)
class LoadedDataset:
    """A loaded dataset together with its source information."""

    data: pd.DataFrame
    source_path: Path
    file_type: str
    file_size_bytes: int
    selected_sheet: str | None = None
    available_sheets: tuple[str, ...] = ()
    csv_encoding: str | None = None
    csv_delimiter: str | None = None

    @property
    def row_count(self) -> int:
        """Number of data rows, including fully blank rows."""

        return len(self.data.index)

    @property
    def column_count(self) -> int:
        """Number of columns in the dataset."""

        return len(self.data.columns)


def validate_file_path(file_path: str | Path) -> Path:
    """Validate the selected path and return an absolute Path object."""

    path = Path(file_path).expanduser()

    if not path.exists():
        raise DatasetLoadError(f"File does not exist: {path}")

    if not path.is_file():
        raise DatasetLoadError(f"The selected path is not a file: {path}")

    extension = path.suffix.lower()

    if extension not in SUPPORTED_EXTENSIONS:
        supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        raise UnsupportedFileError(
            f"Unsupported file type '{extension or 'none'}'. "
            f"Supported types: {supported}"
        )

    if path.stat().st_size == 0:
        raise EmptyDatasetError("The selected file is empty.")

    return path.resolve()


def _read_csv_sample(path: Path) -> tuple[str, str]:
    """Detect a usable text encoding and likely CSV delimiter."""

    last_error: UnicodeDecodeError | None = None

    for encoding in CSV_ENCODINGS:
        try:
            with path.open("r", encoding=encoding, newline="") as handle:
                sample = handle.read(16_384)
        except UnicodeDecodeError as exc:
            last_error = exc
            continue

        if not sample.strip():
            raise EmptyDatasetError("The selected CSV file contains no data.")

        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            delimiter = dialect.delimiter
        except csv.Error:
            delimiter = ","

        return encoding, delimiter

    raise DatasetLoadError(
        "The CSV text encoding could not be read safely."
    ) from last_error


def _load_csv(path: Path) -> LoadedDataset:
    """Load a CSV while preserving blank rows and original text values."""

    encoding, delimiter = _read_csv_sample(path)

    try:
        data = pd.read_csv(
            path,
            sep=delimiter,
            encoding=encoding,
            dtype=object,
            keep_default_na=False,
            skip_blank_lines=False,
        )
    except pd.errors.EmptyDataError as exc:
        raise EmptyDatasetError(
            "The selected CSV file contains no columns or records."
        ) from exc
    except (OSError, ValueError, csv.Error) as exc:
        raise DatasetLoadError(f"Could not read the CSV file: {exc}") from exc

    if data.shape[1] == 0:
        raise EmptyDatasetError("The selected CSV file contains no columns.")

    return LoadedDataset(
        data=data,
        source_path=path,
        file_type="CSV",
        file_size_bytes=path.stat().st_size,
        csv_encoding=encoding,
        csv_delimiter=delimiter,
    )


def list_excel_sheets(file_path: str | Path) -> tuple[str, ...]:
    """Return the worksheet names in an Excel workbook."""

    path = validate_file_path(file_path)

    if path.suffix.lower() != ".xlsx":
        raise UnsupportedFileError(
            "Worksheet discovery is only available for .xlsx files."
        )

    try:
        with pd.ExcelFile(path, engine="openpyxl") as workbook:
            sheet_names = tuple(workbook.sheet_names)
    except (OSError, ValueError, KeyError) as exc:
        raise DatasetLoadError(
            f"Could not read the Excel workbook: {exc}"
        ) from exc

    if not sheet_names:
        raise EmptyDatasetError("The Excel workbook contains no worksheets.")

    return sheet_names


def _load_excel(path: Path, sheet_name: str | None) -> LoadedDataset:
    """Load one worksheet from an Excel workbook."""

    available_sheets = list_excel_sheets(path)
    selected_sheet = sheet_name or available_sheets[0]

    if selected_sheet not in available_sheets:
        available = ", ".join(available_sheets)
        raise DatasetLoadError(
            f"Worksheet '{selected_sheet}' was not found. "
            f"Available worksheets: {available}"
        )

    try:
        data = pd.read_excel(
            path,
            sheet_name=selected_sheet,
            engine="openpyxl",
            dtype=object,
            keep_default_na=False,
        )
    except (OSError, ValueError, KeyError) as exc:
        raise DatasetLoadError(
            f"Could not read worksheet '{selected_sheet}': {exc}"
        ) from exc

    if data.shape[1] == 0:
        raise EmptyDatasetError(
            f"Worksheet '{selected_sheet}' contains no columns."
        )

    return LoadedDataset(
        data=data,
        source_path=path,
        file_type="Excel",
        file_size_bytes=path.stat().st_size,
        selected_sheet=selected_sheet,
        available_sheets=available_sheets,
    )


def load_dataset(
    file_path: str | Path,
    sheet_name: str | None = None,
) -> LoadedDataset:
    """Load a supported dataset without modifying the source file."""

    path = validate_file_path(file_path)

    if path.suffix.lower() == ".csv":
        if sheet_name is not None:
            raise DatasetLoadError(
                "A worksheet name cannot be used when loading a CSV file."
            )
        return _load_csv(path)

    return _load_excel(path, sheet_name)