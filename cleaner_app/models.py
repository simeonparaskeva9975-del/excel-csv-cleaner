"""Shared data models for analysis, cleaning, and reporting."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import StrEnum
from typing import Any


class IssueSeverity(StrEnum):
    """How strongly an issue should be presented to the user."""

    INFO = "info"
    WARNING = "warning"
    ERROR = "error"


class CleaningAction(StrEnum):
    """Supported cleaning actions for the first release."""

    REMOVE_DUPLICATES = "remove_duplicates"
    REMOVE_BLANK_ROWS = "remove_blank_rows"
    REMOVE_BLANK_COLUMNS = "remove_blank_columns"
    TRIM_WHITESPACE = "trim_whitespace"
    STANDARDIZE_COLUMN_NAMES = "standardize_column_names"
    STANDARDIZE_TEXT_CASE = "standardize_text_case"
    CONVERT_NUMERIC = "convert_numeric"
    CONVERT_DATES = "convert_dates"


@dataclass(frozen=True, slots=True)
class DataIssue:
    """One detected quality problem."""

    issue_type: str
    message: str
    severity: IssueSeverity = IssueSeverity.WARNING
    column: str | None = None
    row_number: int | None = None
    count: int = 1
    examples: tuple[str, ...] = ()
    safe_to_fix: bool = False

    def __post_init__(self) -> None:
        if self.count < 1:
            raise ValueError("Issue count must be at least 1.")
        if self.row_number is not None and self.row_number < 1:
            raise ValueError("Row numbers are one-based and must be positive.")


@dataclass(frozen=True, slots=True)
class CleaningOption:
    """A cleaning choice shown before changes are applied."""

    action: CleaningAction
    enabled: bool
    description: str
    target_columns: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class CleaningLogEntry:
    """Auditable record of one applied cleaning operation."""

    action: CleaningAction
    description: str
    affected_cells: int = 0
    affected_rows: int = 0
    column: str | None = None
    before_value: Any = None
    after_value: Any = None
    created_at_utc: datetime = field(
        default_factory=lambda: datetime.now(timezone.utc)
    )

    def __post_init__(self) -> None:
        if self.affected_cells < 0 or self.affected_rows < 0:
            raise ValueError("Affected counts cannot be negative.")

@dataclass(slots=True)
class CleaningResult:
    """Cleaned dataset and its complete audit log."""

    source_name: str
    data: Any
    original_row_count: int
    original_column_count: int
    log_entries: tuple[CleaningLogEntry, ...] = ()

    def __post_init__(self) -> None:
        if self.original_row_count < 0:
            raise ValueError("Original row count cannot be negative.")

        if self.original_column_count < 0:
            raise ValueError("Original column count cannot be negative.")

    @property
    def cleaned_row_count(self) -> int:
        return len(self.data.index)

    @property
    def cleaned_column_count(self) -> int:
        return len(self.data.columns)

    @property
    def removed_row_count(self) -> int:
        return self.original_row_count - self.cleaned_row_count
    
@dataclass(frozen=True, slots=True)
class NumericColumnFinding:
    """Evidence that a column is intended to contain numeric values."""

    column: str
    nonblank_count: int
    numeric_value_count: int
    numeric_text_count: int
    invalid_count: int
    numeric_text_examples: tuple[str, ...] = ()
    invalid_examples: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        counts = (
            self.nonblank_count,
            self.numeric_value_count,
            self.numeric_text_count,
            self.invalid_count,
        )
        if any(count < 0 for count in counts):
            raise ValueError("Numeric finding counts cannot be negative.")
        if self.convertible_count + self.invalid_count != self.nonblank_count:
            raise ValueError(
                "Numeric finding counts must equal the nonblank count."
            )

    @property
    def convertible_count(self) -> int:
        """Values already numeric or recognizable as numeric text."""

        return self.numeric_value_count + self.numeric_text_count

    @property
    def conversion_ratio(self) -> float:
        """Share of nonblank values recognizable as numeric."""

        if self.nonblank_count == 0:
            return 0.0
        return self.convertible_count / self.nonblank_count


@dataclass(frozen=True, slots=True)
class DateColumnFinding:
    """Evidence that a column is intended to contain dates."""

    column: str
    nonblank_count: int
    date_value_count: int
    recognized_text_count: int
    ambiguous_text_count: int
    invalid_count: int
    detected_formats: tuple[str, ...] = ()
    ambiguous_examples: tuple[str, ...] = ()
    invalid_examples: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        counts = (
            self.nonblank_count,
            self.date_value_count,
            self.recognized_text_count,
            self.ambiguous_text_count,
            self.invalid_count,
        )
        if any(count < 0 for count in counts):
            raise ValueError("Date finding counts cannot be negative.")
        if self.recognizable_count + self.invalid_count != self.nonblank_count:
            raise ValueError("Date finding counts must equal the nonblank count.")

    @property
    def recognizable_count(self) -> int:
        """Values recognizable as dates, including ambiguous dates."""

        return (
            self.date_value_count
            + self.recognized_text_count
            + self.ambiguous_text_count
        )

    @property
    def recognition_ratio(self) -> float:
        """Share of nonblank values recognizable as dates."""

        if self.nonblank_count == 0:
            return 0.0
        return self.recognizable_count / self.nonblank_count

@dataclass(frozen=True, slots=True)
class SuspiciousNumericFinding:
    """Potentially invalid or unusual values in a numeric column."""

    column: str
    negative_count: int = 0
    outlier_count: int = 0
    negative_examples: tuple[str, ...] = ()
    outlier_examples: tuple[str, ...] = ()
    lower_outlier_bound: float | None = None
    upper_outlier_bound: float | None = None

    def __post_init__(self) -> None:
        if self.negative_count < 0 or self.outlier_count < 0:
            raise ValueError(
                "Suspicious-value counts cannot be negative."
            )
        
@dataclass(slots=True)
class DatasetProfile:
    """Analysis results consumed by the GUI and report exporter."""

    source_name: str
    row_count: int
    column_count: int
    duplicate_rows: int = 0
    blank_rows: int = 0
    blank_columns: tuple[str, ...] = ()
    missing_by_column: dict[str, int] = field(default_factory=dict)
    dtypes: dict[str, str] = field(default_factory=dict)
    whitespace_cells_by_column: dict[str, int] = field(
        default_factory=dict
    )
    column_name_suggestions: dict[str, str] = field(
        default_factory=dict
    )
    case_variants_by_column: dict[
        str,
        dict[str, tuple[str, ...]],
    ] = field(default_factory=dict)
    numeric_findings: dict[
        str,
        NumericColumnFinding,
    ] = field(default_factory=dict)
    date_findings: dict[
        str,
        DateColumnFinding,
    ] = field(default_factory=dict)
    suspicious_numeric_findings: dict[
        str,
        SuspiciousNumericFinding,
    ] = field(default_factory=dict)
    issues: list[DataIssue] = field(default_factory=list)

    def __post_init__(self) -> None:
        if self.row_count < 0 or self.column_count < 0:
            raise ValueError(
                "Dataset dimensions cannot be negative."
            )