"""Non-destructive data-quality analysis."""

from __future__ import annotations

import re
from collections import Counter
from datetime import date, datetime
from numbers import Number
from typing import Any
from unicodedata import numeric

import pandas as pd

from cleaner_app.file_io import LoadedDataset
from cleaner_app.models import (
    DataIssue,
    DatasetProfile,
    DateColumnFinding,
    IssueSeverity,
    NumericColumnFinding,
)

from cleaner_app.models import (
    DataIssue,
    DatasetProfile,
    DateColumnFinding,
    IssueSeverity,
    NumericColumnFinding,
    SuspiciousNumericFinding,
)


NUMERIC_TEXT_PATTERN = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$")
ISO_DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")
SLASH_DATE_PATTERN = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")
DOT_DATE_PATTERN = re.compile(r"^\d{1,2}\.\d{1,2}\.\d{4}$")
MONTH_NAME_DATE_PATTERN = re.compile(r"^\d{1,2}\s+[A-Za-z]{3}\s+\d{4}$")


def is_blank_value(value: Any) -> bool:
    """Return True for missing values and empty or whitespace-only text."""

    if pd.isna(value):
        return True
    return isinstance(value, str) and value.strip() == ""


def has_outer_whitespace(value: Any) -> bool:
    """Return True when text has leading or trailing whitespace."""

    return isinstance(value, str) and value != value.strip()


def normalize_column_name(column_name: Any) -> str:
    """Create a readable snake_case suggestion for a column name."""

    normalized = str(column_name).strip()
    replacements = {"€": " eur ", "$": " usd ", "%": " percent ", "&": " and "}
    for symbol, replacement in replacements.items():
        normalized = normalized.replace(symbol, replacement)
    normalized = re.sub(r"[^\w]+", "_", normalized, flags=re.UNICODE)
    normalized = re.sub(r"_+", "_", normalized).strip("_").casefold()
    return normalized or "unnamed_column"


def build_blank_mask(data: pd.DataFrame) -> pd.DataFrame:
    """Create a boolean DataFrame identifying blank cells."""

    return data.map(is_blank_value)


def find_column_name_suggestions(
    data: pd.DataFrame,
) -> tuple[dict[str, str], set[str]]:
    """Return rename suggestions and any suggestions that would collide."""

    suggestions = {
        str(column): normalize_column_name(column) for column in data.columns
    }
    counts = Counter(suggestions.values())
    collisions = {name for name, count in counts.items() if count > 1}
    changed = {
        original: suggestion
        for original, suggestion in suggestions.items()
        if original != suggestion
    }
    return changed, collisions


def find_outer_whitespace(
    data: pd.DataFrame,
) -> tuple[dict[str, int], dict[str, tuple[str, ...]]]:
    """Count text cells containing leading or trailing whitespace."""

    counts: dict[str, int] = {}
    examples: dict[str, tuple[str, ...]] = {}
    for column in data.columns:
        mask = data[column].map(has_outer_whitespace)
        count = int(mask.sum())
        if not count:
            continue
        name = str(column)
        counts[name] = count
        values = data.loc[mask, column].drop_duplicates().head(3).tolist()
        examples[name] = tuple(repr(value) for value in values)
    return counts, examples


def find_case_variants(
    data: pd.DataFrame,
) -> dict[str, dict[str, tuple[str, ...]]]:
    """Find text values that differ only by capitalization."""

    findings: dict[str, dict[str, tuple[str, ...]]] = {}
    for column in data.columns:
        grouped: dict[str, set[str]] = {}
        for value in data[column]:
            if not isinstance(value, str) or not value.strip():
                continue
            stripped = value.strip()
            grouped.setdefault(stripped.casefold(), set()).add(stripped)
        variants = {
            key: tuple(sorted(values, key=lambda item: (item.casefold(), item)))
            for key, values in grouped.items()
            if len(values) > 1
        }
        if variants:
            findings[str(column)] = variants
    return findings


def is_numeric_value(value: Any) -> bool:
    """Return True for real numeric values, excluding booleans."""

    return (
        isinstance(value, Number)
        and not isinstance(value, bool)
        and not pd.isna(value)
    )


def is_safe_numeric_text(value: Any) -> bool:
    """Return True for an unambiguous numeric text representation."""

    return (
        isinstance(value, str)
        and bool(value.strip())
        and bool(NUMERIC_TEXT_PATTERN.fullmatch(value.strip()))
    )


def find_numeric_columns(
    data: pd.DataFrame,
    minimum_ratio: float = 0.80,
) -> dict[str, NumericColumnFinding]:
    """Find columns whose nonblank values are predominantly numeric."""

    findings: dict[str, NumericColumnFinding] = {}
    for column in data.columns:
        numeric_count = 0
        numeric_text: list[str] = []
        invalid: list[str] = []
        for value in data[column]:
            if is_blank_value(value):
                continue
            if is_numeric_value(value):
                numeric_count += 1
            elif is_safe_numeric_text(value):
                numeric_text.append(str(value))
            else:
                invalid.append(str(value))
        nonblank = numeric_count + len(numeric_text) + len(invalid)
        convertible = numeric_count + len(numeric_text)
        if not nonblank or convertible / nonblank < minimum_ratio:
            continue
        name = str(column)
        findings[name] = NumericColumnFinding(
            column=name,
            nonblank_count=nonblank,
            numeric_value_count=numeric_count,
            numeric_text_count=len(numeric_text),
            invalid_count=len(invalid),
            numeric_text_examples=tuple(dict.fromkeys(numeric_text))[:3],
            invalid_examples=tuple(dict.fromkeys(invalid))[:3],
        )
    return findings

def numeric_value_for_analysis(
    value: Any,
) -> float | None:
    """Return a numeric value for statistical analysis when safe."""

    if is_numeric_value(value):
        return float(value)

    if is_safe_numeric_text(value):
        return float(value.strip())

    return None


def format_numeric_example(value: float) -> str:
    """Format a numeric example without unnecessary decimal zeros."""

    return f"{value:g}"


def find_suspicious_numeric_values(
    data: pd.DataFrame,
    numeric_findings: dict[str, NumericColumnFinding],
) -> dict[str, SuspiciousNumericFinding]:
    """Flag negative values and IQR outliers without changing them."""

    findings: dict[str, SuspiciousNumericFinding] = {}

    for column in numeric_findings:
        numeric_values = [
            parsed_value
            for value in data[column]
            if (
                parsed_value := numeric_value_for_analysis(value)
            )
            is not None
        ]

        if not numeric_values:
            continue

        numeric_series = pd.Series(
            numeric_values,
            dtype="float64",
        )

        negative_values = numeric_series[
            numeric_series < 0
        ].tolist()

        lower_bound: float | None = None
        upper_bound: float | None = None
        outlier_values: list[float] = []

        # Avoid applying the IQR rule to tiny or nearly constant samples.
        if (
            len(numeric_series) >= 8
            and numeric_series.nunique() >= 4
        ):
            first_quartile = float(
                numeric_series.quantile(0.25)
            )
            third_quartile = float(
                numeric_series.quantile(0.75)
            )
            interquartile_range = (
                third_quartile - first_quartile
            )

            if interquartile_range > 0:
                lower_bound = (
                    first_quartile
                    - 1.5 * interquartile_range
                )
                upper_bound = (
                    third_quartile
                    + 1.5 * interquartile_range
                )

                outlier_values = numeric_series[
                    (numeric_series < lower_bound)
                    | (numeric_series > upper_bound)
                ].tolist()

        if not negative_values and not outlier_values:
            continue

        findings[column] = SuspiciousNumericFinding(
            column=column,
            negative_count=len(negative_values),
            outlier_count=len(outlier_values),
            negative_examples=tuple(
                dict.fromkeys(
                    format_numeric_example(value)
                    for value in negative_values
                )
            )[:3],
            outlier_examples=tuple(
                dict.fromkeys(
                    format_numeric_example(value)
                    for value in outlier_values
                )
            )[:3],
            lower_outlier_bound=lower_bound,
            upper_outlier_bound=upper_bound,
        )

    return findings
def _matches_date_format(value: str, date_format: str) -> bool:
    """Return True when text is a valid calendar date."""

    try:
        datetime.strptime(value, date_format)
    except ValueError:
        return False
    return True


def classify_date_text(value: Any) -> tuple[str, str | None]:
    """Classify text as recognized, ambiguous, or invalid date text."""

    if not isinstance(value, str) or not value.strip():
        return "invalid", None
    text = value.strip()
    if ISO_DATE_PATTERN.fullmatch(text):
        if _matches_date_format(text, "%Y-%m-%d"):
            return "recognized", "yyyy-mm-dd"
        return "invalid", None
    if MONTH_NAME_DATE_PATTERN.fullmatch(text):
        if _matches_date_format(text, "%d %b %Y"):
            return "recognized", "dd Mon yyyy"
        return "invalid", None
    if DOT_DATE_PATTERN.fullmatch(text):
        if _matches_date_format(text, "%d.%m.%Y"):
            return "recognized", "dd.mm.yyyy"
        return "invalid", None
    match = SLASH_DATE_PATTERN.fullmatch(text)
    if not match:
        return "invalid", None
    first, second = int(match.group(1)), int(match.group(2))
    if first <= 12 and second <= 12:
        return "ambiguous", "dd/mm/yyyy or mm/dd/yyyy"
    if _matches_date_format(text, "%d/%m/%Y"):
        return "recognized", "dd/mm/yyyy"
    if _matches_date_format(text, "%m/%d/%Y"):
        return "recognized", "mm/dd/yyyy"
    return "invalid", None


def find_date_columns(
    data: pd.DataFrame,
    minimum_ratio: float = 0.80,
) -> dict[str, DateColumnFinding]:
    """Find columns whose nonblank values are predominantly dates."""

    findings: dict[str, DateColumnFinding] = {}
    for column in data.columns:
        date_count = 0
        recognized_count = 0
        ambiguous: list[str] = []
        invalid: list[str] = []
        formats: list[str] = []
        for value in data[column]:
            if is_blank_value(value):
                continue
            if isinstance(value, (pd.Timestamp, datetime, date)):
                date_count += 1
                formats.append("Excel date value")
                continue
            classification, detected_format = classify_date_text(value)
            if classification == "recognized":
                recognized_count += 1
                if detected_format:
                    formats.append(detected_format)
            elif classification == "ambiguous":
                ambiguous.append(str(value))
                if detected_format:
                    formats.append(detected_format)
            else:
                invalid.append(str(value))
        nonblank = date_count + recognized_count + len(ambiguous) + len(invalid)
        recognizable = date_count + recognized_count + len(ambiguous)
        if not nonblank or recognizable / nonblank < minimum_ratio:
            continue
        name = str(column)
        findings[name] = DateColumnFinding(
            column=name,
            nonblank_count=nonblank,
            date_value_count=date_count,
            recognized_text_count=recognized_count,
            ambiguous_text_count=len(ambiguous),
            invalid_count=len(invalid),
            detected_formats=tuple(dict.fromkeys(formats)),
            ambiguous_examples=tuple(dict.fromkeys(ambiguous))[:3],
            invalid_examples=tuple(dict.fromkeys(invalid))[:3],
        )
    return findings


def analyze_dataset(dataset: LoadedDataset) -> DatasetProfile:
    """Analyze a loaded dataset without modifying it."""

    data = dataset.data
    blank_mask = build_blank_mask(data)
    blank_rows_mask = blank_mask.all(axis=1)
    blank_row_count = int(blank_rows_mask.sum())
    blank_columns = tuple(
        str(column) for column in data.columns if bool(blank_mask[column].all())
    )
    business_data = data.loc[~blank_rows_mask]
    business_blank_mask = blank_mask.loc[~blank_rows_mask]
    duplicate_count = int(business_data.duplicated(keep="first").sum())
    missing = {
        str(column): int(business_blank_mask[column].sum())
        for column in data.columns
    }
    dtypes = {str(column): str(dtype) for column, dtype in data.dtypes.items()}
    whitespace, whitespace_examples = find_outer_whitespace(business_data)
    suggestions, collisions = find_column_name_suggestions(data)
    case_variants = find_case_variants(business_data)
    numeric = find_numeric_columns(business_data)
    suspicious_numeric = find_suspicious_numeric_values(
    business_data,
    numeric,
    )
    dates = find_date_columns(business_data)
    issues: list[DataIssue] = []

    if duplicate_count:
        issues.append(
            DataIssue(
                "duplicate_rows",
                f"{duplicate_count} exact duplicate rows were found.",
                count=duplicate_count,
                safe_to_fix=True,
            )
        )
    if blank_row_count:
        issues.append(
            DataIssue(
                "blank_rows",
                f"{blank_row_count} fully blank rows were found.",
                IssueSeverity.INFO,
                count=blank_row_count,
                safe_to_fix=True,
            )
        )
    for column in blank_columns:
        issues.append(
            DataIssue(
                "blank_column",
                f"Column '{column}' is fully blank.",
                IssueSeverity.INFO,
                column=column,
                safe_to_fix=True,
            )
        )
    for column, count in missing.items():
        if count and column not in blank_columns:
            issues.append(
                DataIssue(
                    "missing_values",
                    f"Column '{column}' contains {count} missing values.",
                    column=column,
                    count=count,
                )
            )
    for column, count in whitespace.items():
        issues.append(
            DataIssue(
                "outer_whitespace",
                f"Column '{column}' contains {count} cells with outer whitespace.",
                IssueSeverity.INFO,
                column=column,
                count=count,
                examples=whitespace_examples[column],
                safe_to_fix=True,
            )
        )
    for original, suggestion in suggestions.items():
        collision = suggestion in collisions
        issues.append(
            DataIssue(
                "column_name_collision" if collision else "column_name_format",
                f"Column '{original}' can be normalized to '{suggestion}'.",
                IssueSeverity.WARNING if collision else IssueSeverity.INFO,
                column=original,
                safe_to_fix=not collision,
            )
        )
    for column, groups in case_variants.items():
        for variants in groups.values():
            issues.append(
                DataIssue(
                    "case_variants",
                    f"Column '{column}' contains capitalization variants.",
                    column=column,
                    count=len(variants),
                    examples=variants,
                )
            )
    for column, finding in suspicious_numeric.items():
            if finding.negative_count:
                issues.append(
                    DataIssue(
                        issue_type="negative_numeric_values",
                        message=(
                            f"Column '{column}' contains "
                            f"{finding.negative_count} negative "
                            f"{'value' if finding.negative_count == 1 else 'values'}."
                        ),
                        severity=IssueSeverity.WARNING,
                        column=column,
                        count=finding.negative_count,
                        examples=finding.negative_examples,
                        safe_to_fix=False,
                )
             )

            if finding.outlier_count:
                issues.append(
                    DataIssue(
                        issue_type="numeric_outliers",
                        message=(
                            f"Column '{column}' contains "
                            f"{finding.outlier_count} potential statistical "
                            f"{'outlier' if finding.outlier_count == 1 else 'outliers'}."
                        ),
                        severity=IssueSeverity.WARNING,
                        column=column,
                        count=finding.outlier_count,
                        examples=finding.outlier_examples,
                        safe_to_fix=False,
                    )
            )
    for column, finding in numeric.items():
        if finding.invalid_count:
            issues.append(
                DataIssue(
                    "mixed_numeric_values",
                    f"Column '{column}' contains uncertain numeric values.",
                    column=column,
                    count=finding.invalid_count,
                    examples=finding.invalid_examples,
                )
            )

        if finding.numeric_text_count:
            issues.append(
                DataIssue(
                    "numeric_text",
                    f"Column '{column}' contains numeric-looking text.",
                    IssueSeverity.INFO,
                    column=column,
                    count=finding.numeric_text_count,
                    examples=finding.numeric_text_examples,
                    safe_to_fix=not finding.invalid_count,
                )
            )

    for column, finding in suspicious_numeric.items():
        if finding.negative_count:
            issues.append(
                DataIssue(
                    issue_type="negative_numeric_values",
                    message=(
                        f"Column '{column}' contains "
                        f"{finding.negative_count} negative "
                        f"{'value' if finding.negative_count == 1 else 'values'}."
                    ),
                    severity=IssueSeverity.WARNING,
                    column=column,
                    count=finding.negative_count,
                    examples=finding.negative_examples,
                    safe_to_fix=False,
                )
            )

        if finding.outlier_count:
            issues.append(
                DataIssue(
                    issue_type="numeric_outliers",
                    message=(
                        f"Column '{column}' contains "
                        f"{finding.outlier_count} potential statistical "
                        f"{'outlier' if finding.outlier_count == 1 else 'outliers'}."
                    ),
                    severity=IssueSeverity.WARNING,
                    column=column,
                    count=finding.outlier_count,
                    examples=finding.outlier_examples,
                    safe_to_fix=False,
                )
            )

    for column, finding in dates.items():
        if len(finding.detected_formats) > 1:
            issues.append(
                DataIssue(
                    "mixed_date_formats",
                    f"Column '{column}' contains multiple date formats.",
                    column=column,
                    count=len(finding.detected_formats),
                    examples=finding.detected_formats,
                )
            )
        if finding.ambiguous_text_count:
            issues.append(
                DataIssue(
                    "ambiguous_dates",
                    f"Column '{column}' contains ambiguous slash dates.",
                    column=column,
                    count=finding.ambiguous_text_count,
                    examples=finding.ambiguous_examples,
                )
            )
        if finding.invalid_count:
            issues.append(
                DataIssue(
                    "invalid_date_values",
                    f"Column '{column}' contains invalid date values.",
                    column=column,
                    count=finding.invalid_count,
                    examples=finding.invalid_examples,
                )
            )

    return DatasetProfile(
        source_name=dataset.source_path.name,
        row_count=dataset.row_count,
        column_count=dataset.column_count,
        duplicate_rows=duplicate_count,
        blank_rows=blank_row_count,
        blank_columns=blank_columns,
        missing_by_column=missing,
        dtypes=dtypes,
        whitespace_cells_by_column=whitespace,
        column_name_suggestions=suggestions,
        case_variants_by_column=case_variants,
        numeric_findings=numeric,
        date_findings=dates,
        suspicious_numeric_findings=suspicious_numeric,
        issues=issues,
    )
