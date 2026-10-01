"""Conservative and auditable dataset cleaning."""

from __future__ import annotations

from collections.abc import Iterable

from collections import Counter

import pandas as pd

from datetime import datetime

from cleaner_app.file_io import LoadedDataset
from cleaner_app.models import (
    CleaningAction,
    CleaningLogEntry,
    CleaningResult,
)
from cleaner_app.analysis import (
    build_blank_mask,
    classify_date_text,
    find_column_name_suggestions,
    has_outer_whitespace,
    is_safe_numeric_text,
)


FOUNDATION_ACTIONS = frozenset(
    {
        CleaningAction.REMOVE_BLANK_ROWS,
        CleaningAction.REMOVE_DUPLICATES,
        CleaningAction.REMOVE_BLANK_COLUMNS,
        CleaningAction.TRIM_WHITESPACE,
        CleaningAction.STANDARDIZE_COLUMN_NAMES,


    }
)

SUPPORTED_ACTIONS = FOUNDATION_ACTIONS | {
    CleaningAction.STANDARDIZE_TEXT_CASE,
    CleaningAction.CONVERT_NUMERIC,
    CleaningAction.CONVERT_DATES,
}

def standardize_selected_text_case(
    data: pd.DataFrame,
    target_columns: Iterable[str],
) -> list[CleaningLogEntry]:
    """Use the unique most-common capitalization in selected columns."""

    log_entries: list[CleaningLogEntry] = []

    for column in dict.fromkeys(target_columns):
        if column not in data.columns:
            raise ValueError(
                f"Text-case column was not found: {column}"
            )

        grouped_counts: dict[str, Counter[str]] = {}

        for value in data[column]:
            if not isinstance(value, str):
                continue

            stripped_value = value.strip()

            if not stripped_value:
                continue

            comparison_key = stripped_value.casefold()

            grouped_counts.setdefault(
                comparison_key,
                Counter(),
            )[stripped_value] += 1

        for variants in grouped_counts.values():
            if len(variants) < 2:
                continue

            highest_count = max(variants.values())
            most_common_variants = [
                variant
                for variant, count in variants.items()
                if count == highest_count
            ]

            # A tie is uncertain, so no value is changed.
            if len(most_common_variants) != 1:
                continue

            canonical_value = most_common_variants[0]

            for old_value in variants:
                if old_value == canonical_value:
                    continue

                matching_cells = data[column].map(
                    lambda value, old=old_value: (
                        isinstance(value, str)
                        and value.strip() == old
                    )
                )
                changed_count = int(matching_cells.sum())

                if changed_count == 0:
                    continue

                def replace_case(value: str) -> str:
                    leading_length = len(value) - len(value.lstrip())
                    trailing_start = len(value.rstrip())

                    return (
                        value[:leading_length]
                        + canonical_value
                        + value[trailing_start:]
                    )

                data.loc[matching_cells, column] = (
                    data.loc[matching_cells, column].map(
                        replace_case
                    )
                )

                log_entries.append(
                    CleaningLogEntry(
                        action=CleaningAction.STANDARDIZE_TEXT_CASE,
                        description=(
                            f"Standardized '{old_value}' to "
                            f"'{canonical_value}' in column "
                            f"'{column}'."
                        ),
                        affected_cells=changed_count,
                        column=column,
                        before_value=old_value,
                        after_value=canonical_value,
                    )
                )

    return log_entries

def parse_safe_numeric_text(value: str) -> int | float:
    """Convert unambiguous numeric text to an int or float."""

    stripped_value = value.strip()

    if "." in stripped_value:
        return float(stripped_value)

    return int(stripped_value)


def convert_selected_numeric_text(
    data: pd.DataFrame,
    target_columns: Iterable[str],
) -> list[CleaningLogEntry]:
    """Convert only unambiguous numeric text in selected columns."""

    log_entries: list[CleaningLogEntry] = []

    for column in dict.fromkeys(target_columns):
        if column not in data.columns:
            raise ValueError(
                f"Numeric conversion column was not found: {column}"
            )

        conversions: Counter[tuple[str, int | float]] = Counter()

        for row_index, value in data[column].items():
            if not is_safe_numeric_text(value):
                continue

            converted_value = parse_safe_numeric_text(value)
            data.at[row_index, column] = converted_value

            conversions[(value, converted_value)] += 1

        for (
            before_value,
            after_value,
        ), changed_count in conversions.items():
            log_entries.append(
                CleaningLogEntry(
                    action=CleaningAction.CONVERT_NUMERIC,
                    description=(
                        f"Converted numeric text "
                        f"{before_value!r} to {after_value!r} "
                        f"in column '{column}'."
                    ),
                    affected_cells=changed_count,
                    column=column,
                    before_value=before_value,
                    after_value=after_value,
                )
            )

    return log_entries

DATE_FORMAT_CODES = {
    "yyyy-mm-dd": "%Y-%m-%d",
    "dd/mm/yyyy": "%d/%m/%Y",
    "mm/dd/yyyy": "%m/%d/%Y",
    "dd.mm.yyyy": "%d.%m.%Y",
    "dd Mon yyyy": "%d %b %Y",
}


def convert_selected_date_text(
    data: pd.DataFrame,
    target_columns: Iterable[str],
) -> list[CleaningLogEntry]:
    """Convert only recognized, unambiguous dates."""

    log_entries: list[CleaningLogEntry] = []

    for column in dict.fromkeys(target_columns):
        if column not in data.columns:
            raise ValueError(
                f"Date conversion column was not found: {column}"
            )

        conversions_by_format: Counter[str] = Counter()

        for row_index, value in data[column].items():
            classification, detected_format = classify_date_text(
                value
            )

            if classification != "recognized":
                continue

            if detected_format not in DATE_FORMAT_CODES:
                continue

            parsed_value = datetime.strptime(
                value.strip(),
                DATE_FORMAT_CODES[detected_format],
            )

            data.at[row_index, column] = pd.Timestamp(parsed_value)
            conversions_by_format[detected_format] += 1

        for detected_format, changed_count in (
            conversions_by_format.items()
        ):
            log_entries.append(
                CleaningLogEntry(
                    action=CleaningAction.CONVERT_DATES,
                    description=(
                        f"Converted {changed_count} dates from "
                        f"format '{detected_format}' in column "
                        f"'{column}'."
                    ),
                    affected_cells=changed_count,
                    column=column,
                    before_value=detected_format,
                    after_value="datetime",
                )
            )

    return log_entries

def clean_dataset(
    dataset: LoadedDataset,
    actions: Iterable[CleaningAction] = FOUNDATION_ACTIONS,
    *,
    text_case_columns: Iterable[str] = (),
    numeric_columns: Iterable[str] = (),
    date_columns: Iterable[str] = (),

) -> CleaningResult:
    """Apply selected cleaning actions to a copy of the dataset.

    The source DataFrame is never modified.
    """

    selected_actions = frozenset(actions)
    unsupported_actions = selected_actions - SUPPORTED_ACTIONS

    if unsupported_actions:
        action_names = ", ".join(
            sorted(action.value for action in unsupported_actions)
        )
        raise NotImplementedError(
            f"These cleaning actions are not implemented yet: {action_names}"
        )

    cleaned_data = dataset.data.copy(deep=True)
    log_entries: list[CleaningLogEntry] = []

    if CleaningAction.REMOVE_BLANK_ROWS in selected_actions:
        blank_mask = build_blank_mask(cleaned_data)
        blank_rows = blank_mask.all(axis=1)
        removed_count = int(blank_rows.sum())

        if removed_count:
            cleaned_data = cleaned_data.loc[~blank_rows].copy()

        log_entries.append(
            CleaningLogEntry(
                action=CleaningAction.REMOVE_BLANK_ROWS,
                description=(
                    f"Removed {removed_count} fully blank "
                    f"{'row' if removed_count == 1 else 'rows'}."
                ),
                affected_rows=removed_count,
            )
        )

    if CleaningAction.REMOVE_DUPLICATES in selected_actions:
        # Fully blank rows are not treated as business duplicates.
        blank_mask = build_blank_mask(cleaned_data)
        blank_rows = blank_mask.all(axis=1)

        duplicate_rows = (
            cleaned_data.duplicated(keep="first")
            & ~blank_rows
        )
        removed_count = int(duplicate_rows.sum())

        if removed_count:
            cleaned_data = cleaned_data.loc[~duplicate_rows].copy()

        log_entries.append(
            CleaningLogEntry(
                action=CleaningAction.REMOVE_DUPLICATES,
                description=(
                    f"Removed {removed_count} exact duplicate "
                    f"{'row' if removed_count == 1 else 'rows'}."
                ),
                affected_rows=removed_count,
            )
        )
    if CleaningAction.REMOVE_BLANK_COLUMNS in selected_actions:
        blank_mask = build_blank_mask(cleaned_data)

        blank_columns = [
            column
            for column in cleaned_data.columns
            if bool(blank_mask[column].all())
        ]

        row_count = len(cleaned_data.index)

        for column in blank_columns:
            cleaned_data.drop(columns=[column], inplace=True)

            log_entries.append(
                CleaningLogEntry(
                    action=CleaningAction.REMOVE_BLANK_COLUMNS,
                    description=(
                        f"Removed fully blank column '{column}'."
                    ),
                    affected_cells=row_count,
                    column=str(column),
                    before_value=str(column),
                    after_value=None,
                )
            )
    if CleaningAction.TRIM_WHITESPACE in selected_actions:
        for column in cleaned_data.columns:
            whitespace_mask = cleaned_data[column].map(
                has_outer_whitespace
            )
            changed_count = int(whitespace_mask.sum())

            if changed_count == 0:
                continue

            cleaned_data.loc[whitespace_mask, column] = (
                cleaned_data.loc[whitespace_mask, column].map(
                    str.strip
                )
            )

            log_entries.append(
                CleaningLogEntry(
                    action=CleaningAction.TRIM_WHITESPACE,
                    description=(
                        f"Trimmed outer whitespace from "
                        f"{changed_count} "
                        f"{'cell' if changed_count == 1 else 'cells'} "
                        f"in column '{column}'."
                    ),
                    affected_cells=changed_count,
                    column=str(column),
                )
            )
    if CleaningAction.STANDARDIZE_TEXT_CASE in selected_actions:
            selected_case_columns = tuple(text_case_columns)

            if not selected_case_columns:
                raise ValueError(
                    "Select at least one column for "
                    "text-case standardization."
                )

            case_log_entries = standardize_selected_text_case(
                cleaned_data,
                selected_case_columns,
            )

            log_entries.extend(case_log_entries)

    if CleaningAction.CONVERT_DATES in selected_actions:
        selected_date_columns = tuple(date_columns)

        if not selected_date_columns:
            raise ValueError(
                "Select at least one column for date conversion."
            )

        date_log_entries = convert_selected_date_text(
            cleaned_data,
            selected_date_columns,
        )

        log_entries.extend(date_log_entries)

    if CleaningAction.CONVERT_NUMERIC in selected_actions:
        selected_numeric_columns = tuple(numeric_columns)

        if not selected_numeric_columns:
            raise ValueError(
                "Select at least one column for numeric conversion."
            )

        numeric_log_entries = convert_selected_numeric_text(
            cleaned_data,
            selected_numeric_columns,
        )

        log_entries.extend(numeric_log_entries)      
    
    if CleaningAction.STANDARDIZE_COLUMN_NAMES in selected_actions:
        suggestions, collisions = find_column_name_suggestions(
            cleaned_data
        )

        rename_map = {
            original: suggestion
            for original, suggestion in suggestions.items()
            if suggestion not in collisions
        }

        if rename_map:
            cleaned_data.rename(
                columns=rename_map,
                inplace=True,
            )

        for original, suggestion in rename_map.items():
            log_entries.append(
                CleaningLogEntry(
                    action=CleaningAction.STANDARDIZE_COLUMN_NAMES,
                    description=(
                        f"Renamed column '{original}' "
                        f"to '{suggestion}'."
                    ),
                    affected_cells=1,
                    column=original,
                    before_value=original,
                    after_value=suggestion,
                )
            )

    cleaned_data.reset_index(drop=True, inplace=True)
    cleaned_data.reset_index(drop=True, inplace=True)

    return CleaningResult(
        source_name=dataset.source_path.name,
        data=cleaned_data,
        original_row_count=dataset.row_count,
        original_column_count=dataset.column_count,
        log_entries=tuple(log_entries),
    )