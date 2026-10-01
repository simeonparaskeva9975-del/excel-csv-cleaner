"""Desktop interface for the Excel/CSV Cleaner."""

from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any

import pandas as pd
import ttkbootstrap as ttk
from tksheet import Sheet

from cleaner_app import __version__
from cleaner_app.analysis import analyze_dataset
from cleaner_app.cleaning import clean_dataset
from cleaner_app.file_io import (
    DatasetLoadError,
    LoadedDataset,
    load_dataset,
)
from cleaner_app.models import (
    CleaningAction,
    CleaningResult,
    DatasetProfile,
)
from cleaner_app.reporting import (
    export_cleaned_csv,
    export_excel_report,
)
from cleaner_app.styles import (
    BACKGROUND,
    BORDER,
    ERROR,
    SURFACE,
    SUCCESS,
    TEXT_PRIMARY,
    TEXT_SECONDARY,
    WARNING,
    apply_application_styles,
)


class CleanerApplication:
    """Main desktop application window."""

    WINDOW_WIDTH = 1180
    WINDOW_HEIGHT = 760
    PREVIEW_ROWS = 20

    def __init__(self, root: ttk.Window) -> None:
        self.root = root
        self.loaded_dataset: LoadedDataset | None = None
        self.profile: DatasetProfile | None = None
        self.cleaning_result: CleaningResult | None = None
        self.after_profile: DatasetProfile | None = None
        self.safe_action_vars: dict[CleaningAction, tk.BooleanVar] = {}
        self.text_case_vars: dict[str, tk.BooleanVar] = {}
        self.numeric_vars: dict[str, tk.BooleanVar] = {}
        self.date_vars: dict[str, tk.BooleanVar] = {}
        self.last_export_directory: Path | None = None
        self.last_directory = Path.cwd()
        self.step_number_labels: list[ttk.Label] = []
        self.step_text_labels: list[ttk.Label] = []

        self.root.title("Excel/CSV Cleaner + Report Generator")
        self.root.minsize(980, 650)
        self.root.configure(background=BACKGROUND)

        apply_application_styles(self.root.style)

        self.status_var = tk.StringVar(
            master=self.root,
            value="Ready — select an Excel or CSV file to begin.",
        )

        self._configure_window()
        self._build_interface()

    def _configure_window(self) -> None:
        """Set the initial size and center the window."""

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x_position = max(
            0,
            (screen_width - self.WINDOW_WIDTH) // 2,
        )
        y_position = max(
            0,
            (screen_height - self.WINDOW_HEIGHT) // 2,
        )

        self.root.geometry(
            f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}"
            f"+{x_position}+{y_position}"
        )

        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

    def _build_interface(self) -> None:
        """Construct the main application layout."""

        main_container = ttk.Frame(
            self.root,
            style="App.TFrame",
            padding=(30, 24, 30, 18),
        )
        main_container.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(2, weight=1)

        self._build_header(main_container)
        self._build_workflow_steps(main_container)
        self._build_workspace(main_container)
        self._build_status_bar(main_container)

    def _build_header(self, parent: ttk.Frame) -> None:
        """Create the application heading."""

        header = ttk.Frame(parent, style="App.TFrame")
        header.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(0, 22),
        )
        header.columnconfigure(1, weight=1)

        logo = ttk.Label(
            header,
            text="DQ",
            style="StepActive.TLabel",
            width=4,
        )
        logo.grid(
            row=0,
            column=0,
            rowspan=2,
            sticky="w",
            padx=(0, 14),
        )

        ttk.Label(
            header,
            text="Excel/CSV Cleaner",
            style="Title.TLabel",
        ).grid(
            row=0,
            column=1,
            sticky="w",
        )

        ttk.Label(
            header,
            text=(
                "Analyze, clean and export professional "
                "data-quality reports."
            ),
            style="Subtitle.TLabel",
        ).grid(
            row=1,
            column=1,
            sticky="w",
            pady=(3, 0),
        )

        ttk.Label(
            header,
            text=f"Version {__version__}",
            style="Version.TLabel",
        ).grid(
            row=0,
            column=2,
            rowspan=2,
            sticky="e",
        )

    def _build_workflow_steps(
        self,
        parent: ttk.Frame,
    ) -> None:
        """Create the four-stage workflow indicator."""

        workflow = ttk.Frame(parent, style="App.TFrame")
        workflow.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(0, 22),
        )

        for column in range(7):
            workflow.columnconfigure(
                column,
                weight=0 if column % 2 == 0 else 1,
            )

        steps = (
            ("1", "Select File"),
            ("2", "Analyze"),
            ("3", "Clean"),
            ("4", "Export"),
        )

        self.step_number_labels.clear()
        self.step_text_labels.clear()

        for index, (number, title) in enumerate(steps):
            step_column = index * 2
            is_active = index == 0

            step_container = ttk.Frame(
                workflow,
                style="App.TFrame",
            )
            step_container.grid(
                row=0,
                column=step_column,
                sticky="w",
            )

            number_label = ttk.Label(
                step_container,
                text=number,
                width=3,
                style=(
                    "StepActive.TLabel"
                    if is_active
                    else "StepInactive.TLabel"
                ),
            )
            number_label.pack(
                side="left",
                padx=(0, 8),
            )

            text_label = ttk.Label(
                step_container,
                text=title,
                style=(
                    "StepTextActive.TLabel"
                    if is_active
                    else "StepTextInactive.TLabel"
                ),
            )
            text_label.pack(side="left")

            self.step_number_labels.append(number_label)
            self.step_text_labels.append(text_label)

            if index < len(steps) - 1:
                ttk.Separator(
                    workflow,
                    orient="horizontal",
                ).grid(
                    row=0,
                    column=step_column + 1,
                    sticky="ew",
                    padx=16,
                )

    def _set_active_step(self, active_step: int) -> None:
        """Highlight one stage in the four-step workflow."""

        for index, (number_label, text_label) in enumerate(
            zip(
                self.step_number_labels,
                self.step_text_labels,
                strict=True,
            ),
            start=1,
        ):
            is_active = index == active_step
            number_label.configure(
                style=(
                    "StepActive.TLabel"
                    if is_active
                    else "StepInactive.TLabel"
                )
            )
            text_label.configure(
                style=(
                    "StepTextActive.TLabel"
                    if is_active
                    else "StepTextInactive.TLabel"
                )
            )

    def _build_workspace(self, parent: ttk.Frame) -> None:
        """Create the replaceable central workspace."""

        self.workspace = ttk.Frame(
            parent,
            style="App.TFrame",
        )
        self.workspace.grid(
            row=2,
            column=0,
            sticky="nsew",
        )
        self.workspace.rowconfigure(0, weight=1)
        self.workspace.columnconfigure(0, weight=1)

        self._show_empty_workspace()

    def _clear_workspace(self) -> None:
        """Remove all widgets from the central workspace."""

        for child in self.workspace.winfo_children():
            child.destroy()

    def _show_empty_workspace(self) -> None:
        """Show the initial file-selection state."""

        self._clear_workspace()

        card = ttk.Frame(
            self.workspace,
            style="Card.TFrame",
            padding=36,
        )
        card.grid(
            row=0,
            column=0,
            sticky="nsew",
        )
        card.columnconfigure(0, weight=1)
        card.rowconfigure(0, weight=1)
        card.rowconfigure(4, weight=1)

        content = ttk.Frame(card, style="Surface.TFrame")
        content.grid(row=2, column=0)

        ttk.Label(
            content,
            text="Start with your dataset",
            style="Heading.TLabel",
        ).pack(pady=(0, 8))

        ttk.Label(
            content,
            text=(
                "Select an Excel workbook or CSV file. "
                "Your original file will never be modified."
            ),
            style="Body.TLabel",
            justify="center",
        ).pack(pady=(0, 20))

        ttk.Button(
            content,
            text="Select File",
            style="Primary.TButton",
            command=self._select_file,
        ).pack()

        ttk.Label(
            content,
            text="Supported formats: .xlsx and .csv",
            style="Small.TLabel",
        ).pack(pady=(12, 0))

    def _select_file(self) -> None:
        """Ask the user to select and load a dataset."""

        selected_path = filedialog.askopenfilename(
            parent=self.root,
            title="Select an Excel or CSV file",
            initialdir=self.last_directory,
            filetypes=(
                ("Supported datasets", "*.xlsx *.csv"),
                ("Excel workbooks", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ),
        )

        if not selected_path:
            self.status_var.set("File selection cancelled.")
            return

        path = Path(selected_path)
        self.last_directory = path.parent
        self.status_var.set(f"Loading {path.name}...")
        self.root.update_idletasks()

        try:
            dataset = load_dataset(path)
        except DatasetLoadError as exc:
            self.status_var.set("The selected file could not be loaded.")
            messagebox.showerror(
                "Unable to Load File",
                str(exc),
                parent=self.root,
            )
            return
        except Exception as exc:
            self.status_var.set("An unexpected loading error occurred.")
            messagebox.showerror(
                "Unexpected Error",
                (
                    "The selected file could not be loaded.\n\n"
                    f"Details: {exc}"
                ),
                parent=self.root,
            )
            return

        self.loaded_dataset = dataset
        self.profile = None
        self.cleaning_result = None
        self.after_profile = None
        self._set_active_step(1)
        self._show_loaded_workspace(dataset)
        self.status_var.set(
            f"Loaded {dataset.source_path.name} successfully."
        )

    def _show_loaded_workspace(
        self,
        dataset: LoadedDataset,
    ) -> None:
        """Show file information and a scrollable data preview."""

        self._clear_workspace()

        card = ttk.Frame(
            self.workspace,
            style="Card.TFrame",
            padding=22,
        )
        card.grid(
            row=0,
            column=0,
            sticky="nsew",
        )
        card.columnconfigure(0, weight=1)
        card.rowconfigure(3, weight=1)

        file_header = ttk.Frame(
            card,
            style="Surface.TFrame",
        )
        file_header.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(0, 16),
        )
        file_header.columnconfigure(0, weight=1)

        file_text = ttk.Frame(
            file_header,
            style="Surface.TFrame",
        )
        file_text.grid(row=0, column=0, sticky="w")

        ttk.Label(
            file_text,
            text=dataset.source_path.name,
            style="FileName.TLabel",
        ).pack(anchor="w")

        metadata_parts = [
            dataset.file_type,
            self._format_file_size(dataset.file_size_bytes),
        ]

        if dataset.selected_sheet:
            metadata_parts.append(
                f"Worksheet: {dataset.selected_sheet}"
            )

        ttk.Label(
            file_text,
            text="  •  ".join(metadata_parts),
            style="FileMeta.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        button_group = ttk.Frame(
            file_header,
            style="Surface.TFrame",
        )
        button_group.grid(row=0, column=1, sticky="e")

        ttk.Button(
            button_group,
            text="Choose Another File",
            style="Secondary.TButton",
            command=self._select_file,
        ).pack(side="left", padx=(0, 10))

        ttk.Button(
            button_group,
            text="Analyze Data",
            style="Primary.TButton",
            command=self._analyze_data,
        ).pack(side="left")

        metrics = ttk.Frame(
            card,
            style="Surface.TFrame",
        )
        metrics.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(0, 16),
        )

        for column in range(4):
            metrics.columnconfigure(column, weight=1)

        metric_values = (
            ("Rows", f"{dataset.row_count:,}"),
            ("Columns", f"{dataset.column_count:,}"),
            ("File Type", dataset.file_type),
            ("File Size", self._format_file_size(
                dataset.file_size_bytes
            )),
        )

        for column, (label, value) in enumerate(metric_values):
            self._create_metric_card(
                metrics,
                column,
                label,
                value,
            )

        preview_header = ttk.Frame(
            card,
            style="Surface.TFrame",
        )
        preview_header.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(0, 8),
        )
        preview_header.columnconfigure(0, weight=1)

        ttk.Label(
            preview_header,
            text="Data Preview",
            style="Heading.TLabel",
        ).grid(row=0, column=0, sticky="w")

        shown_rows = min(
            self.PREVIEW_ROWS,
            dataset.row_count,
        )

        ttk.Label(
            preview_header,
            text=(
                f"Showing {shown_rows:,} of "
                f"{dataset.row_count:,} rows"
            ),
            style="Small.TLabel",
        ).grid(row=0, column=1, sticky="e")

        self._create_preview_table(
            card,
            dataset,
        )

    def _create_preview_table(
        self,
        parent: ttk.Frame,
        dataset: LoadedDataset,
    ) -> None:
        """Create a smooth, read-only spreadsheet preview."""

        preview = dataset.data.head(self.PREVIEW_ROWS)

        headers = [
            str(column)
            for column in preview.columns
        ]

        table_data = [
            [
                self._display_value(value)
                for value in row
            ]
            for row in preview.itertuples(
                index=False,
                name=None,
            )
        ]

        table_container = tk.Frame(
            parent,
            background="#FFFFFF",
            highlightthickness=0,
        )
        table_container.grid(
            row=3,
            column=0,
            sticky="nsew",
        )
        table_container.rowconfigure(0, weight=1)
        table_container.columnconfigure(0, weight=1)

        self.preview_sheet = Sheet(
            table_container,
            data=table_data,
            headers=headers,
            width=1_040,
            height=310,
            show_row_index=True,
            show_header=True,
            show_x_scrollbar=True,
            show_y_scrollbar=True,
            default_column_width=140,
            default_row_index_width=52,
            default_row_height=28,
            default_header_height=34,
            align="w",
            header_align="w",
            row_index_align="center",
            table_wrap="",
            header_wrap="",
            empty_horizontal=0,
            empty_vertical=0,
            font=("Segoe UI", 9, "normal"),
            header_font=("Segoe UI", 9, "bold"),
            theme="light blue",
            scrollbar_theme_inheritance="clam",
        )
        self.preview_sheet.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        self.preview_sheet.enable_bindings(
            "single_select",
            "drag_select",
            "select_all",
            "column_select",
            "row_select",
            "rc_select",
            "arrowkeys",
            "copy",
            "column_width_resize",
            "double_click_column_resize",
        )

        for column_index, header in enumerate(headers):
            column_values = [
                row[column_index]
                for row in table_data
            ]

            longest_length = max(
                [len(header)]
                + [len(value) for value in column_values]
            )

            width = min(
                max(longest_length * 8 + 24, 110),
                240,
            )

            self.preview_sheet.column_width(
                column=column_index,
                width=width,
                redraw=False,
            )

        self.preview_sheet.redraw()

    def _create_metric_card(
        self,
        parent: ttk.Frame,
        column: int,
        label: str,
        value: str,
    ) -> None:
        """Create one compact file-summary card."""

        metric = ttk.Frame(
            parent,
            style="Card.TFrame",
            padding=(14, 10),
        )
        metric.grid(
            row=0,
            column=column,
            sticky="ew",
            padx=(
                0 if column == 0 else 5,
                0 if column == 3 else 5,
            ),
        )

        ttk.Label(
            metric,
            text=value,
            style="MetricValue.TLabel",
        ).pack(anchor="w")

        ttk.Label(
            metric,
            text=label,
            style="MetricName.TLabel",
        ).pack(anchor="w", pady=(2, 0))

    @staticmethod
    def _display_value(value: Any) -> str:
        """Convert a table value to safe preview text."""

        try:
            if pd.isna(value):
                return ""
        except (TypeError, ValueError):
            pass

        if hasattr(value, "strftime"):
            try:
                return value.strftime("%Y-%m-%d")
            except (TypeError, ValueError):
                pass

        return str(value)

    @staticmethod
    def _format_file_size(size_bytes: int) -> str:
        """Return a human-readable file size."""

        if size_bytes < 1_024:
            return f"{size_bytes} B"

        size_kb = size_bytes / 1_024

        if size_kb < 1_024:
            return f"{size_kb:.1f} KB"

        size_mb = size_kb / 1_024
        return f"{size_mb:.1f} MB"

    def _analyze_data(self) -> None:
        """Run non-destructive analysis and show its findings."""

        if self.loaded_dataset is None:
            return

        self.status_var.set("Analyzing data quality...")
        self.root.configure(cursor="watch")
        self.root.update_idletasks()

        try:
            profile = analyze_dataset(self.loaded_dataset)
        except Exception as exc:
            self.status_var.set("Data analysis failed.")
            messagebox.showerror(
                "Analysis Error",
                (
                    "The dataset could not be analyzed.\n\n"
                    f"Details: {exc}"
                ),
                parent=self.root,
            )
            return
        finally:
            self.root.configure(cursor="")

        self.profile = profile
        self._set_active_step(2)
        self._show_analysis_workspace(profile)
        self.status_var.set(
            f"Analysis complete — {len(profile.issues):,} "
            "quality findings detected."
        )

    def _show_analysis_workspace(
        self,
        profile: DatasetProfile,
    ) -> None:
        """Display summary metrics and detailed analysis findings."""

        self._clear_workspace()

        card = ttk.Frame(
            self.workspace,
            style="Card.TFrame",
            padding=22,
        )
        card.grid(row=0, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)

        header = ttk.Frame(card, style="Surface.TFrame")
        header.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(0, 16),
        )
        header.columnconfigure(0, weight=1)

        heading_group = ttk.Frame(
            header,
            style="Surface.TFrame",
        )
        heading_group.grid(row=0, column=0, sticky="w")

        ttk.Label(
            heading_group,
            text="Data Quality Analysis",
            style="FileName.TLabel",
        ).pack(anchor="w")

        ttk.Label(
            heading_group,
            text=(
                f"Results for {profile.source_name} — "
                "the source file has not been modified."
            ),
            style="FileMeta.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        button_group = ttk.Frame(
            header,
            style="Surface.TFrame",
        )
        button_group.grid(row=0, column=1, sticky="e")

        ttk.Button(
            button_group,
            text="Back to Preview",
            style="Secondary.TButton",
            command=self._return_to_preview,
        ).pack(side="left", padx=(0, 10))

        ttk.Button(
            button_group,
            text="Review Cleaning Options",
            style="Primary.TButton",
            command=self._show_cleaning_options,
        ).pack(side="left")

        safe_findings = sum(
            issue.safe_to_fix
            for issue in profile.issues
        )
        review_findings = len(profile.issues) - safe_findings
        missing_cells = sum(profile.missing_by_column.values())

        metrics = ttk.Frame(card, style="Surface.TFrame")
        metrics.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(0, 16),
        )

        for column in range(4):
            metrics.columnconfigure(column, weight=1)

        metric_values = (
            ("Quality Findings", f"{len(profile.issues):,}"),
            ("Recommended Fixes", f"{safe_findings:,}"),
            ("Review Required", f"{review_findings:,}"),
            ("Missing Cells", f"{missing_cells:,}"),
        )

        for column, (label, value) in enumerate(metric_values):
            self._create_metric_card(
                metrics,
                column,
                label,
                value,
            )

        notebook = ttk.Notebook(card)
        notebook.grid(
            row=2,
            column=0,
            sticky="nsew",
        )

        issues_tab = ttk.Frame(
            notebook,
            style="Surface.TFrame",
            padding=10,
        )
        columns_tab = ttk.Frame(
            notebook,
            style="Surface.TFrame",
            padding=10,
        )

        notebook.add(
            issues_tab,
            text=f"Detected Issues ({len(profile.issues)})",
        )
        notebook.add(
            columns_tab,
            text="Column Summary",
        )

        self._build_issues_table(issues_tab, profile)
        self._build_column_summary_table(columns_tab, profile)

    def _build_issues_table(
        self,
        parent: ttk.Frame,
        profile: DatasetProfile,
    ) -> None:
        """Create a concise issues table with full selected details."""

        parent.rowconfigure(0, weight=1, minsize=125)
        parent.columnconfigure(0, weight=1)

        column_ids = (
            "severity",
            "issue",
            "column",
            "count",
            "decision",
        )

        table_frame = ttk.Frame(
            parent,
            style="Surface.TFrame",
        )
        table_frame.grid(row=0, column=0, sticky="nsew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        tree = ttk.Treeview(
            table_frame,
            columns=column_ids,
            show="headings",
            height=5,
            style="Data.Treeview",
        )
        tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=tree.yview,
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)

        self._configure_responsive_columns(
            tree,
            (
                ("severity", "Severity", 100, 1, "center"),
                ("issue", "Issue", 190, 3, "w"),
                ("column", "Column", 190, 3, "w"),
                ("count", "Count", 80, 1, "center"),
                ("decision", "Action", 130, 2, "center"),
            ),
        )

        tree.tag_configure("info", foreground=TEXT_SECONDARY)
        tree.tag_configure("warning", foreground=WARNING)
        tree.tag_configure("error", foreground=ERROR)
        tree.tag_configure("safe", foreground=SUCCESS)

        issue_by_item: dict[str, Any] = {}

        for issue in profile.issues:
            tag = issue.severity.value
            if issue.safe_to_fix and issue.severity.value == "info":
                tag = "safe"

            item_id = tree.insert(
                "",
                "end",
                values=(
                    issue.severity.value.title(),
                    issue.issue_type.replace("_", " ").title(),
                    issue.column or "—",
                    f"{issue.count:,}",
                    (
                        "Recommended"
                        if issue.safe_to_fix
                        else "Review"
                    ),
                ),
                tags=(tag,),
            )

            issue_by_item[item_id] = issue

        detail_card = ttk.Frame(
            parent,
            style="Card.TFrame",
            padding=(14, 9),
        )
        detail_card.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(10, 0),
        )
        detail_card.columnconfigure(0, weight=1)

        detail_title_var = tk.StringVar(
            master=self.root,
            value="Select a finding to view its full details.",
        )
        ttk.Label(
            detail_card,
            textvariable=detail_title_var,
            style="Heading.TLabel",
        ).grid(row=0, column=0, sticky="w")

        detail_text = self._create_detail_text(
            detail_card,
            row=1,
            height=3,
        )

        def show_selected_issue(_event: Any = None) -> None:
            selected_items = tree.selection()

            if not selected_items:
                return

            issue = issue_by_item[selected_items[0]]
            title = issue.issue_type.replace("_", " ").title()

            detail_parts = [issue.message]

            if issue.examples:
                detail_parts.append(
                    "Examples: " + ", ".join(issue.examples)
                )

            detail_parts.append(
                "Handling: "
                + (
                    "Recommended safe fix."
                    if issue.safe_to_fix
                    else "Review required before making changes."
                )
            )

            detail_title_var.set(title)
            self._set_detail_text(
                detail_text,
                "\n\n".join(detail_parts),
            )

        tree.bind("<<TreeviewSelect>>", show_selected_issue)

        first_item = tree.get_children()
        if first_item:
            tree.selection_set(first_item[0])
            tree.focus(first_item[0])
            show_selected_issue()

    def _build_column_summary_table(
        self,
        parent: ttk.Frame,
        profile: DatasetProfile,
    ) -> None:
        """Create a compact column table with wrapped details below."""

        parent.rowconfigure(0, weight=1, minsize=125)
        parent.columnconfigure(0, weight=1)

        column_ids = (
            "column",
            "dtype",
            "missing",
            "whitespace",
            "finding_count",
        )

        table_frame = ttk.Frame(
            parent,
            style="Surface.TFrame",
        )
        table_frame.grid(row=0, column=0, sticky="nsew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        tree = ttk.Treeview(
            table_frame,
            columns=column_ids,
            show="headings",
            height=5,
            style="Data.Treeview",
        )
        tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=tree.yview,
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)

        self._configure_responsive_columns(
            tree,
            (
                ("column", "Column", 220, 4, "w"),
                ("dtype", "Data Type", 140, 2, "w"),
                ("missing", "Missing", 110, 1, "center"),
                ("whitespace", "Whitespace", 125, 2, "center"),
                ("finding_count", "Issues", 90, 1, "center"),
            ),
        )

        issues_by_column: dict[str, set[str]] = {}

        for issue in profile.issues:
            if issue.column is None:
                continue

            label = issue.issue_type.replace("_", " ").title()
            issues_by_column.setdefault(issue.column, set()).add(label)

        findings_by_item: dict[str, tuple[str, list[str]]] = {}

        for column, dtype in profile.dtypes.items():
            findings = sorted(issues_by_column.get(column, set()))
            item_id = tree.insert(
                "",
                "end",
                values=(
                    column,
                    dtype,
                    f"{profile.missing_by_column.get(column, 0):,}",
                    f"{profile.whitespace_cells_by_column.get(column, 0):,}",
                    f"{len(findings):,}",
                ),
            )

            findings_by_item[item_id] = (column, findings)

        detail_card = ttk.Frame(
            parent,
            style="Card.TFrame",
            padding=(14, 9),
        )
        detail_card.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(10, 0),
        )
        detail_card.columnconfigure(0, weight=1)

        detail_title_var = tk.StringVar(
            master=self.root,
            value="Select a column to view its findings.",
        )
        ttk.Label(
            detail_card,
            textvariable=detail_title_var,
            style="Heading.TLabel",
        ).grid(row=0, column=0, sticky="w")

        detail_text = self._create_detail_text(
            detail_card,
            row=1,
            height=2,
        )

        def show_selected_column(_event: Any = None) -> None:
            selected_items = tree.selection()

            if not selected_items:
                return

            column, findings = findings_by_item[selected_items[0]]
            detail_title_var.set(f"Column: {column}")
            self._set_detail_text(
                detail_text,
                "Potential issues: "
                + (
                    ", ".join(findings)
                    if findings
                    else "No column-specific issues detected."
                ),
            )

        tree.bind("<<TreeviewSelect>>", show_selected_column)

        first_item = tree.get_children()
        if first_item:
            tree.selection_set(first_item[0])
            tree.focus(first_item[0])
            show_selected_column()

    def _configure_responsive_columns(
        self,
        tree: ttk.Treeview,
        specifications: tuple[
            tuple[str, str, int, int, str], ...
        ],
    ) -> None:
        """Keep headings separated and use all available table width."""

        for column_id, heading, minimum, _weight, anchor in specifications:
            tree.heading(column_id, text=heading, anchor=anchor)
            tree.column(
                column_id,
                width=minimum,
                minwidth=minimum,
                stretch=False,
                anchor=anchor,
            )

        def resize_columns(event: tk.Event) -> None:
            available_width = max(event.width - 4, 1)
            minimum_width = sum(item[2] for item in specifications)
            extra_width = max(available_width - minimum_width, 0)
            total_weight = sum(item[3] for item in specifications)

            assigned_extra = 0
            for index, (
                column_id,
                _heading,
                minimum,
                weight,
                _anchor,
            ) in enumerate(specifications):
                if index == len(specifications) - 1:
                    extra = extra_width - assigned_extra
                else:
                    extra = round(extra_width * weight / total_weight)
                    assigned_extra += extra

                tree.column(column_id, width=minimum + extra)

        tree.bind("<Configure>", resize_columns, add="+")

    def _create_detail_text(
        self,
        parent: ttk.Frame,
        row: int,
        height: int,
    ) -> tk.Text:
        """Create a readable, word-wrapped details panel."""

        text_frame = ttk.Frame(parent, style="Card.TFrame")
        text_frame.grid(
            row=row,
            column=0,
            sticky="ew",
            pady=(7, 0),
        )
        text_frame.columnconfigure(0, weight=1)

        detail_text = tk.Text(
            text_frame,
            height=height,
            wrap="word",
            font=("Segoe UI", 10),
            foreground=TEXT_SECONDARY,
            background=SURFACE,
            selectbackground="#DDE9F3",
            selectforeground=TEXT_PRIMARY,
            relief="flat",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            padx=10,
            pady=6,
            cursor="arrow",
        )
        detail_text.grid(row=0, column=0, sticky="ew")

        scrollbar = ttk.Scrollbar(
            text_frame,
            orient="vertical",
            command=detail_text.yview,
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        detail_text.configure(yscrollcommand=scrollbar.set)
        detail_text.configure(state="disabled")

        return detail_text

    @staticmethod
    def _set_detail_text(widget: tk.Text, value: str) -> None:
        """Replace the contents of a read-only details panel."""

        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", value)
        widget.configure(state="disabled")
        widget.yview_moveto(0.0)

    def _return_to_preview(self) -> None:
        """Return to the loaded-file preview without reloading it."""

        if self.loaded_dataset is None:
            return

        self._set_active_step(1)
        self._show_loaded_workspace(self.loaded_dataset)
        self.status_var.set("Returned to the source-data preview.")

    def _show_cleaning_options(self) -> None:
        """Show safe defaults and optional column-specific cleaning."""

        if self.profile is None or self.loaded_dataset is None:
            return

        self._clear_workspace()
        self._set_active_step(3)

        card = ttk.Frame(
            self.workspace,
            style="Card.TFrame",
            padding=22,
        )
        card.grid(row=0, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)

        header = ttk.Frame(card, style="Surface.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)

        heading_group = ttk.Frame(header, style="Surface.TFrame")
        heading_group.grid(row=0, column=0, sticky="w")

        ttk.Label(
            heading_group,
            text="Review Cleaning Options",
            style="FileName.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            heading_group,
            text=(
                "Safe fixes are preselected. Optional conversions require "
                "your approval."
            ),
            style="FileMeta.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        ttk.Button(
            header,
            text="Back to Analysis",
            style="Secondary.TButton",
            command=self._return_to_analysis,
        ).grid(row=0, column=1, sticky="e")

        content = ttk.Frame(card, style="Surface.TFrame")
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1, uniform="options")
        content.columnconfigure(1, weight=1, uniform="options")
        content.rowconfigure(0, weight=1)

        safe_card = ttk.Frame(
            content,
            style="Card.TFrame",
            padding=(18, 15),
        )
        safe_card.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 7),
        )
        safe_card.columnconfigure(0, weight=1)

        ttk.Label(
            safe_card,
            text="Recommended Safe Fixes",
            style="Heading.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            safe_card,
            text="Transparent changes that do not infer missing data.",
            style="Small.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 12))

        whitespace_count = sum(
            self.profile.whitespace_cells_by_column.values()
        )
        safe_options = (
            (
                CleaningAction.REMOVE_BLANK_ROWS,
                "Remove fully blank rows",
                f"{self.profile.blank_rows:,} rows detected",
            ),
            (
                CleaningAction.REMOVE_DUPLICATES,
                "Remove exact duplicate rows",
                f"{self.profile.duplicate_rows:,} duplicates detected",
            ),
            (
                CleaningAction.REMOVE_BLANK_COLUMNS,
                "Remove fully blank columns",
                f"{len(self.profile.blank_columns):,} columns detected",
            ),
            (
                CleaningAction.TRIM_WHITESPACE,
                "Trim outer whitespace",
                f"{whitespace_count:,} text cells detected",
            ),
            (
                CleaningAction.STANDARDIZE_COLUMN_NAMES,
                "Standardize column names",
                (
                    f"{len(self.profile.column_name_suggestions):,} "
                    "rename suggestions"
                ),
            ),
        )

        self.safe_action_vars = {}
        for row, (action, title, description) in enumerate(
            safe_options,
            start=2,
        ):
            variable = tk.BooleanVar(master=self.root, value=True)
            self.safe_action_vars[action] = variable
            self._create_cleaning_option(
                safe_card,
                row,
                variable,
                title,
                description,
            )

        review_card = ttk.Frame(
            content,
            style="Card.TFrame",
            padding=(18, 15),
        )
        review_card.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(7, 0),
        )
        review_card.columnconfigure(0, weight=1)

        ttk.Label(
            review_card,
            text="Review Required",
            style="Heading.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            review_card,
            text="Select only the columns you want converted.",
            style="Small.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 10))

        review_content = ttk.Frame(
            review_card,
            style="Surface.TFrame",
        )
        review_content.grid(row=2, column=0, sticky="nsew")
        review_content.columnconfigure(0, weight=1)

        self.text_case_vars = self._create_column_choice_group(
            review_content,
            0,
            "Standardize capitalization",
            tuple(self.profile.case_variants_by_column),
        )

        numeric_columns = tuple(
            column
            for column, finding in self.profile.numeric_findings.items()
            if finding.numeric_text_count > 0
        )
        self.numeric_vars = self._create_column_choice_group(
            review_content,
            1,
            "Convert numeric text",
            numeric_columns,
        )

        date_columns = tuple(
            column
            for column, finding in self.profile.date_findings.items()
            if finding.recognized_text_count > 0
        )
        self.date_vars = self._create_column_choice_group(
            review_content,
            2,
            "Convert unambiguous dates",
            date_columns,
        )

        notice = ttk.Label(
            review_card,
            text=(
                "Missing, ambiguous, invalid and suspicious values are "
                "reported only — they will not be changed automatically."
            ),
            style="Notice.TLabel",
            justify="left",
            wraplength=470,
        )
        notice.grid(row=3, column=0, sticky="ew", pady=(12, 0))

        footer = ttk.Frame(card, style="Surface.TFrame")
        footer.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        footer.columnconfigure(0, weight=1)

        ttk.Label(
            footer,
            text="The original source file will remain unchanged.",
            style="Small.TLabel",
        ).grid(row=0, column=0, sticky="w")

        ttk.Button(
            footer,
            text="Clean Data",
            style="Primary.TButton",
            command=self._clean_data,
        ).grid(row=0, column=1, sticky="e")

        self.status_var.set(
            "Review the selected cleaning actions before continuing."
        )

    def _create_cleaning_option(
        self,
        parent: ttk.Frame,
        row: int,
        variable: tk.BooleanVar,
        title: str,
        description: str,
    ) -> None:
        """Create one checkbox with a concise impact description."""

        option = ttk.Frame(parent, style="Surface.TFrame")
        option.grid(row=row, column=0, sticky="ew", pady=4)
        option.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            option,
            variable=variable,
            style="Option.TCheckbutton",
        ).grid(row=0, column=0, rowspan=2, sticky="n", padx=(0, 8))

        ttk.Label(
            option,
            text=title,
            style="OptionTitle.TLabel",
        ).grid(row=0, column=1, sticky="w")
        ttk.Label(
            option,
            text=description,
            style="Small.TLabel",
        ).grid(row=1, column=1, sticky="w", pady=(1, 0))

    def _create_column_choice_group(
        self,
        parent: ttk.Frame,
        row: int,
        title: str,
        columns: tuple[str, ...],
    ) -> dict[str, tk.BooleanVar]:
        """Create an optional conversion group for detected columns."""

        group = ttk.Frame(parent, style="Surface.TFrame")
        group.grid(row=row, column=0, sticky="ew", pady=(0, 9))
        group.columnconfigure(0, weight=1)

        ttk.Label(
            group,
            text=title,
            style="OptionTitle.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(0, 3))

        variables: dict[str, tk.BooleanVar] = {}

        if not columns:
            ttk.Label(
                group,
                text="No applicable columns detected.",
                style="Small.TLabel",
            ).grid(row=1, column=0, sticky="w")
            return variables

        choices = ttk.Frame(group, style="Surface.TFrame")
        choices.grid(row=1, column=0, sticky="ew")

        for index, column in enumerate(columns):
            variable = tk.BooleanVar(master=self.root, value=False)
            variables[column] = variable
            ttk.Checkbutton(
                choices,
                text=column,
                variable=variable,
                style="Option.TCheckbutton",
            ).grid(
                row=index // 2,
                column=index % 2,
                sticky="w",
                padx=(0, 14),
                pady=1,
            )

        return variables

    def _clean_data(self) -> None:
        """Apply selected actions to a copy and re-analyze the result."""

        if self.loaded_dataset is None or self.profile is None:
            return

        actions = {
            action
            for action, variable in self.safe_action_vars.items()
            if variable.get()
        }
        text_case_columns = tuple(
            column
            for column, variable in self.text_case_vars.items()
            if variable.get()
        )
        numeric_columns = tuple(
            column
            for column, variable in self.numeric_vars.items()
            if variable.get()
        )
        date_columns = tuple(
            column
            for column, variable in self.date_vars.items()
            if variable.get()
        )

        if text_case_columns:
            actions.add(CleaningAction.STANDARDIZE_TEXT_CASE)
        if numeric_columns:
            actions.add(CleaningAction.CONVERT_NUMERIC)
        if date_columns:
            actions.add(CleaningAction.CONVERT_DATES)

        self.status_var.set("Cleaning a copy of the dataset...")
        self.root.configure(cursor="watch")
        self.root.update_idletasks()

        try:
            result = clean_dataset(
                self.loaded_dataset,
                actions=actions,
                text_case_columns=text_case_columns,
                numeric_columns=numeric_columns,
                date_columns=date_columns,
            )
            cleaned_dataset = LoadedDataset(
                data=result.data,
                source_path=Path(
                    f"cleaned_{self.loaded_dataset.source_path.name}"
                ),
                file_type=self.loaded_dataset.file_type,
                file_size_bytes=0,
                selected_sheet=self.loaded_dataset.selected_sheet,
            )
            after_profile = analyze_dataset(cleaned_dataset)
        except Exception as exc:
            self.status_var.set("Data cleaning failed.")
            messagebox.showerror(
                "Cleaning Error",
                (
                    "The selected cleaning actions could not be applied."
                    "\n\n"
                    f"Details: {exc}"
                ),
                parent=self.root,
            )
            return
        finally:
            self.root.configure(cursor="")

        self.cleaning_result = result
        self.after_profile = after_profile
        self._show_cleaning_results(result, after_profile)
        self.status_var.set(
            f"Cleaning complete — {len(result.log_entries):,} "
            "operations were logged."
        )

    def _show_cleaning_results(
        self,
        result: CleaningResult,
        after_profile: DatasetProfile,
    ) -> None:
        """Show before/after metrics and a preview of cleaned data."""

        self._clear_workspace()
        self._set_active_step(3)

        card = ttk.Frame(
            self.workspace,
            style="Card.TFrame",
            padding=22,
        )
        card.grid(row=0, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(3, weight=1)

        header = ttk.Frame(card, style="Surface.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)

        heading_group = ttk.Frame(header, style="Surface.TFrame")
        heading_group.grid(row=0, column=0, sticky="w")

        ttk.Label(
            heading_group,
            text="Cleaned Data Preview",
            style="FileName.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            heading_group,
            text=(
                "Review the cleaned copy before exporting. "
                "The source file is unchanged."
            ),
            style="FileMeta.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        button_group = ttk.Frame(header, style="Surface.TFrame")
        button_group.grid(row=0, column=1, sticky="e")
        ttk.Button(
            button_group,
            text="Back to Options",
            style="Secondary.TButton",
            command=self._show_cleaning_options,
        ).pack(side="left", padx=(0, 10))
        ttk.Button(
            button_group,
            text="Export Results",
            style="Primary.TButton",
            command=self._show_export_workspace,
        ).pack(side="left")

        metrics = ttk.Frame(card, style="Surface.TFrame")
        metrics.grid(row=1, column=0, sticky="ew", pady=(0, 14))
        for column in range(4):
            metrics.columnconfigure(column, weight=1)

        metric_values = (
            (
                "Rows Before → After",
                f"{result.original_row_count:,} → {result.cleaned_row_count:,}",
            ),
            (
                "Columns Before → After",
                (
                    f"{result.original_column_count:,} → "
                    f"{result.cleaned_column_count:,}"
                ),
            ),
            ("Cleaning Log Entries", f"{len(result.log_entries):,}"),
            ("Remaining Findings", f"{len(after_profile.issues):,}"),
        )
        for column, (label, value) in enumerate(metric_values):
            self._create_metric_card(metrics, column, label, value)

        preview_header = ttk.Frame(card, style="Surface.TFrame")
        preview_header.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        preview_header.columnconfigure(0, weight=1)
        ttk.Label(
            preview_header,
            text="Cleaned Data",
            style="Heading.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            preview_header,
            text=(
                f"Showing {min(self.PREVIEW_ROWS, result.cleaned_row_count):,} "
                f"of {result.cleaned_row_count:,} rows"
            ),
            style="Small.TLabel",
        ).grid(row=0, column=1, sticky="e")

        preview_dataset = LoadedDataset(
            data=result.data,
            source_path=Path(result.source_name),
            file_type=self.loaded_dataset.file_type,
            file_size_bytes=0,
        )
        self._create_preview_table(card, preview_dataset)

    def _return_to_analysis(self) -> None:
        """Return to the current non-destructive analysis results."""

        if self.profile is None:
            return

        self._set_active_step(2)
        self._show_analysis_workspace(self.profile)
        self.status_var.set("Returned to the analysis results.")

    def _show_export_workspace(self) -> None:
        """Show professional Excel and optional CSV export controls."""

        if (
            self.loaded_dataset is None
            or self.profile is None
            or self.cleaning_result is None
            or self.after_profile is None
        ):
            return

        self._clear_workspace()
        self._set_active_step(4)

        card = ttk.Frame(
            self.workspace,
            style="Card.TFrame",
            padding=22,
        )
        card.grid(row=0, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)

        header = ttk.Frame(card, style="Surface.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        header.columnconfigure(0, weight=1)

        heading_group = ttk.Frame(header, style="Surface.TFrame")
        heading_group.grid(row=0, column=0, sticky="w")
        ttk.Label(
            heading_group,
            text="Export Results",
            style="FileName.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            heading_group,
            text="Choose the deliverables and where they should be saved.",
            style="FileMeta.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        ttk.Button(
            header,
            text="Back to Cleaned Preview",
            style="Secondary.TButton",
            command=lambda: self._show_cleaning_results(
                self.cleaning_result,
                self.after_profile,
            ),
        ).grid(row=0, column=1, sticky="e")

        metrics = ttk.Frame(card, style="Surface.TFrame")
        metrics.grid(row=1, column=0, sticky="ew", pady=(0, 16))
        for column in range(4):
            metrics.columnconfigure(column, weight=1)

        metric_values = (
            ("Cleaned Rows", f"{self.cleaning_result.cleaned_row_count:,}"),
            (
                "Cleaned Columns",
                f"{self.cleaning_result.cleaned_column_count:,}",
            ),
            (
                "Cleaning Log Entries",
                f"{len(self.cleaning_result.log_entries):,}",
            ),
            ("Remaining Findings", f"{len(self.after_profile.issues):,}"),
        )
        for column, (label, value) in enumerate(metric_values):
            self._create_metric_card(metrics, column, label, value)

        export_card = ttk.Frame(
            card,
            style="Card.TFrame",
            padding=(20, 17),
        )
        export_card.grid(row=2, column=0, sticky="nsew")
        export_card.columnconfigure(1, weight=1)

        default_directory = (
            self.loaded_dataset.source_path.parent / "cleaned_output"
        )
        source_stem = self.loaded_dataset.source_path.stem

        self.export_directory_var = tk.StringVar(
            master=self.root,
            value=str(default_directory),
        )
        self.export_excel_var = tk.BooleanVar(
            master=self.root,
            value=True,
        )
        self.export_csv_var = tk.BooleanVar(
            master=self.root,
            value=False,
        )
        self.excel_filename_var = tk.StringVar(
            master=self.root,
            value=f"{source_stem}_cleaned_report.xlsx",
        )
        self.csv_filename_var = tk.StringVar(
            master=self.root,
            value=f"{source_stem}_cleaned.csv",
        )

        ttk.Label(
            export_card,
            text="Destination Folder",
            style="OptionTitle.TLabel",
        ).grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Entry(
            export_card,
            textvariable=self.export_directory_var,
        ).grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(6, 16),
            ipady=6,
        )
        ttk.Button(
            export_card,
            text="Browse...",
            style="Secondary.TButton",
            command=self._choose_export_directory,
        ).grid(row=1, column=2, sticky="e", padx=(10, 0), pady=(6, 16))

        ttk.Separator(export_card).grid(
            row=2,
            column=0,
            columnspan=3,
            sticky="ew",
            pady=(0, 14),
        )

        excel_check = ttk.Checkbutton(
            export_card,
            text="Excel report",
            variable=self.export_excel_var,
            style="Option.TCheckbutton",
            command=self._update_export_controls,
        )
        excel_check.grid(row=3, column=0, sticky="w")
        ttk.Label(
            export_card,
            text=(
                "Cleaned Data, Data Quality Summary, Cleaning Log and "
                "meaningful Charts"
            ),
            style="Small.TLabel",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(2, 5))
        self.excel_filename_entry = ttk.Entry(
            export_card,
            textvariable=self.excel_filename_var,
        )
        self.excel_filename_entry.grid(
            row=5,
            column=0,
            columnspan=3,
            sticky="ew",
            ipady=5,
            pady=(0, 14),
        )

        csv_check = ttk.Checkbutton(
            export_card,
            text="Also export cleaned CSV",
            variable=self.export_csv_var,
            style="Option.TCheckbutton",
            command=self._update_export_controls,
        )
        csv_check.grid(row=6, column=0, sticky="w")
        ttk.Label(
            export_card,
            text="UTF-8 CSV containing only the cleaned dataset.",
            style="Small.TLabel",
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(2, 5))
        self.csv_filename_entry = ttk.Entry(
            export_card,
            textvariable=self.csv_filename_var,
        )
        self.csv_filename_entry.grid(
            row=8,
            column=0,
            columnspan=3,
            sticky="ew",
            ipady=5,
        )

        footer = ttk.Frame(card, style="Surface.TFrame")
        footer.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        footer.columnconfigure(0, weight=1)

        self.open_folder_button = ttk.Button(
            footer,
            text="Open Export Folder",
            style="Secondary.TButton",
            command=self._open_export_directory,
            state="disabled",
        )
        self.open_folder_button.grid(row=0, column=0, sticky="w")

        ttk.Button(
            footer,
            text="Export Files",
            style="Primary.TButton",
            command=self._export_files,
        ).grid(row=0, column=1, sticky="e")

        self._update_export_controls()
        self.status_var.set("Choose the export destination and file types.")

    def _choose_export_directory(self) -> None:
        """Let the user choose a destination folder."""

        current_value = self.export_directory_var.get().strip()
        initial_directory = Path(current_value).expanduser()

        if not initial_directory.exists():
            initial_directory = self.loaded_dataset.source_path.parent

        selected_directory = filedialog.askdirectory(
            parent=self.root,
            title="Choose Export Folder",
            initialdir=initial_directory,
        )

        if selected_directory:
            self.export_directory_var.set(selected_directory)

    def _update_export_controls(self) -> None:
        """Enable filename fields only for selected deliverables."""

        self.excel_filename_entry.configure(
            state="normal" if self.export_excel_var.get() else "disabled"
        )
        self.csv_filename_entry.configure(
            state="normal" if self.export_csv_var.get() else "disabled"
        )

    @staticmethod
    def _validated_export_filename(
        filename: str,
        extension: str,
    ) -> str:
        """Validate one filename and append its extension when omitted."""

        cleaned_name = filename.strip()

        if not cleaned_name:
            raise ValueError("Export filenames cannot be empty.")

        if Path(cleaned_name).name != cleaned_name:
            raise ValueError(
                "Enter filenames only; choose the folder separately."
            )

        suffix = Path(cleaned_name).suffix.lower()

        if not suffix:
            cleaned_name += extension
        elif suffix != extension:
            raise ValueError(
                f"The filename '{cleaned_name}' must use "
                f"the {extension} extension."
            )

        return cleaned_name

    def _export_files(self) -> None:
        """Export selected deliverables with overwrite confirmation."""

        if (
            self.cleaning_result is None
            or self.profile is None
            or self.after_profile is None
        ):
            return

        export_excel = self.export_excel_var.get()
        export_csv = self.export_csv_var.get()

        if not export_excel and not export_csv:
            messagebox.showwarning(
                "Nothing Selected",
                "Select at least one file type to export.",
                parent=self.root,
            )
            return

        try:
            directory_text = self.export_directory_var.get().strip()
            if not directory_text:
                raise ValueError("Choose an export destination folder.")

            output_directory = Path(directory_text).expanduser()
            output_paths: list[Path] = []
            excel_path: Path | None = None
            csv_path: Path | None = None

            if export_excel:
                excel_name = self._validated_export_filename(
                    self.excel_filename_var.get(),
                    ".xlsx",
                )
                excel_path = output_directory / excel_name
                output_paths.append(excel_path)

            if export_csv:
                csv_name = self._validated_export_filename(
                    self.csv_filename_var.get(),
                    ".csv",
                )
                csv_path = output_directory / csv_name
                output_paths.append(csv_path)
        except (OSError, ValueError) as exc:
            messagebox.showerror(
                "Invalid Export Settings",
                str(exc),
                parent=self.root,
            )
            return

        existing_paths = [path for path in output_paths if path.exists()]
        if existing_paths:
            existing_names = "\n".join(
                f"• {path.name}" for path in existing_paths
            )
            should_overwrite = messagebox.askyesno(
                "Replace Existing Files?",
                (
                    "The following files already exist:\n\n"
                    f"{existing_names}\n\n"
                    "Replace them?"
                ),
                parent=self.root,
            )
            if not should_overwrite:
                self.status_var.set("Export cancelled — no files changed.")
                return

        self.status_var.set("Generating export files...")
        self.root.configure(cursor="watch")
        self.root.update_idletasks()

        try:
            exported_paths: list[Path] = []
            if excel_path is not None:
                exported_paths.append(
                    export_excel_report(
                        excel_path,
                        self.cleaning_result,
                        self.profile,
                        self.after_profile,
                    )
                )
            if csv_path is not None:
                exported_paths.append(
                    export_cleaned_csv(
                        csv_path,
                        self.cleaning_result,
                    )
                )
        except Exception as exc:
            self.status_var.set("Export failed.")
            messagebox.showerror(
                "Export Error",
                (
                    "The selected files could not be exported.\n\n"
                    f"Details: {exc}"
                ),
                parent=self.root,
            )
            return
        finally:
            self.root.configure(cursor="")

        self.last_export_directory = output_directory.resolve()
        self.open_folder_button.configure(state="normal")
        exported_names = "\n".join(
            f"• {path.name}" for path in exported_paths
        )
        self.status_var.set(
            f"Export complete — {len(exported_paths):,} "
            f"{'file' if len(exported_paths) == 1 else 'files'} created."
        )
        messagebox.showinfo(
            "Export Complete",
            (
                "The following files were created successfully:\n\n"
                f"{exported_names}\n\n"
                f"Saved in:\n{self.last_export_directory}"
            ),
            parent=self.root,
        )

    def _open_export_directory(self) -> None:
        """Open the most recent export folder in the system file manager."""

        if self.last_export_directory is None:
            return

        try:
            if os.name == "nt":
                os.startfile(self.last_export_directory)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(self.last_export_directory)])
            else:
                subprocess.Popen(
                    ["xdg-open", str(self.last_export_directory)]
                )
        except OSError as exc:
            messagebox.showerror(
                "Unable to Open Folder",
                str(exc),
                parent=self.root,
            )

    def _build_status_bar(
        self,
        parent: ttk.Frame,
    ) -> None:
        """Create the application status bar."""

        ttk.Label(
            parent,
            textvariable=self.status_var,
            style="Status.TLabel",
            anchor="w",
        ).grid(
            row=3,
            column=0,
            sticky="ew",
            pady=(14, 0),
        )


def run_app() -> None:
    """Create and run the desktop application."""

    root = ttk.Window(themename="litera")
    CleanerApplication(root)
    root.mainloop()
