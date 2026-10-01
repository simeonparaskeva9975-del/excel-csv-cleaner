"""Visual constants and ttk styles for the desktop application."""

from __future__ import annotations

import ttkbootstrap as ttk


BACKGROUND = "#F4F6F8"
SURFACE = "#FFFFFF"
BORDER = "#DDE2E8"
TEXT_PRIMARY = "#182230"
TEXT_SECONDARY = "#647184"
ACCENT = "#25476A"
ACCENT_HOVER = "#1C3854"
DISABLED = "#A8B1BC"
SUCCESS = "#287A5B"
WARNING = "#A66A18"
ERROR = "#B54747"

FONT_FAMILY = "Segoe UI"
FONT_TITLE = (FONT_FAMILY, 20, "bold")
FONT_HEADING = (FONT_FAMILY, 13, "bold")
FONT_BODY = (FONT_FAMILY, 10)
FONT_SMALL = (FONT_FAMILY, 9)
FONT_BUTTON = (FONT_FAMILY, 10, "bold")


def apply_application_styles(style: ttk.Style) -> None:
    """Register the custom styles used by the application."""

    style.configure("App.TFrame", background=BACKGROUND)
    style.configure("Surface.TFrame", background=SURFACE)
    style.configure(
        "Card.TFrame",
        background=SURFACE,
        bordercolor=BORDER,
        borderwidth=1,
        relief="solid",
    )

    style.configure(
        "Title.TLabel",
        background=BACKGROUND,
        foreground=TEXT_PRIMARY,
        font=FONT_TITLE,
    )
    style.configure(
        "Subtitle.TLabel",
        background=BACKGROUND,
        foreground=TEXT_SECONDARY,
        font=FONT_BODY,
    )
    style.configure(
        "Heading.TLabel",
        background=SURFACE,
        foreground=TEXT_PRIMARY,
        font=FONT_HEADING,
    )
    style.configure(
        "Body.TLabel",
        background=SURFACE,
        foreground=TEXT_SECONDARY,
        font=FONT_BODY,
    )
    style.configure(
        "Small.TLabel",
        background=SURFACE,
        foreground=TEXT_SECONDARY,
        font=FONT_SMALL,
    )
    style.configure(
        "OptionTitle.TLabel",
        background=SURFACE,
        foreground=TEXT_PRIMARY,
        font=FONT_BUTTON,
    )
    style.configure(
        "Notice.TLabel",
        background="#FFF7E8",
        foreground="#78511B",
        font=FONT_SMALL,
        padding=(10, 8),
    )
    style.configure(
        "Option.TCheckbutton",
        background=SURFACE,
        foreground=TEXT_PRIMARY,
        font=FONT_BODY,
        padding=(2, 2),
    )
    style.map(
        "Option.TCheckbutton",
        background=[("active", SURFACE)],
        foreground=[("disabled", DISABLED)],
    )
    style.configure(
        "Version.TLabel",
        background=BACKGROUND,
        foreground=TEXT_SECONDARY,
        font=FONT_SMALL,
    )
    style.configure(
        "FileName.TLabel",
        background=SURFACE,
        foreground=TEXT_PRIMARY,
        font=(FONT_FAMILY, 12, "bold"),
    )
    style.configure(
        "FileMeta.TLabel",
        background=SURFACE,
        foreground=TEXT_SECONDARY,
        font=FONT_SMALL,
    )
    style.configure(
        "MetricValue.TLabel",
        background=SURFACE,
        foreground=TEXT_PRIMARY,
        font=(FONT_FAMILY, 17, "bold"),
    )
    style.configure(
        "MetricName.TLabel",
        background=SURFACE,
        foreground=TEXT_SECONDARY,
        font=FONT_SMALL,
    )

    style.configure(
        "StepActive.TLabel",
        background=ACCENT,
        foreground="#FFFFFF",
        font=FONT_BUTTON,
        padding=(10, 6),
        anchor="center",
    )
    style.configure(
        "StepInactive.TLabel",
        background="#E8ECF1",
        foreground=TEXT_SECONDARY,
        font=FONT_BUTTON,
        padding=(10, 6),
        anchor="center",
    )
    style.configure(
        "StepTextActive.TLabel",
        background=BACKGROUND,
        foreground=TEXT_PRIMARY,
        font=FONT_BUTTON,
    )
    style.configure(
        "StepTextInactive.TLabel",
        background=BACKGROUND,
        foreground=TEXT_SECONDARY,
        font=FONT_BODY,
    )

    style.configure(
        "Primary.TButton",
        background=ACCENT,
        foreground="#FFFFFF",
        font=FONT_BUTTON,
        padding=(18, 10),
        borderwidth=0,
    )
    style.map(
        "Primary.TButton",
        background=[
            ("disabled", DISABLED),
            ("pressed", ACCENT_HOVER),
            ("active", ACCENT_HOVER),
        ],
        foreground=[
            ("disabled", "#EEF1F4"),
            ("pressed", "#FFFFFF"),
            ("active", "#FFFFFF"),
        ],
    )

    style.configure(
        "Secondary.TButton",
        background="#E9EDF2",
        foreground=TEXT_PRIMARY,
        font=FONT_BUTTON,
        padding=(14, 9),
        borderwidth=0,
    )
    style.map(
        "Secondary.TButton",
        background=[
            ("pressed", "#D8DEE6"),
            ("active", "#D8DEE6"),
        ],
    )

    style.configure(
        "Data.Treeview",
        background=SURFACE,
        fieldbackground=SURFACE,
        foreground=TEXT_PRIMARY,
        bordercolor=BORDER,
        rowheight=29,
        font=FONT_SMALL,
    )
    style.configure(
        "Data.Treeview.Heading",
        background="#E9EDF2",
        foreground=TEXT_PRIMARY,
        font=FONT_BUTTON,
        padding=(9, 8),
        bordercolor=BORDER,
        borderwidth=1,
        relief="solid",
    )
    style.map(
        "Data.Treeview",
        background=[("selected", "#DDE9F3")],
        foreground=[("selected", TEXT_PRIMARY)],
    )

    style.configure(
        "Status.TLabel",
        background="#E9EDF2",
        foreground=TEXT_SECONDARY,
        font=FONT_SMALL,
        padding=(12, 8),
    )
