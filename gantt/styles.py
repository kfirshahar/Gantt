"""Shared fills, fonts and helpers so every sheet reads as one document."""

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Palette
INK = "1F2933"
MUTED = "7B8794"
BAND_INPUT = "1F4E79"       # header band on tabs the user edits
BAND_OUTPUT = "3B5A40"      # header band on read-only views
BAND_SECTION = "E4E7EB"     # block sub-headings

FILL_INPUT_HDR = PatternFill("solid", fgColor=BAND_INPUT)
FILL_OUTPUT_HDR = PatternFill("solid", fgColor=BAND_OUTPUT)
FILL_SECTION = PatternFill("solid", fgColor=BAND_SECTION)
FILL_DERIVED = PatternFill("solid", fgColor="F5F7FA")
FILL_OVER = PatternFill("solid", fgColor="F8C9C9")
FILL_WARN = PatternFill("solid", fgColor="FDE9C9")
FILL_OK = PatternFill("solid", fgColor="D8EFD8")
FILL_BAR = PatternFill("solid", fgColor="9FC5E8")
FILL_WEEKEND = PatternFill("solid", fgColor="EEF1F4")

FONT_TITLE = Font(bold=True, size=14, color=INK)
FONT_HDR = Font(bold=True, size=10, color="FFFFFF")
FONT_SECTION = Font(bold=True, size=10, color=INK)
FONT_DERIVED = Font(italic=True, size=10, color=MUTED)
FONT_NOTE = Font(italic=True, size=9, color=MUTED)
FONT_BODY = Font(size=10, color=INK)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")

THIN = Side(style="thin", color="D2D6DC")
THICK = Side(style="medium", color=INK)
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_GROUP_TOP = Border(left=THIN, right=THIN, top=THICK, bottom=THIN)


def title(ws, text: str, note: str | None = None) -> None:
    ws["A1"] = text
    ws["A1"].font = FONT_TITLE
    if note:
        ws["A2"] = note
        ws["A2"].font = FONT_NOTE


def header_row(ws, row: int, first_col: int, values, output: bool = False) -> None:
    """Write a banded header row starting at `first_col`."""
    fill = FILL_OUTPUT_HDR if output else FILL_INPUT_HDR
    for offset, value in enumerate(values):
        cell = ws.cell(row=row, column=first_col + offset, value=value)
        cell.fill = fill
        cell.font = FONT_HDR
        cell.alignment = CENTER
        cell.border = BORDER_ALL


def section(ws, row: int, text: str, span: int = 12) -> None:
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = FONT_SECTION
    for c in range(1, span + 1):
        ws.cell(row=row, column=c).fill = FILL_SECTION


def widths(ws, mapping: dict) -> None:
    for letter, width in mapping.items():
        ws.column_dimensions[letter].width = width


def mark_derived(cell) -> None:
    cell.font = FONT_DERIVED
    cell.fill = FILL_DERIVED


def grey_inactive_weeks(ws, first_col: int, last_col: int,
                        first_row: int, last_row: int, header_row: int) -> None:
    """Dim week columns that fall outside the horizon set on Config.

    The columns are always present; this is what makes them read as switched
    off rather than as weeks with no work in them.
    """
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.utils import get_column_letter

    first = get_column_letter(first_col)
    ws.conditional_formatting.add(
        f"{first}{first_row}:{get_column_letter(last_col)}{last_row}",
        Rule(type="expression",
             dxf=DifferentialStyle(fill=FILL_WEEKEND, font=Font(color=MUTED)),
             formula=[f"{first}${header_row}>CfgStartWeek+CfgHorizon-1"]))
