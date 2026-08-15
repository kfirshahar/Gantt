"""Recalculating a generated workbook so its formula results can be checked.

openpyxl writes formulas but never evaluates them, so verifying the numbers
needs a real spreadsheet engine. Two are supported, and which one runs matters
more than it might appear.

LibreOffice is available everywhere and recalculates numbers faithfully, but it
does not reproduce Excel's formatting or comparison semantics. Every rendering
bug this project has shipped — a differential fill keyed on fgColor rather than
bgColor, a colour padded to a transparent alpha, and `"" > 0` evaluating true —
was invisible to it. Excel, driven over COM, is the engine the file is actually
opened in, so on Windows it is the more trustworthy of the two and is preferred.

Selection, in order:

- ``GANTT_RECALC`` — ``excel``, ``libreoffice`` or ``none`` to force a choice.
- otherwise Excel if this is Windows and pywin32 is importable,
- otherwise LibreOffice if it can be found,
- otherwise nothing, and callers skip their numeric checks.

``GANTT_SOFFICE`` overrides the search for the LibreOffice binary.
"""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
from pathlib import Path

ENV_BACKEND = "GANTT_RECALC"
ENV_SOFFICE = "GANTT_SOFFICE"

_SOFFICE_CANDIDATES = {
    "Darwin": [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ],
    "Windows": [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ],
    "Linux": [
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/snap/bin/libreoffice",
    ],
}


def find_libreoffice() -> Path | None:
    """The soffice binary, or None. An explicit override wins outright."""
    explicit = os.environ.get(ENV_SOFFICE)
    if explicit:
        path = Path(explicit)
        return path if path.exists() else None

    for candidate in _SOFFICE_CANDIDATES.get(platform.system(), []):
        if Path(candidate).exists():
            return Path(candidate)

    for name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return Path(found)
    return None


def excel_available() -> bool:
    """Whether Excel can be driven over COM from this interpreter."""
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


class Backend:
    """Recalculates a workbook and returns a path to the recalculated copy."""

    name = "none"

    def recalculate(self, src: Path, outdir: Path) -> Path:
        raise NotImplementedError


class LibreOffice(Backend):
    name = "libreoffice"

    def __init__(self, soffice: Path) -> None:
        self.soffice = soffice

    def recalculate(self, src: Path, outdir: Path) -> Path:
        outdir.mkdir(parents=True, exist_ok=True)
        subprocess.run(
            [str(self.soffice), "--headless", "--norestore",
             "--convert-to", "xlsx", "--outdir", str(outdir), str(src)],
            check=True, capture_output=True, timeout=300)
        return outdir / src.name


class ExcelCOM(Backend):
    """Drives the installed Excel. Windows only.

    NOT EXERCISED ON macOS — written against the COM API but verified only on
    the target machine. `DispatchEx` deliberately starts a private instance so
    the run cannot disturb, or be disturbed by, a copy the user already has open.
    """

    name = "excel"

    def recalculate(self, src: Path, outdir: Path) -> Path:
        import win32com.client

        outdir.mkdir(parents=True, exist_ok=True)
        dst = outdir / src.name
        xlOpenXMLWorkbook = 51

        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        try:
            book = app.Workbooks.Open(str(src.resolve()))
            app.CalculateFullRebuild()
            book.SaveAs(str(dst.resolve()), FileFormat=xlOpenXMLWorkbook)
            book.Close(SaveChanges=False)
        finally:
            app.Quit()
        return dst


def get_backend() -> Backend | None:
    """The backend to use, honouring GANTT_RECALC, or None if none is usable."""
    choice = os.environ.get(ENV_BACKEND, "auto").strip().lower()

    if choice == "none":
        return None

    if choice in ("excel", "auto") and excel_available():
        return ExcelCOM()
    if choice == "excel":
        return None

    if choice in ("libreoffice", "auto"):
        soffice = find_libreoffice()
        if soffice:
            return LibreOffice(soffice)
    return None


def describe() -> str:
    backend = get_backend()
    if backend is None:
        return ("no recalculation backend found; set GANTT_SOFFICE to a soffice "
                "binary, or install pywin32 on Windows to drive Excel")
    if isinstance(backend, LibreOffice):
        return f"libreoffice at {backend.soffice}"
    return "excel via COM"
