"""Shared fixtures.

Workbook dimensions live in module-level globals on `gantt.layout`, because
formulas and defined names are generated from them and every sheet builder has
to agree. Resizing is therefore process-wide, and a test that builds a larger
workbook would silently change the shape of every test that follows it.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from gantt import layout as L, names as N  # noqa: E402


@pytest.fixture(autouse=True)
def restore_layout_sizes():
    """Put the layout back however a test left it."""
    saved = {name: getattr(L, name) for name in L.SIZES}
    yield
    if any(getattr(L, name) != value for name, value in saved.items()):
        L.configure(**saved)
        N.refresh()
