"""Assemble the workbook."""

from pathlib import Path

from openpyxl import Workbook

from . import (layout as L, names as N, sheet_calc, sheet_guide, sheet_inputs,
               sheet_views, sheet_work)

SHEET_ORDER = [
    L.GUIDE,
    L.GANTT_HIGH, L.GANTT_DEEP,
    L.TASKS, L.SUBTASKS, L.ASSIGNEES, L.CAPACITY, L.EQUIPMENT, L.HOLIDAYS,
    L.CONFIG, L.CALC_WEEK, L.CALC_DAY,
]

BUILDERS = {
    L.GUIDE: sheet_guide.build_guide,
    L.CONFIG: sheet_inputs.build_config,
    L.ASSIGNEES: sheet_inputs.build_assignees,
    L.CAPACITY: sheet_inputs.build_capacity,
    L.EQUIPMENT: sheet_inputs.build_equipment,
    L.HOLIDAYS: sheet_inputs.build_holidays,
    L.TASKS: sheet_work.build_tasks,
    L.SUBTASKS: sheet_work.build_subtasks,
    L.CALC_WEEK: sheet_calc.build_calc_week,
    L.CALC_DAY: sheet_calc.build_calc_day,
    L.GANTT_HIGH: sheet_views.build_gantt_high,
    L.GANTT_DEEP: sheet_views.build_gantt_deep,
}


def build(path: str | Path, sizes: dict | None = None) -> Path:
    """Generate the workbook. `sizes` resizes it to fit a given dataset.

    Sizing has to happen before anything is written: block positions and the
    defined names are derived from the row counts, so changing them afterwards
    would leave formulas pointing at the wrong ranges.
    """
    if sizes:
        L.configure(**sizes)
        N.refresh()

    wb = Workbook()
    wb.remove(wb.active)
    for name in SHEET_ORDER:
        wb.create_sheet(name)

    # Defined names must exist before any formula that references them is
    # written, so that a round-trip through openpyxl resolves cleanly.
    N.register(wb)

    for name, builder in BUILDERS.items():
        builder(wb[name])

    for name in SHEET_ORDER:
        wb[name].sheet_view.showGridLines = False

    wb.active = wb.index(wb[L.GUIDE])
    path = Path(path)
    wb.save(path)
    return path
