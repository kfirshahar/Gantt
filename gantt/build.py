"""Assemble the workbook."""

from pathlib import Path

from openpyxl import Workbook

from . import layout as L, names as N, sheet_calc, sheet_inputs, sheet_views, sheet_work

SHEET_ORDER = [
    L.GANTT_HIGH, L.GANTT_DEEP,
    L.TASKS, L.SUBTASKS, L.ASSIGNEES, L.CAPACITY, L.EQUIPMENT, L.HOLIDAYS,
    L.CONFIG, L.CALC_WEEK, L.CALC_DAY,
]

BUILDERS = {
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


def build(path: str | Path) -> Path:
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

    wb.active = wb.index(wb[L.GANTT_HIGH])
    path = Path(path)
    wb.save(path)
    return path
