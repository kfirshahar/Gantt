"""Moving the input data in and out of a workbook as JSON.

Phase 1 of the roadmap: export only.

One invariant shapes everything here — **nothing computed is ever read**. Every
value comes from a cell the user types into, which holds a literal. That matters
because openpyxl does not evaluate formulas: a workbook generated a moment ago
has no cached results at all, so an exporter that reached for a derived column
would return `None` for a freshly built file and a stale number for an edited
one. Derived facts the JSON needs, such as a sub-task's ID, are recomputed here
using the same rule the workbook uses.

The second consequence is that export needs no recalculation engine, so it runs
anywhere.

Week-indexed grids are exported by **offset from `config.start_week`**, not by
calendar week. Calendar labels live in computed cells, and a positional list
plus the start week says the same thing without depending on them.
"""

from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from . import layout as L, names as N

SCHEMA_VERSION = 1


def _clean(value: Any) -> Any:
    """A cell's value, with blanks normalised to None."""
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        return stripped or None
    return value


def _row_is_empty(ws, row: int, columns) -> bool:
    return all(_clean(ws.cell(row=row, column=c).value) is None for c in columns)


def _as_iso(value: Any) -> str | None:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, str):
        return value.strip() or None
    return None


def _config(wb) -> dict:
    ws = wb[L.CONFIG]
    complexity = []
    for i in range(L.CFG_COMPLEXITY_COUNT):
        row = L.CFG_COMPLEXITY_FIRST + i
        name = _clean(ws.cell(row=row, column=1).value)
        if name is not None:
            complexity.append({"name": name,
                               "base_days": _clean(ws.cell(row=row, column=2).value)})

    priorities = []
    for i in range(L.CFG_PRIORITY_COUNT):
        row = L.CFG_PRIORITY_FIRST + i
        name = _clean(ws.cell(row=row, column=1).value)
        if name is not None:
            priorities.append({"name": name,
                               "order": _clean(ws.cell(row=row, column=2).value)})

    return {
        "year": _clean(ws.cell(row=L.CFG_YEAR_ROW, column=2).value),
        "start_week": _clean(ws.cell(row=L.CFG_START_WEEK_ROW, column=2).value),
        "horizon_weeks": _clean(ws.cell(row=L.CFG_HORIZON_ROW, column=2).value),
        "complexity": complexity,
        "priorities": priorities,
    }


def _assignees(wb) -> list[dict]:
    ws = wb[L.ASSIGNEES]
    out = []
    for row in range(L.GRID_FIRST_DATA_ROW, N.LAST_ASG_ROW + 1):
        name = _clean(ws.cell(row=row, column=1).value)
        if name is None:
            continue
        out.append({"name": name,
                    "proficiency": _clean(ws.cell(row=row, column=2).value),
                    "notes": _clean(ws.cell(row=row, column=3).value)})
    return out


def _week_grid(wb, sheet: str, labels: list[str], key: str, values_key: str) -> list[dict]:
    """A per-week grid, exported by offset from the project start week.

    `labels` supplies the row identities: the Capacity sheet's own name column is
    a formula mirroring Assignees, so it cannot be read, and the row order is
    what ties the two together.
    """
    ws = wb[sheet]
    out = []
    for i, label in enumerate(labels):
        row = L.GRID_FIRST_DATA_ROW + i
        series = [_clean(ws.cell(row=row, column=L.GRID_FIRST_WEEK_COL + w).value)
                  for w in range(L.WEEK_COLS)]
        out.append({key: label, values_key: series})
    return out


def _equipment_names(wb) -> list[str]:
    ws = wb[L.EQUIPMENT]
    names = []
    for row in range(L.GRID_FIRST_DATA_ROW, N.LAST_EQP_ROW + 1):
        name = _clean(ws.cell(row=row, column=1).value)
        if name is not None:
            names.append(name)
    return names


def _holidays(wb) -> list[dict]:
    ws = wb[L.HOLIDAYS]
    out = []
    for row in range(L.GRID_FIRST_DATA_ROW, N.LAST_HOL_ROW + 1):
        when = _as_iso(ws.cell(row=row, column=1).value)
        if when is None:
            continue
        out.append({"date": when, "name": _clean(ws.cell(row=row, column=2).value)})
    return out


def _tasks(wb) -> list[dict]:
    ws = wb[L.TASKS]
    out = []
    for row in range(L.TASK_FIRST_ROW, N.LAST_TASK_ROW + 1):
        if _row_is_empty(ws, row, L.TASK_INPUT_COLS):
            continue
        out.append({
            "id": _clean(ws.cell(row=row, column=L.T_ID).value),
            "name": _clean(ws.cell(row=row, column=L.T_NAME).value),
            "category": _clean(ws.cell(row=row, column=L.T_CATEGORY).value),
            "priority": _clean(ws.cell(row=row, column=L.T_PRIORITY).value),
            "complexity": _clean(ws.cell(row=row, column=L.T_COMPLEXITY).value),
            "equipment": _clean(ws.cell(row=row, column=L.T_EQUIPMENT).value),
            "default_assignee": _clean(ws.cell(row=row, column=L.T_DEF_ASSIGNEE).value),
            "earliest_start_week": _clean(ws.cell(row=row, column=L.T_START_WW).value),
        })
    return out


def _sub_tasks(wb) -> list[dict]:
    """Sub-tasks, with the ID recomputed rather than read.

    The workbook builds the ID with a formula, so the cell holds no value until
    something recalculates. The rule is simply the parent plus a running count
    of that parent from the top, which is cheap to reproduce — and reproducing it
    keeps export working on a file that has never been opened in Excel.
    """
    ws = wb[L.SUBTASKS]
    seen: dict[str, int] = {}
    out = []
    for row in range(L.SUB_FIRST_ROW, N.LAST_SUB_ROW + 1):
        if _row_is_empty(ws, row, L.SUB_INPUT_COLS):
            continue
        parent = _clean(ws.cell(row=row, column=L.S_PARENT).value)
        index = None
        if parent is not None:
            seen[parent] = seen.get(parent, 0) + 1
            index = seen[parent]
        out.append({
            "id": f"{parent}.{index:02d}" if parent is not None else None,
            "parent": parent,
            "name": _clean(ws.cell(row=row, column=L.S_NAME).value),
            "complexity": _clean(ws.cell(row=row, column=L.S_COMPLEXITY).value),
            "assignee": _clean(ws.cell(row=row, column=L.S_ASSIGNEE).value),
        })
    return out


def export(path: str | Path) -> dict:
    """Every input tab of a workbook, as a plain dictionary."""
    wb = load_workbook(Path(path), data_only=False)
    assignees = _assignees(wb)

    return {
        "schema_version": SCHEMA_VERSION,
        "source": Path(path).name,
        "config": _config(wb),
        "assignees": assignees,
        "capacity": _week_grid(wb, L.CAPACITY, [a["name"] for a in assignees],
                               "assignee", "days_by_week_offset"),
        "equipment": _week_grid(wb, L.EQUIPMENT, _equipment_names(wb),
                                "type", "units_by_week_offset"),
        "holidays": _holidays(wb),
        "tasks": _tasks(wb),
        "sub_tasks": _sub_tasks(wb),
    }


def to_json(data: dict) -> str:
    return json.dumps(data, indent=2, ensure_ascii=False, sort_keys=False)


def _main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(
        prog="python -m gantt.exchange",
        description="Move a Gantt workbook's input data in and out as JSON.")
    sub = parser.add_subparsers(dest="command", required=True)

    exporter = sub.add_parser("export", help="write a workbook's input tabs as JSON")
    exporter.add_argument("workbook", type=Path)
    exporter.add_argument("-o", "--output", type=Path,
                          help="file to write; stdout if omitted")

    args = parser.parse_args(argv)
    if args.command == "export":
        text = to_json(export(args.workbook))
        if args.output:
            args.output.write_text(text + "\n", encoding="utf-8")
            print(f"wrote {args.output}")
        else:
            print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
