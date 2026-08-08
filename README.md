# Gantt Template

Generates `Gantt_Template.xlsx` — a self-contained Excel planning template with
live formulas. No Python at runtime: edit an input cell and the Gantt views
recalculate in Excel.

## Build

```sh
python3 -m venv .venv && .venv/bin/pip install openpyxl
.venv/bin/python build_template.py            # -> Gantt_Template.xlsx
```

## Test

Numerical tests recalculate the workbook with LibreOffice headless and compare
the results against `gantt/reference.py`, an independent Python model of the
scheduling algorithm. They skip automatically if LibreOffice is absent.

```sh
brew install --cask libreoffice     # optional, enables the numerical half
.venv/bin/pip install pytest
.venv/bin/python -m pytest tests/ -q
```

## How it works

Effort is derived, never typed: `base_days(complexity) ÷ proficiency(assignee)`
per sub-task, summed up to the parent task. Sub-tasks are ranked by priority,
then start week, then row, and then packed into weeks — each one taking whatever
days its assignee has left after higher-ranked work, spilling into later weeks
until its effort is exhausted. Durations fall out of capacity rather than being
entered.

`Gantt-Deep` runs the identical algorithm at day granularity over a windowed
range of weeks, sharing the rank order so both views always agree.

Equipment is validated, not scheduled: a shortage is flagged, it does not
silently stretch the plan.

See `docs/superpowers/specs/` for the full design and the decisions behind it.

## Layout

| Tab | Role |
|---|---|
| `Gantt-High` | Assignee load, task timeline, equipment, checks — read-only |
| `Gantt-Deep` | Sub-task × day view, windowed — read-only but for the two window cells |
| `Tasks`, `Sub-Tasks` | What needs doing |
| `Assignees`, `Capacity` | Who, and how many days each week |
| `Equipment`, `Holidays` | Shared equipment pool, company-wide off days |
| `Config` | Complexity → base days, priorities, horizon |

Grey italic columns are computed — typing in them breaks the schedule.
`CalcWeek` and `CalcDay` are hidden: right-click a sheet tab → *Unhide*.

## Changing the horizon

26 week columns are built into the file; `Horizon (weeks)` on `Config` sets how
many are active. Weeks past it are switched off and greyed out, so going from 12
to 16 weeks is one cell edit — no regeneration. Beyond 26, raise `WEEK_COLS` in
`gantt/layout.py` and rebuild.

## Conventions

Work week is Sunday–Thursday. Week 1 is the Sunday-start week containing
1 January, so a week number maps to exactly one work week.
