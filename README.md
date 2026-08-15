# Gantt Template

Generates `Gantt_Template.xlsx` — a self-contained Excel planning template with
live formulas. No Python at runtime: edit an input cell and the Gantt views
recalculate in Excel.

## Build

macOS / Linux:

```sh
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python build_template.py            # -> Gantt_Template.xlsx
```

Windows (PowerShell), Python 3.10+:

```powershell
py -3.12 -m venv .venv
.venv\Scripts\pip install -r requirements.txt
$env:PYTHONUTF8 = "1"        # the workbook uses non-ASCII; see below
.venv\Scripts\python build_template.py
```

`PYTHONUTF8=1` matters on Windows. The check columns use a warning sign and the
notes use em dashes, and those characters reach the console through pytest
failure messages. A console at the default code page raises `UnicodeEncodeError`
when it tries to print them.

## Test

Structural tests run anywhere. Numerical tests recalculate the workbook and
compare the results against `gantt/reference.py`, an independent Python model of
the scheduling algorithm; they skip if no engine is available.

```sh
.venv/bin/python -m pytest tests/ -q
python -c "from gantt import recalc; print(recalc.describe())"   # which engine
```

### Choosing the engine

| Engine | Platform | Notes |
|---|---|---|
| Excel over COM | Windows | Preferred. Needs `pywin32`. |
| LibreOffice | any | `brew install --cask libreoffice`, or install and let it be found |

**Excel is the more trustworthy of the two, and not by a small margin.** Every
rendering bug this project has shipped was invisible to LibreOffice: a
conditional fill keyed on `fgColor` rather than `bgColor`, a colour padded to a
transparent alpha, and `"" > 0` evaluating true so every empty cell shaded. The
numbers were right in all three cases. On Windows, prefer Excel.

Overrides:

- `GANTT_RECALC` — `excel`, `libreoffice` or `none` to force the choice.
- `GANTT_SOFFICE` — an explicit path to the `soffice` binary.

### Proving which engine actually ran

Selecting a backend is not the same as it working. This builds a throwaway
workbook, recalculates it, and reports both the value that came back and the
application that produced it — every engine stamps its name into
`docProps/app.xml`, so the answer cannot be faked:

```sh
python -m gantt.recalc
```

```
selected     : excel via COM
backend      : excel
computed 1+1 : 2  (expected 2)
written by   : Microsoft Excel

OK: the selected engine really recalculated the workbook.
```

A computed `2` proves something evaluated the formula; the stamp proves what.
Exit status is non-zero if the engine that ran is not the one selected, so this
is safe to put in a setup script.

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

## Getting the data out

```sh
python -m gantt.exchange export plan.xlsx -o plan.json
python -m gantt.exchange export plan.xlsx            # stdout
```

Covers every input tab: `Config`, `Assignees`, `Capacity`, `Equipment`,
`Holidays`, `Tasks`, `Sub-Tasks`. Derived columns are deliberately absent —
they are recomputed by the workbook and would be stale the moment anything
upstream changed.

Nothing computed is ever read, so **export needs no Excel or LibreOffice** and
works on a workbook that has never been opened. A sub-task's ID is rebuilt in
Python using the same rule the formula uses, rather than read from the cell.

Week grids are exported by offset from `config.start_week`, not by calendar
week, since calendar labels live in computed cells:

```json
{"assignee": "Alice", "days_by_week_offset": [5, 5, 4.5, 5, ...]}
```

## Getting the data back in

```sh
python -m gantt.exchange import plan.xlsx --json update.json --dry-run -v
python -m gantt.exchange import plan.xlsx --json update.json
python -m gantt.exchange rebuild plan.json -o new.xlsx
```

Two modes, because a first ingest and a weekly update want opposite things.

**`merge`** (the default) respects field ownership. The JSON owns facts observed
in the outside world; the workbook owns planning decisions:

| Owned by the JSON | Owned by the workbook |
|---|---|
| task name, category | priority, complexity, equipment |
| sub-task name | default assignee, earliest start week |
| the existence of new work | capacity, equipment availability, holidays, config |

Without that rule every weekly import would silently revert whatever you had
changed since the last one. Ownership only governs *updates* — a row that does
not exist yet is written in full.

**`replace`** treats the JSON as the whole truth, which is what a first ingest
needs. It writes rows from the top and clears whatever the JSON does not
mention, so no demo data is left mixed in. `rebuild` generates a fresh workbook and replaces into it, and is the
version-upgrade path: export from the old template, rebuild with the new, and
fields the old export never knew about take the generator's defaults.

Two details worth knowing:

- **Sub-tasks match on (parent, name), not on their ID.** Sub-task IDs are
  positional — `T-01.01` is simply the first `T-01` row from the top — so
  inserting a row renames every one below it and an update would attach to the
  wrong work.
- **A new assignee or equipment type is added even in merge mode.** Referential
  integrity beats ownership: a task naming an assignee the workbook has never
  heard of fails its dropdown and schedules nothing.
- **`rebuild` sizes the workbook to the data**, with headroom, so a plan larger
  than the shipped 30 tasks / 600 sub-tasks needs no source edit. The dimensions
  are recorded inside the file, so a later export reads it at its real size
  rather than assuming the defaults and truncating it.

## Layout

The workbook opens on a **Guide** tab explaining the fill order, how scheduling
works, what Rank is, and what each Check message means. It lives inside the file
so it travels with a copy.

| Tab | Role |
|---|---|
| `Guide` | How to use the file |
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
