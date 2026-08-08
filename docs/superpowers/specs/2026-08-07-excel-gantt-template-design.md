# Excel Gantt Template — Design

Date: 2026-08-07
Status: approved, first draft in progress

## Goal

A self-contained Excel template for project planning under three simultaneous
constraints — assignee capacity, equipment availability, and calendar time.
The user edits input tabs; the Gantt views recalculate live. No Python at
runtime, no database, no backend.

The template must be generic: duplicate it, replace the lookup values on
`Config`, and it fits a different project.

## Decisions

These were settled during brainstorming and are not open questions.

| Decision | Choice | Rationale |
|---|---|---|
| Scheduling | Manual assignment + validation | User picks owner and start week; the workbook flags over-allocation rather than reassigning work |
| Logic location | Live Excel formulas | Delivered `.xlsx` is portable and self-contained; editing an input recalculates immediately |
| Effort model | Derived from complexity | `base_days(complexity) ÷ proficiency(assignee)`, no hand-entered durations |
| Task effort | Sum of its sub-tasks | Guarantees the high-level and deep-dive views agree by construction |
| Equipment | Typed pools | Multiple types with per-week unit counts; each task requires one type |
| Excel target | 2016 / 2019 | No dynamic arrays — no `LET`, `FILTER`, `XLOOKUP`, `SEQUENCE` |
| Spread rule | Fill available days, spill over | Duration is computed from capacity, never typed |
| Contention order | Priority, then start week, then row | Higher priority claims capacity first; P1 is never starved by a P2 |
| Sub-tasks | Own tab, explicit rows | Each sub-task carries its own complexity and optional own assignee |
| Parent column | Repeats hidden, not merged | Looks merged; sorting, filtering and formulas all keep working |
| Work week | Sunday–Thursday | Five working days, Friday and Saturday are weekend |
| Week numbering | Sunday-start, week 1 contains Jan 1 | Puts the five working days contiguously at the front of each numbered week |
| Deep-dive scope | Day columns windowed | Start week + week count inputs, defaulting to project start |
| Calendar detail | Weeks tab + holiday list | Available days per assignee per week, plus company-wide off dates |

### Merged parent column

Excel stores a merged range's value only in its top-left cell, leaving the rest
empty, and refuses to sort or AutoFilter across merges. Since `Sub-Tasks` is a
table the user edits, the real parent ID is written into every row and
conditional formatting renders repeats in white-on-white with a heavy top border
per group. Visually identical to a merge; nothing breaks.

## Scale

- ~30 parent tasks, 1–20 sub-tasks each (~600 sub-tasks)
- 3–10 assignees
- 2–10 equipment types
- 12-week horizon (extendable)

Pre-built row counts: 30 tasks, 600 sub-tasks, 10 assignees, 10 equipment types,
50 holidays.

At this scale the heaviest grid is 600 × 12 ≈ 7,200 cells. Row-filtering the
deep-dive was designed and then dropped as unnecessary — day columns are
windowed, sub-task rows are not.

## Architecture

### Tabs

Inputs (user edits):

| Tab | Contents |
|---|---|
| `Config` | Project year, start week, horizon; complexity → base-days lookup; priority list |
| `Assignees` | Name, proficiency factor |
| `Capacity` | Assignee × week grid of available days |
| `Equipment` | Equipment type × week grid of available units |
| `Holidays` | Company-wide off dates |
| `Tasks` | ID, name, category, priority, complexity, equipment type, default assignee, earliest start week |
| `Sub-Tasks` | Parent ID, sub-task ID, name, complexity, assignee (blank inherits parent's default) |

Outputs (read-only, formula-driven):

| Tab | Contents |
|---|---|
| `Gantt-High` | Assignee rows × week columns showing load vs available; task timeline block; equipment block; checks block |
| `Gantt-Deep` | Sub-task rows × day columns, windowed to a start week and week count |

Hidden: `CalcWeek` and `CalcDay`, the scheduling grids.

### One engine, not two

The engine runs **once, at sub-task level**. The high-level view is a rollup of
sub-task allocations grouped by assignee; it is not a separate computation.
This is what allows a sub-task to carry its own assignee without the two views
disagreeing, and it removes a whole duplicate ranking model.

### Effort

```
sub_task_effort = base_days(sub_task.complexity) / proficiency(effective_assignee)
task_effort     = SUM(effort of its sub-tasks)
```

`effective_assignee` is the sub-task's own assignee, or the parent task's
default assignee when that cell is left blank. Leaving it blank is the common
case and keeps data entry cheap.

Task-level complexity remains as metadata for grouping and color. It does not
drive duration.

### Ranking

Each sub-task gets a unique integer rank from a composite sort key:

```
key  = parent_priority × 10^9 + parent_start_week × 10^6 + parent_row × 10^3 + subtask_index
rank = RANK(key, all_keys, ascending)
```

The `parent_row` and `subtask_index` terms guarantee uniqueness, so `RANK`
never ties. Everything recomputes when a priority changes.

### The spill-over grid

`CalcWeek` has one row per rank and one column per week. Row *k* pulls in
whichever sub-task currently holds rank *k* via `INDEX`/`MATCH`. Each cell is:

```
allocation = MIN( effort not yet consumed by this sub-task in earlier weeks,
                  assignee's available days this week
                    − days already claimed by higher-ranked sub-tasks )
```

bounded below at zero, and forced to zero before the parent's earliest start
week.

**Why the grid is ordered by rank.** The second term is a `SUMIF` over grid rows
*above* only. A cumulative that scanned the whole column would include the cell
itself and Excel would reject the workbook as a circular reference — and
referencing rows above *and* below is equally circular, because a row below
would in turn reference this one. Physically ordering the grid by rank is what
makes the cumulative one-directional and therefore legal. This is why `CalcWeek`
is a re-ordered projection of `Sub-Tasks` rather than a copy of it.

A sub-task's start and end weeks are the first and last non-zero cell in its
row. Duration is never typed.

### Deep-dive

`CalcDay` runs the identical algorithm at day granularity over the windowed
days, sharing the same rank order so the two views stay consistent.

Per-day capacity:

```
day_capacity = week_available_days / count of non-holiday Sun–Thu days in that week
```

and zero on a holiday.

Sub-tasks that began before the window are seeded with the effort they have left
at the window boundary:

```
remaining_at_window_start = total_effort − SUM(weekly allocations for weeks before the window)
```

so a window opened mid-project shows correct residual work.

### Equipment is validation, not scheduling

Per week, demand for an equipment type is the count of *parent tasks* with a
non-zero allocation that require it. The cell turns red where demand exceeds
supply on the `Equipment` tab. Consistent with the manual-plus-validation
model: a shortage warns, it does not silently stretch the plan.

## Validation and error handling

- Dropdowns on every cross-referencing column: assignee, equipment type,
  complexity, priority, parent ID. A typo cannot silently unschedule a task.
- A `Checks` block on `Gantt-High` reporting: sub-tasks whose parent ID does not
  exist, tasks with no sub-tasks, work that does not finish inside the horizon,
  assignees with zero capacity across the whole horizon, and equipment demand
  exceeding supply.

**There is deliberately no "over capacity" check.** Brainstorming assumed one,
and building it proved the assumption wrong: because the spill-over engine caps
each week's allocation at what the assignee actually has, allocated days can
never exceed available days. Any such check would report zero forever, and a
red "over-allocated" cell would never fire. Excess work does not show up as an
overloaded week — it shows up as work pushed past the end of the horizon.

The two surviving signals are therefore:

- **Work days that do not fit in the horizon** — the real error condition, and
  the number to act on.
- **Assignee-weeks with no slack left** — saturation, shown in red. Information,
  not an error: it marks who is the bottleneck holding the plan back.
- Every derived column is guarded against division by zero and empty input, so
  a freshly cleared template opens without a wall of `#DIV/0!`.

## Known limitations

1. **No task dependencies.** Explicitly out of scope per the original brief.
2. **Task ID lists in `Gantt-High` cells.** `TEXTJOIN` does not exist in Excel
   2016. The draft shows load-vs-available in the assignee cells and puts task
   identity in a task timeline block below. If the target is Excel 2019+, the
   in-cell ID list can be restored with `TEXTJOIN`.
3. **Priority changes reshuffle the grid** and can move unrelated tasks, because
   capacity is a shared pool. Correct, occasionally surprising.
4. **Widening the deep-dive window** past ~12 weeks will slow recalculation.
   The week-count input carries a soft cap and a note.

## Build approach

A Python generator, `build_template.py`, writes `Gantt_Template.xlsx` — headers,
named ranges, formula strings, conditional formats, data validation, and demo
data. The output has no runtime dependency; Python never runs again. The
template is versioned as reviewable code rather than as an undiffable binary.

Rejected: a Python model plus solver plus exporter. Since the logic must be
expressible in formulas anyway, a parallel Python solver would duplicate it and
immediately drift.

## Testing

- **Structural**: regenerate the workbook, read it back with `openpyxl`, assert
  formulas, named ranges, and validation rules are as specified.
- **Numerical**: LibreOffice headless recalculates the workbook; computed values
  are compared against a Python reference implementation of the spill-over
  algorithm run on the demo dataset. `openpyxl` cannot evaluate formulas, so
  this step is what catches a formula that is structurally perfect and
  numerically wrong.

## First draft dataset

Demo data shipped in the template, sized to make input ergonomics visible:

- 5 parent tasks with 3, 8, 1, 15 and 6 sub-tasks (33 total) — exercises the
  full 1–20 range
- 3 assignees at differing proficiency
- 2 equipment types
- 12-week horizon, WW33–WW44 2026 (Sun 2026-08-09 through Thu 2026-10-29)

The draft's purpose is to evaluate how easy the input tabs are to fill in.
