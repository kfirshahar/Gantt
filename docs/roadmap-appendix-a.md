# Roadmap — Appendix A and B

Date: 2026-08-15
Status: Phases 0, 1 and 2 complete; Phases 3-6 proposed

Decisions taken (Appendix A): `% done` is needed; `Done` work appears in the
Gantt where it actually happened; the field-ownership split is confirmed; real
data exists in a v4 workbook.

Appendix B adds the constraint that all of this must run on Windows with Python
3.12 and Office 2019, be driven by Opencode rather than Claude, and arrive there
through a GitHub remote. That reorders the plan again: portability now comes
first, because until the suite runs on the target machine nothing done after it
can be verified there.

## What the decisions imply

### Showing history means the horizon should stop moving

The original complaint was that a finished task reads "start outside horizon"
because its start week now sits before `Config`'s start week. The instinct is to
make the check tolerate that. The better fix is to stop moving the horizon.

Anchor it at the project start and add a separate **current week** pointer:

- Weeks before *current* hold **actuals** — what really happened.
- Weeks from *current* onward hold the **plan** — what the engine computes.
- Nothing falls outside the horizon merely because time has passed.

The false alarm becomes structurally impossible rather than suppressed, and
"outside horizon" recovers a real meaning: a start week beyond the horizon's
*end*, which genuinely cannot be scheduled.

### Actuals have to be recorded, because formulas cannot remember

A spreadsheet cannot observe that a sub-task became `In Progress` last Tuesday.
Without macros there is no state, so actual dates are **input**, not derived:
`Actual start WW` and `Actual end WW` on `Sub-Tasks`.

That is the deeper reason Appendix A's two items belong together. Recording
actuals by hand every week is tedious enough that the feature would go unused;
the agent ingest is what makes it viable.

### Live v4 data moves the exchange work early

Adding `Status`, `% done` and the actuals columns changes the layout of
`Sub-Tasks`. A populated v4 workbook cannot simply be reopened by the new
generator, and the tooling that rescues it is export/import. Proving the
round-trip on today's stable schema before the schema grows is the lower-risk
order anyway.

## Appendix B analysis: what is actually platform-dependent

Audited rather than assumed. The result is narrower than expected.

| Concern | Finding |
|---|---|
| Hardcoded paths | **One**: `SOFFICE` in `tests/test_template.py:21` |
| Path handling | `pathlib` throughout, no `os.path`, no separators assumed |
| Shell usage | No `shell=True`, no bash, no `os.system` |
| File encoding | No `open()` without explicit encoding anywhere |
| Python version | No 3.13-only syntax; the code is valid 3.12 |
| Generator | Filesystem-agnostic, no temp-dir assumptions |

So the reading in Appendix B is essentially right: **LibreOffice is the only
real platform dependency**, and it lives entirely in the test suite. The
generator itself already runs anywhere.

Three smaller things the audit did surface:

- **Console encoding.** The workbook uses `⚠` and `—`, which appear in test
  failure messages. A Windows console at cp1252 raises `UnicodeEncodeError` when
  pytest prints them. Fixed by `PYTHONUTF8=1`; no source change needed.
- **No dependency manifest.** There is no `requirements.txt` or `pyproject.toml`,
  so the target machine has nothing to install from.
- **No `.gitattributes`.** Line endings will churn across the two platforms.

### The recalculation backend is worth more than a path fix

The obvious fix is to discover LibreOffice instead of hardcoding it. Worth doing,
but it misses something this project has already been bitten by three times.

Every formatting bug that reached you — `fgColor` instead of `bgColor`, the
transparent alpha, and `"" > 0` being true — was **invisible to LibreOffice**.
It recalculates numbers faithfully and does not reproduce Excel's formatting or
comparison semantics. The numeric tests passed throughout while Office 2019
showed a blank grid.

The target machine has Office 2019. So the backend should be pluggable:

- **LibreOffice** — both platforms, found via `GANTT_SOFFICE`, then the usual
  install locations, then `PATH`.
- **Excel via COM** — Windows only, through `pywin32`. Higher fidelity: it is
  the actual engine the file will be opened in.
- **Neither** — numeric tests skip, as they already do.

On Windows the numeric tests then become genuinely trustworthy rather than
merely indicative. The structural tests that read `styles.xml` stay as they are;
they are platform-independent and are what caught the Excel-only bugs.

### Keeping the agent thin

Appendix B asks for Opencode on Windows with PowerShell instead of bash, and for
the skill to be specified here but implemented there.

The design principle that makes this nearly a non-issue: **put every piece of
logic in the Python CLI**. If the skill's whole job is to call
`python -m gantt.exchange import --json update.json plan.xlsx`, then it needs no
process handling, no text munging, no shell primitives — and bash versus
PowerShell stops mattering. A skill that shells out to `sed`, `jq` or `find` has
to be rewritten for Windows; one that calls a single Python entry point does not.

So the deliverable for Opencode is a specification plus a working CLI, not a
script. Per Appendix B, Claude specifies; the target agent implements.

## Phases

### Phase 0 — Portability and handover — **DONE**

Delivered: pluggable recalculation backend (`GANTT_RECALC`, `GANTT_SOFFICE`)
with Excel over COM preferred on Windows and LibreOffice as fallback;
`python -m gantt.recalc` proving which engine really ran; dependency manifests;
`.gitattributes`; a Windows README section; and the repository pushed to
`github.com/kfirshahar/Gantt`.

Verified on the target machine: Python 3.12, Office 2019, Excel COM confirmed by
the `docProps/app.xml` fingerprint reading "Microsoft Excel", 38 tests passing
with no warnings.

**It also found a bug that had nothing to do with portability.** Excel dropped 8
of the workbook's 37 conditional-format ranges the first time it saved the file,
because classic conditional formatting cannot reference another worksheet and
Excel promotes any rule that does into an x14 extension openpyxl neither reads
nor preserves. Nine rules reached into `CalcWeek`/`CalcDay`; each sheet now
mirrors the flag it needs into a hidden row of its own.

That is worth recording because it was **the risk this roadmap named for Phase
2, discovered before Phase 2 was written**. Had import been built first, it would
have silently stripped the colours from a user's workbook and the cause would
have been considerably harder to find.

### Known issue, deferred

A faulthandler "Windows fatal exception" still appears intermittently on the
target machine, in a different test each run. It does not fail the suite. The
varying location points at Excel process teardown timing rather than anything a
particular test does, so the next thing to try is reusing a single Excel instance
for the whole session rather than starting and quitting one per recalculation.
Deferred by agreement.

### Phase 1 — Export XLSX → JSON (schema v1) — **DONE**

- All input tabs: `Config`, `Assignees`, `Capacity`, `Equipment`, `Holidays`,
  `Tasks`, `Sub-Tasks`.
- `schema_version` at the root, `1` for today's layout.
- `python -m gantt.exchange export plan.xlsx -o plan.json`
- Must read the **v4 workbook as it exists on the target machine**, which is
  today's format, so this is exactly the file the exporter is written against.

Delivered with one invariant worth carrying into Phase 2: **nothing computed is
ever read**. openpyxl does not evaluate formulas, so a freshly generated
workbook has no cached results at all — an exporter reaching for a derived
column would return null for a new file and a stale number for an edited one.
Export therefore reads only cells the user types into, and rebuilds the few
derived facts it needs (a sub-task's ID) using the same rule the workbook uses.

Two consequences: export needs no recalculation engine, so it runs anywhere; and
it cannot silently emit stale data.

Size: S. Protects the live data before anything else moves.

### Phase 2 — Import JSON → XLSX, and rebuild (schema v1) — **DONE**

Verified feasible on both platforms. An openpyxl load/save round-trip preserves
all defined names, conditional-format ranges, validations, differential styles
and hidden sheets — and since Phase 0 removed the cross-sheet rules, that now
holds for a workbook Excel has saved too, which is the case that matters.
Import writes values into the user's own file, so their formatting and notes
survive.

- Match by ID; unknown IDs appended, known IDs updated in place.
- **Field ownership** (confirmed): JSON owns observed facts — status, % done,
  actual dates, newly discovered work. The workbook owns planning decisions —
  priority, earliest start week, assignee, capacity, equipment. Without this
  rule every weekly import silently reverts the planner's changes.
- `rebuild` mode generates a fresh workbook from JSON. **Export from the old
  version, rebuild with the new** is the upgrade path; unknown fields take
  defaults by construction.
- `--dry-run` prints the diff first.
- Test: export → import → export is byte-identical.

Size: M.

### Phase 3 — Status and remaining effort (schema v2)

- `Status` on `Sub-Tasks`: `TODO` / `In Progress` / `Done`, default `TODO`.
- `% done`, 0–100, meaningful while `In Progress`.
- Remaining effort = `effort × (1 − % done)`; `Done` is zero.
- Scheduling consumes **remaining** effort, so finished work stops competing for
  capacity and part-done work claims only its balance.
- Task status derived from its sub-tasks, for the same reason task effort is
  derived: two sources of truth for one fact will disagree.
- A v1 file imports with `status = TODO`, `% done = 0` — which is the first real
  exercise of the backward-compatibility rule.

Size: M.

### Phase 4 — Current week, actuals, and history in the Gantt

- `Current week` on `Config`, typed rather than `=TODAY()` so the plan does not
  shift underneath you between openings, with a helper cell showing today's week.
- `Actual start WW` / `Actual end WW` on `Sub-Tasks`.
- Planned work starts at `MAX(earliest start, current week)`, so nothing is
  scheduled into the past and a stale earliest-start means "as soon as possible".
- Consumed effort spreads evenly across the actual span — approximate, labelled
  as history, not worth modelling precisely.
- Actual and planned cells are told apart **by colour**, not by extra rows.
- Assignee load reads continuously across the current-week boundary: actuals
  behind it, plan ahead of it.

Size: L, and the phase that most increases what a reader must hold in their head.

### Phase 5 — Convergence diagnostics

- Quantified columns on `Tasks`: `Remaining`, `Scheduled`, `Shortfall`, and the
  week it finishes or "beyond horizon".
- The binding constraint per task: no capacity in range, horizon too short, or
  higher-priority work consuming the days.
- A **Plan health** block on `Gantt-High`: total shortfall, how many more weeks
  would absorb it, the bottleneck assignee, and the smallest single change that
  would make the plan converge.

Size: M, mostly formulas over grids that already exist.

### Phase 6 — Agent skill specification for Opencode

**Specified here, implemented there.** Not executed on this machine.

- A `SKILL.md` describing the weekly cadence: export current state, apply the
  status update, re-import, recompute, report what moved and what newly fails to
  converge.
- Every step is one `python -m gantt.exchange` call, so no PowerShell beyond invoking
  Python.
- The JSON payload shape, with worked examples for first ingest and for an
  update introducing new sub-tasks.
- Handover notes for the target agent's skill-creator: what the CLI guarantees,
  what it refuses, and what a dry-run diff looks like.

Size: M for the specification.

## Sequencing

```
0 Portability ──► 1 Export ──► 2 Import/rebuild ──► 3 Status ──► 4 Actuals ──► 5 Diagnostics ──► 6 Skill spec
    DONE            DONE            DONE             (schema v2)   (history)     (actionable)      (Opencode)
```

Phase 2 added one lesson worth carrying forward: **a value round-trip is not a
behaviour round-trip**. Export renders dates as ISO strings, and writing one
back verbatim left text in the Holidays cells. Excel compares text against a
date serial without ever matching, so every holiday was silently ignored and
every week looked like a full five days — while `export -> rebuild -> export`
remained a perfect fixed point, because export normalised both sides to strings.
Only recalculating the rebuilt file exposed it. There is now a test that
recalculates and compares the working-day row rather than the cell values.

Phase 0 first because the target machine cannot verify anything until the suite
runs there. Phases 1–2 are insurance for the live v4 data. Phases 3–5 deliver
"manage the project after it started". Phase 6 hands over.

A reasonable stopping point is the end of Phase 5: the workbook is fully usable
for a running project, with JSON exchange driven by hand rather than by an agent.

## Two risks worth naming

**Conceptual load.** The workbook is already at the edge of what one person can
hold in their head — that is why the Guide tab exists. Phase 4 adds a
current-week pointer, two date columns, a plan/actual distinction in every grid,
and a second colour language. Grow the Guide in the same commit as each phase,
never afterwards.

**Keep actuals display-only.** They should set remaining effort and nothing more.
If history feeds back into how future work is scheduled, the engine gains a
second code path and loses the "one engine, not two" property that currently
guarantees `Gantt-High` and `Gantt-Deep` agree. Everything Phase 4 needs can be
rendered from recorded inputs without the spill-over algorithm knowing that
history exists.

## Open question

Where should the split of work fall between this machine and the target? Phases
0–2 are the natural handover point: after them the live v4 data is safe and the
suite runs on Windows, so Phases 3–5 could be built here and pulled, or built
there against a verified baseline.
