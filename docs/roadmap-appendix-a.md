# Roadmap — Appendix A and B

Date: 2026-08-15
Status: Phases 0-5 complete; Phases 6-7 proposed

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

### Phase 3 — Status and remaining effort (schema v2) — **DONE**

- `Status` on `Sub-Tasks`: `TODO` / `In Progress` / `Done`, default `TODO`.
- `% done`, 0–100, meaningful while `In Progress`.
- Remaining effort = `effort × (1 − % done)`; `Done` is zero.
- Scheduling consumes **remaining** effort, so finished work stops competing for
  capacity and part-done work claims only its balance.
- Task status derived from its sub-tasks, for the same reason task effort is
  derived: two sources of truth for one fact will disagree.
- A v1 file imports with `status = TODO`, `% done = 0` — which is the first real
  exercise of the backward-compatibility rule.

Delivered as specified. Size: M.

**Found afterward, the hard way.** Migrating a real v5 workbook to v6 put
assignee names in `Status` and day counts in `% done`. Export read every file
by *today's* column positions, and Phase 3 had inserted the two new columns
exactly where a pre-v3 file kept its derived `Effective assignee` and `Effort`
— so the exporter collected formulas, or in a file Excel had saved, their
cached results, which look exactly like ordinary data. The Phase 1 invariant
("nothing computed is ever read") was right; the bug was assuming the column
map never moves. Fixed by stamping the schema version a workbook was built
with into the file itself (`TplSchemaVersion`) and having export read each
file by *its own* map, plus a hard refusal to import into a workbook older
than the tool. Worth carrying forward: **any column insertion is a migration
hazard**, not just a formula-writing exercise, the moment a populated
workbook exists on the other end.

### Phase 4 — Current week, actuals, and history in the Gantt — **DONE**

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

Delivered as specified, in two sub-commits (4a: the current-week anchor; 4b:
recording actuals and drawing history) plus two follow-on fixes. Size: L, and
the phase that most increases what a reader must hold in their head, as
predicted below.

Lessons worth carrying forward:

- **Finished work has to be checked first.** A Done task has no remaining
  effort, so every downstream check fired — "not scheduled" for work that is
  simply complete. Status is now the first thing `Check` looks at.
- **Status is authoritative over % done, not merely a synonym for it.**
  Marking a sub-task Done and leaving `% done` blank is safe — remaining goes
  to zero regardless. The reverse is deliberately not symmetric: `In Progress`
  with a blank `% done` reads as no progress and claims full effort, which
  overstates what is left rather than understating it. The safer direction,
  but worth knowing.
- **Every hardcoded column letter had to go.** Inserting the two actual-date
  columns meant first removing some thirty hand-written `$F{r}`-style
  references from `sheet_work.py` — a letter written in by hand silently
  changes meaning when anything to its left moves. This is the same fault
  that caused the v5→v6 migration bug, one layer down. Formulas now look up
  every column letter from `layout.py`.
- **Export never read the current-week pointer**, so a rebuild silently reset
  "now" to whatever the template ships with and every planned start moved
  with it — nothing would have *looked* broken. Schema v3: actual dates are
  JSON-owned observed facts, and a v2 payload imports with them blank.

### Phase 5 — Convergence diagnostics — **DONE**

- Quantified columns on `Tasks`: `Remaining`, `Scheduled`, `Shortfall`, and the
  week it finishes or "beyond horizon".
- The binding constraint per task: no capacity in range, horizon too short, or
  higher-priority work consuming the days.
- A **Plan health** block on `Gantt-High`: total shortfall, how many more weeks
  would absorb it, the bottleneck assignee, and the smallest single change that
  would make the plan converge.

Delivered as specified. Size: M, mostly formulas over grids that already
existed — confirmed true, no new hidden-sheet infrastructure was needed.

**`Scheduled` and `Remaining` are not the same axis, and the demo data makes
that concrete.** T-04 has ~42 days of untouched (`Remaining`) work, but only a
sliver of it — the `Shortfall` — actually fails to land inside the horizon.
Reading `Remaining` alone (all Phase 3 gave you) cannot tell those apart; a
task can be almost entirely unscheduled work and still converge fine, or be
mostly scheduled and still miss by a day. This is the concrete answer to
Appendix A's original complaint that "it is not clear how to make changes in
order to make the plan converge."

**The binding-constraint classification is deliberately simplified**, and
worth knowing precisely how: it checks capacity against a task's *default*
assignee only, never the true mix across sub-tasks with an assignee override.
Modelling the real mix would mean deduplicating shared capacity across
sub-tasks and tasks — which is what the spill-over engine already does, so a
second, independent attempt at it risks becoming a second scheduling engine
that can disagree with the first. Good enough to point at the right lever
(horizon, a specific assignee's capacity, or priority order); not a proof.

**"Weeks to absorb it" assumes the shortfall spreads evenly across everyone
at their average capacity.** That is optimistic exactly when the real
constraint is concentrated in one person or one skill — the case the
`Bottleneck assignee` row exists to surface separately, precisely because the
average-rate number understates the fix needed when there is a dominant
bottleneck. The Guide says this plainly rather than letting the number read
as more precise than it is.

**A formula that evaluates to `""` is stored as a blank cell on
recalculation, not a literal empty string** — caught by a test comparing
against a hardcoded `""` and failing with `None == ''`. Every prior "blank
when not applicable" column in this workbook happened to also be the *outer*
guard (`IF($A{r}="","",...)`), so this had not come up before; `Binding
constraint` is the first column that returns `""` from an *inner* branch on a
populated row. Worth remembering for any future diagnostic column: normalise
with `cell.value or ""` when a test compares against a literal blank.

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

### Phase 7 — Commit-log calibration for effort and proficiency (proposed)

Every number that drives scheduling — `base_days` per complexity, `proficiency`
per assignee — is currently a hand-entered guess. Phase 4 gives the project its
own history (`Actual start WW` / `Actual end WW`, `Status`, `% done`); once
enough sub-tasks have actually finished, that history can replace the guesses
with numbers the project itself produced. This is a calibration loop, not a
one-shot fix: run it periodically as more work completes and the estimates
should keep converging on reality.

**Depends on Phase 4's actuals existing in real, lived-in data** — this is not
buildable against the demo dataset, which has only a handful of finished
sub-tasks. It needs a project that has been run for a while first.

**Data sources:**

- Finished (`Done`) sub-tasks with recorded `Actual start WW` / `Actual end
  WW`, exported via `python -m gantt.exchange export` — complexity, assignee,
  and observed duration are all already there.
- `git log` for the project repository the work actually happened in (not this
  template's own repo — the *target* project's), read per author and per
  commit date. This is a signal that work happened and roughly when, not a
  measure of how much.

**Computation.** `base_days` and `proficiency` are both unknown and only their
ratio is observed (`effort = base_days[complexity] / proficiency[assignee]`),
so a single finished sub-task cannot solve for either — there are more
unknowns than equations. With enough observations across multiple assignees
and complexities, though, the system is over-determined and a least-squares
fit recovers both, the same way the template already fixes one assignee's
proficiency at 1.0 as a reference point. Sketch:

1. For every Done sub-task, compute observed effort as business days between
   `Actual start WW` and `Actual end WW` (coarse — see risks below).
2. Cross-reference commits authored by that sub-task's assignee falling inside
   the actual-start/actual-end window, as a corroborating signal that work
   happened and roughly how much — not as the primary effort measure.
3. Fit `log(observed_effort) = log(base_days[complexity]) -
   log(proficiency[assignee])` by least squares across all observations, one
   assignee pinned at `proficiency = 1.0` to resolve the scale ambiguity.
4. Report the fitted values alongside the current hand-entered ones and the
   sample size behind each — a complexity/assignee pair with one data point
   should say so, not be presented with the same confidence as one with fifty.

**Output: a report, not a write.** This has to stay advisory. `Config` and
`Assignees` are planning decisions, and the field-ownership rule established
in Phase 2 — the JSON/ingest side owns observed facts, the workbook owns
planning decisions — applies here too: a calibration run must never silently
rewrite `base_days` or `proficiency` underneath the planner. It proposes; a
human decides whether to type the new numbers in. Likely shape: a new
`python -m gantt.calibrate` CLI command (or a script alongside `exchange.py`)
that reads a JSON export plus a git log and prints a suggested-values table.

**Where this lives.** Not inside the workbook or the generator — git history
is specific to the *project* being managed, not to the Gantt template, which
has to stay generic enough to duplicate for any project (per this document's
own goal). A standalone module reading `gantt.exchange export` output plus
`git log` output is the right shape; it should not become a dependency of
`gantt/build.py`.

**Risks, named up front rather than discovered later:**

- Commit count or frequency is a weak, gameable effort proxy — assignees
  commit at wildly different granularities. Treat it as corroboration for the
  actual-date signal, never as the primary measure.
- `Actual start WW` / `Actual end WW` are week-granularity, so observed
  duration carries roughly ±6 days of noise per sub-task. Calibration needs
  many observations per complexity/assignee cell before the fit means
  anything; the report should say so rather than presenting a two-observation
  fit as settled.
- Git authorship does not automatically match the `Assignee` column — pairing,
  handle mismatches, and rebasing all break a naive name join. An explicit
  assignee ↔ git-author mapping is a prerequisite, not an assumption.
- This is opt-in, offline tooling run by a human periodically — not something
  the weekly ingest skill in Phase 6 should call automatically.

Size: L, and the phase most likely to need a second pass once it meets real
data — the sketch above is a starting point, not a finished spec.

## Sequencing

```
0 Portability ──► 1 Export ──► 2 Import/rebuild ──► 3 Status ──► 4 Actuals ──► 5 Diagnostics ──► 6 Skill spec ──► 7 Calibration
    DONE            DONE            DONE               DONE          DONE          DONE            (Opencode)      (proposed)
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
