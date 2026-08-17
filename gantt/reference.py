"""Pure-Python model of the scheduling algorithm.

This exists solely as an oracle for the tests. It is deliberately written
independently of the formula generator: if both agree on the demo dataset, the
formulas encode the algorithm the design describes.
"""

from dataclasses import dataclass, field

from . import calendar_utils as C, demo


@dataclass
class SubTask:
    parent: str
    name: str
    complexity: str
    assignee: str
    effort: float
    status: str
    remaining: float
    key: int
    index: int
    consumed: float = 0.0
    actual_start: int | None = None
    actual_end: int | None = None
    rank: int = 0
    weekly: dict[int, float] = field(default_factory=dict)
    daily: dict[str, float] = field(default_factory=dict)
    # History, kept apart from `weekly` on purpose: the scheduling algorithm
    # never sees it, exactly as in the workbook.
    actual_weekly: dict[int, float] = field(default_factory=dict)


def _weeks(horizon: int | None = None) -> list[int]:
    """The active weeks. `horizon` mirrors the Horizon cell on Config."""
    return [demo.START_WEEK + i for i in range(horizon or demo.HORIZON)]


def _lookups():
    base = dict(demo.COMPLEXITY)
    prof = dict(demo.ASSIGNEES)
    prio = dict(demo.PRIORITIES)
    tasks = {t[0]: t for t in demo.TASKS}
    task_row = {t[0]: i + 1 for i, t in enumerate(demo.TASKS)}
    return base, prof, prio, tasks, task_row


def build_subtasks() -> list[SubTask]:
    """Sub-tasks with effort, sort key and rank resolved."""
    base, prof, prio, tasks, task_row = _lookups()
    seen: dict[str, int] = {}
    rows: list[SubTask] = []

    for sub in demo.subtasks():
        parent, name, complexity = sub["parent"], sub["name"], sub["complexity"]
        seen[parent] = seen.get(parent, 0) + 1
        index = seen[parent]
        task = tasks[parent]
        assignee = sub["assignee"] or task[6]
        effort = base[complexity] / prof[assignee]
        status = sub["status"]
        # Done claims no capacity regardless of what % done says; otherwise the
        # balance is what remains of effort after the recorded progress.
        remaining = 0.0 if status == "Done" else effort * (1 - sub["pct_done"] / 100)
        key = (prio[task[3]] * 1_000_000_000
               + task[7] * 1_000_000
               + task_row[parent] * 1_000
               + index)
        rows.append(SubTask(parent, name, complexity, assignee, effort, status,
                            remaining, key, index,
                            consumed=effort - remaining,
                            actual_start=sub.get("actual_start_week"),
                            actual_end=sub.get("actual_end_week")))

    for rank, row in enumerate(sorted(rows, key=lambda s: s.key), start=1):
        row.rank = rank
    return rows


def planning_start(parent: str, current_week: int | None = None) -> int:
    """The first week work may be planned into.

    The later of the task's earliest start and the current week: nothing is
    scheduled into the past, and an earliest start that has already gone by
    means "as soon as possible" rather than being an error.
    """
    _, _, _, tasks, _ = _lookups()
    current = demo.CURRENT_WEEK if current_week is None else current_week
    return max(tasks[parent][7], current)


def schedule_weekly(rows: list[SubTask], horizon: int | None = None,
                    current_week: int | None = None) -> list[SubTask]:
    """Spill-over across weeks, in rank order, against assignee capacity."""
    _, _, _, tasks, _ = _lookups()
    weeks = _weeks(horizon)
    used: dict[tuple[str, int], float] = {}

    for row in sorted(rows, key=lambda s: s.rank):
        remaining = row.remaining
        start = planning_start(row.parent, current_week)
        for w in weeks:
            if w < start or remaining <= 1e-9:
                row.weekly[w] = 0.0
                continue
            cap = week_capacity(row.assignee, w, weeks)
            free = max(0.0, cap - used.get((row.assignee, w), 0.0))
            take = min(remaining, free)
            row.weekly[w] = take
            used[(row.assignee, w)] = used.get((row.assignee, w), 0.0) + take
            remaining -= take
    return rows


def _holidays() -> set:
    from datetime import datetime
    return {datetime.strptime(d, "%Y-%m-%d").date() for d, _ in demo.HOLIDAYS}


def working_days(week: int) -> int:
    """Sun..Thu days in a week that are not company holidays."""
    holidays = _holidays()
    return sum(1 for d in C.workdays(demo.YEAR, week) if d not in holidays)


def week_capacity(name: str, week: int, weeks: list[int]) -> float:
    """Entered capacity reduced pro-rata for company holidays.

    The entered figure is what the person has in a normal five-day week; a week
    shortened by a holiday scales it down in proportion.
    """
    raw = demo.CAPACITY[name][weeks.index(week)]
    return raw * working_days(week) / C.WORKDAYS_PER_WEEK


def schedule_daily(rows: list[SubTask], window_start: int, window_weeks: int,
                   horizon: int | None = None,
                   current_week: int | None = None) -> list[SubTask]:
    """Same algorithm at day granularity across the visible window."""
    _, _, _, tasks, _ = _lookups()
    weeks = _weeks(horizon)
    holidays = _holidays()
    used: dict[tuple[str, str], float] = {}

    window = [w for w in range(window_start, window_start + window_weeks) if w in weeks]
    day_caps: dict[tuple[str, str], float] = {}
    for w in window:
        for name, _ in demo.ASSIGNEES:
            # Weekly capacity is already pro-rated, so the day rate is just the
            # ordinary daily rate; holidays remove whole days rather than
            # concentrating the week's work into the ones that remain.
            per_day = demo.CAPACITY[name][weeks.index(w)] / C.WORKDAYS_PER_WEEK
            for d in C.workdays(demo.YEAR, w):
                day_caps[(name, d.isoformat())] = 0.0 if d in holidays else per_day

    for row in sorted(rows, key=lambda s: s.rank):
        burned = sum(v for w, v in row.weekly.items() if w < window_start)
        remaining = max(0.0, row.remaining - burned)
        start = planning_start(row.parent, current_week)
        for w in window:
            for d in C.workdays(demo.YEAR, w):
                iso = d.isoformat()
                if w < start or remaining <= 1e-9:
                    row.daily[iso] = 0.0
                    continue
                free = max(0.0, day_caps[(row.assignee, iso)]
                           - used.get((row.assignee, iso), 0.0))
                take = min(remaining, free)
                row.daily[iso] = take
                used[(row.assignee, iso)] = used.get((row.assignee, iso), 0.0) + take
                remaining -= take
    return rows


def solve(window_start: int | None = None, window_weeks: int = 4,
          horizon: int | None = None, current_week: int | None = None) -> list[SubTask]:
    rows = schedule_weekly(build_subtasks(), horizon, current_week)
    rows = schedule_daily(rows, window_start or demo.START_WEEK, window_weeks,
                          horizon, current_week)
    return schedule_actuals(rows, horizon, current_week)


def schedule_actuals(rows: list[SubTask], horizon: int | None = None,
                     current_week: int | None = None) -> list[SubTask]:
    """Spread consumed effort across the weeks it was recorded against.

    Evenly, and clamped to end before the current week, so history and plan
    never share a column and a view can simply add them.
    """
    weeks = _weeks(horizon)
    holidays = _holidays()
    current = demo.CURRENT_WEEK if current_week is None else current_week
    current_pos = weeks.index(current) if current in weeks else 0

    for row in rows:
        row.actual_weekly = {w: 0.0 for w in weeks}
        if row.consumed <= 0 or row.actual_start not in weeks:
            continue
        start = weeks.index(row.actual_start)
        end = weeks.index(row.actual_end) if row.actual_end in weeks else current_pos - 1
        end = min(end, current_pos - 1)
        if end < start:
            continue

        share = row.consumed / (end - start + 1)
        for position in range(start, end + 1):
            week = weeks[position]
            row.actual_weekly[week] = share
            working = [d for d in C.workdays(demo.YEAR, week) if d not in holidays]
            for day in working:
                row.daily[day.isoformat()] = share / len(working)
    return rows


def assignee_load(rows: list[SubTask], horizon: int | None = None) -> dict[str, dict[int, float]]:
    out = {name: {w: 0.0 for w in _weeks(horizon)} for name, _ in demo.ASSIGNEES}
    # What a week actually cost plus what it is planned to cost. Behind the
    # current week these are history, ahead of it the plan; they never overlap.
    for row in rows:
        for w, v in row.weekly.items():
            out[row.assignee][w] += v
        for w, v in row.actual_weekly.items():
            out[row.assignee][w] += v
    return out


def task_load(rows: list[SubTask], horizon: int | None = None) -> dict[str, dict[int, float]]:
    out = {t[0]: {w: 0.0 for w in _weeks(horizon)} for t in demo.TASKS}
    # What a week actually cost plus what it is planned to cost. Behind the
    # current week these are history, ahead of it the plan; they never overlap.
    for row in rows:
        for w, v in row.weekly.items():
            out[row.parent][w] += v
        for w, v in row.actual_weekly.items():
            out[row.parent][w] += v
    return out


def equipment_demand(rows: list[SubTask], horizon: int | None = None) -> dict[str, dict[int, int]]:
    tasks = {t[0]: t for t in demo.TASKS}
    loads = task_load(rows, horizon)
    out = {name: {w: 0 for w in _weeks(horizon)} for name in demo.EQUIPMENT}
    for tid, per_week in loads.items():
        equip = tasks[tid][5]
        for w, v in per_week.items():
            if v > 0:
                out[equip][w] += 1
    return out


# --- Convergence diagnostics (Phase 5) --------------------------------------

def task_effort(tid: str, rows: list[SubTask]) -> float:
    return sum(s.effort for s in rows if s.parent == tid)


def task_scheduled(tid: str, rows: list[SubTask], horizon: int | None = None) -> float:
    """History plus plan, confined to the horizon — what actually landed."""
    return sum(task_load(rows, horizon)[tid].values())


def task_shortfall(tid: str, rows: list[SubTask], horizon: int | None = None) -> float:
    return max(0.0, task_effort(tid, rows) - task_scheduled(tid, rows, horizon))


def binding_constraint(tid: str, rows: list[SubTask], horizon: int | None = None,
                       current_week: int | None = None) -> str:
    """Best-effort diagnosis of why a task does not converge.

    Deliberately simplified to mirror the workbook: capacity is checked
    against the task's *default* assignee only, not the true mix across
    overridden sub-tasks — computing the real mix would mean deduplicating
    shared assignees across sub-tasks, which risks becoming a second
    scheduling engine. "" means the task has no shortfall.
    """
    _, _, _, tasks, _ = _lookups()
    shortfall = task_shortfall(tid, rows, horizon)
    if shortfall <= 1e-9:
        return ""

    weeks = _weeks(horizon)
    start = planning_start(tid, current_week)
    remaining_weeks = [w for w in weeks if w >= start]
    if len(remaining_weeks) * C.WORKDAYS_PER_WEEK < shortfall:
        return "horizon too short"

    default_assignee = tasks[tid][6]
    available = sum(week_capacity(default_assignee, w, weeks) for w in remaining_weeks)
    if available < shortfall:
        return "no capacity in range"

    return "higher-priority work"


def plan_health(rows: list[SubTask], horizon: int | None = None,
                current_week: int | None = None) -> dict:
    """The Gantt-High summary block: how much does not fit, and where to look.

    `weeks_to_absorb` and `bottleneck_assignee` are both computed only over
    weeks from the current one to the horizon's end — saturation in a week
    that has already happened is not something anyone can act on.
    """
    _, _, _, tasks, _ = _lookups()
    weeks = _weeks(horizon)
    current = demo.CURRENT_WEEK if current_week is None else current_week
    future = [w for w in weeks if w >= current]

    total_shortfall = sum(task_shortfall(t[0], rows, horizon) for t in demo.TASKS)

    names = [n for n, _ in demo.ASSIGNEES]
    total_capacity = sum(week_capacity(n, w, weeks) for n in names for w in future)
    avg_weekly_capacity = (total_capacity / len(names) / len(future)
                           if names and future else 0.0)
    weeks_to_absorb = (0 if total_shortfall <= 1e-9 else
                       None if avg_weekly_capacity <= 0 else
                       -(-total_shortfall // avg_weekly_capacity))  # ceil

    loads = assignee_load(rows, horizon)
    saturation = {}
    for name in names:
        saturation[name] = sum(
            1 for w in future
            if week_capacity(name, w, weeks) > 0
            and loads[name][w] >= week_capacity(name, w, weeks) - 1e-9)
    bottleneck = max(saturation, key=saturation.get) if names else None
    if bottleneck is not None and saturation[bottleneck] <= 0:
        bottleneck = "none"

    return {
        "total_shortfall": total_shortfall,
        "weeks_to_absorb": weeks_to_absorb,
        "bottleneck_assignee": bottleneck,
        "saturation": saturation,
    }
