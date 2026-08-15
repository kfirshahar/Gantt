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
    rank: int = 0
    weekly: dict[int, float] = field(default_factory=dict)
    daily: dict[str, float] = field(default_factory=dict)


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
                            remaining, key, index))

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
    return schedule_daily(rows, window_start or demo.START_WEEK, window_weeks,
                          horizon, current_week)


def assignee_load(rows: list[SubTask], horizon: int | None = None) -> dict[str, dict[int, float]]:
    out = {name: {w: 0.0 for w in _weeks(horizon)} for name, _ in demo.ASSIGNEES}
    for row in rows:
        for w, v in row.weekly.items():
            out[row.assignee][w] += v
    return out


def task_load(rows: list[SubTask], horizon: int | None = None) -> dict[str, dict[int, float]]:
    out = {t[0]: {w: 0.0 for w in _weeks(horizon)} for t in demo.TASKS}
    for row in rows:
        for w, v in row.weekly.items():
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
