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
    key: int
    index: int
    rank: int = 0
    weekly: dict[int, float] = field(default_factory=dict)
    daily: dict[str, float] = field(default_factory=dict)


def _weeks() -> list[int]:
    return [demo.START_WEEK + i for i in range(demo.HORIZON)]


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

    for parent, name, complexity, override in demo.subtasks():
        seen[parent] = seen.get(parent, 0) + 1
        index = seen[parent]
        task = tasks[parent]
        assignee = override or task[6]
        effort = base[complexity] / prof[assignee]
        key = (prio[task[3]] * 1_000_000_000
               + task[7] * 1_000_000
               + task_row[parent] * 1_000
               + index)
        rows.append(SubTask(parent, name, complexity, assignee, effort, key, index))

    for rank, row in enumerate(sorted(rows, key=lambda s: s.key), start=1):
        row.rank = rank
    return rows


def schedule_weekly(rows: list[SubTask]) -> list[SubTask]:
    """Spill-over across weeks, in rank order, against assignee capacity."""
    _, _, _, tasks, _ = _lookups()
    weeks = _weeks()
    used: dict[tuple[str, int], float] = {}

    for row in sorted(rows, key=lambda s: s.rank):
        remaining = row.effort
        start = tasks[row.parent][7]
        for w in weeks:
            if w < start or remaining <= 1e-9:
                row.weekly[w] = 0.0
                continue
            cap = demo.CAPACITY[row.assignee][weeks.index(w)]
            free = max(0.0, cap - used.get((row.assignee, w), 0.0))
            take = min(remaining, free)
            row.weekly[w] = take
            used[(row.assignee, w)] = used.get((row.assignee, w), 0.0) + take
            remaining -= take
    return rows


def _holidays() -> set:
    from datetime import datetime
    return {datetime.strptime(d, "%Y-%m-%d").date() for d, _ in demo.HOLIDAYS}


def schedule_daily(rows: list[SubTask], window_start: int, window_weeks: int) -> list[SubTask]:
    """Same algorithm at day granularity across the visible window."""
    _, _, _, tasks, _ = _lookups()
    weeks = _weeks()
    holidays = _holidays()
    used: dict[tuple[str, str], float] = {}

    window = [w for w in range(window_start, window_start + window_weeks) if w in weeks]
    day_caps: dict[tuple[str, str], float] = {}
    for w in window:
        days = C.workdays(demo.YEAR, w)
        working = [d for d in days if d not in holidays]
        for name, _ in demo.ASSIGNEES:
            week_cap = demo.CAPACITY[name][weeks.index(w)]
            per_day = week_cap / len(working) if working else 0.0
            for d in days:
                day_caps[(name, d.isoformat())] = 0.0 if d in holidays else per_day

    for row in sorted(rows, key=lambda s: s.rank):
        burned = sum(v for w, v in row.weekly.items() if w < window_start)
        remaining = max(0.0, row.effort - burned)
        start = tasks[row.parent][7]
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


def solve(window_start: int | None = None, window_weeks: int = 4) -> list[SubTask]:
    rows = schedule_weekly(build_subtasks())
    return schedule_daily(rows, window_start or demo.START_WEEK, window_weeks)


def assignee_load(rows: list[SubTask]) -> dict[str, dict[int, float]]:
    out = {name: {w: 0.0 for w in _weeks()} for name, _ in demo.ASSIGNEES}
    for row in rows:
        for w, v in row.weekly.items():
            out[row.assignee][w] += v
    return out


def task_load(rows: list[SubTask]) -> dict[str, dict[int, float]]:
    out = {t[0]: {w: 0.0 for w in _weeks()} for t in demo.TASKS}
    for row in rows:
        for w, v in row.weekly.items():
            out[row.parent][w] += v
    return out


def equipment_demand(rows: list[SubTask]) -> dict[str, dict[int, int]]:
    tasks = {t[0]: t for t in demo.TASKS}
    loads = task_load(rows)
    out = {name: {w: 0 for w in _weeks()} for name in demo.EQUIPMENT}
    for tid, per_week in loads.items():
        equip = tasks[tid][5]
        for w, v in per_week.items():
            if v > 0:
                out[equip][w] += 1
    return out
