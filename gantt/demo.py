"""Demo dataset shipped inside the template.

Sized to make input ergonomics visible: five parent tasks spanning the full
1..20 sub-task range, three assignees at differing proficiency, two equipment
types, and a 12-week horizon starting WW33 2026.
"""

YEAR = 2026
START_WEEK = 33
HORIZON = 12
# Two weeks in, so the demo shows a project under way rather than one
# about to begin.
CURRENT_WEEK = 35

COMPLEXITY = [("Simple", 1.0), ("Medium", 2.5), ("Complex", 5.0)]
# Order carries the meaning: not started, in progress, finished.
STATUSES = ["TODO", "In Progress", "Done"]
PRIORITIES = [("P1", 1), ("P2", 2), ("P3", 3)]

ASSIGNEES = [
    ("Alice", 1.2),
    ("Bob", 1.0),
    ("Carol", 0.8),
]

# Available days per assignee per week from WW33. Deliberately uneven: Bob takes
# WW37 off entirely, Carol runs at four days a week. Filled out to WEEK_COLS so
# that raising the Horizon on Config immediately has capacity to schedule into.
CAPACITY = {
    "Alice": [5, 5, 4.5, 5, 5, 5, 3, 5, 5, 5, 4, 5] + [5] * 14,
    "Bob":   [5, 5, 5,   5, 0, 5, 5, 5, 5, 4, 5, 5] + [5] * 14,
    "Carol": [4, 4, 4,   4, 4, 4, 4, 3, 4, 4, 4, 4] + [4] * 14,
}

# Units available per equipment type per week. Rig-A dips to one unit in WW35.
EQUIPMENT = {
    "Rig-A": [2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2] + [2] * 14,
    "Rig-B": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1] + [1] * 14,
}

# Example dates only — replace with your own calendar before planning against it.
HOLIDAYS = [
    ("2026-09-13", "Example holiday A"),
    ("2026-09-21", "Example holiday B"),
    ("2026-09-28", "Example holiday C"),
]

# (id, name, category, priority, complexity, equipment, default assignee,
#  earliest start week, number of sub-tasks)
TASKS = [
    ("T-01", "Design review",    "Design",  "P1", "Complex", "Rig-A", "Alice", 33, 3),
    ("T-02", "Prototype build",  "Build",   "P2", "Complex", "Rig-A", "Bob",   33, 8),
    ("T-03", "Vendor selection", "Procure", "P3", "Simple",  "Rig-B", "Carol", 34, 1),
    ("T-04", "Integration test", "Test",    "P1", "Complex", "Rig-B", "Bob",   35, 15),
    ("T-05", "Documentation",    "Docs",    "P3", "Medium",  "Rig-B", "Carol", 36, 6),
]

_SUB_NAMES = {
    "T-01": ["Collect requirements", "Draft layout", "Review with stakeholders"],
    "T-02": ["Source parts", "Machine frame", "Wire harness", "Mount sensors",
             "Firmware flash", "Bench test", "Rework", "Sign-off"],
    "T-03": ["Compare quotes"],
    "T-05": ["Outline", "Write setup guide", "Write ops guide", "Diagrams",
             "Internal review", "Publish"],
}

_CYCLE = ["Simple", "Medium", "Complex"]


# A project part-way through, so the demo exercises the states rather than
# showing every row as untouched: T-01 finished, T-02 under way.
_PROGRESS = {
    "T-01": [("Done", 100), ("Done", 100), ("Done", 100)],
    "T-02": [("Done", 100), ("In Progress", 60), ("In Progress", 25)],
}


def subtasks():
    """One dict per sub-task, in sheet order."""
    rows = []
    for task in TASKS:
        tid, _, _, _, _, _, _, _, n_subs = task
        names = _SUB_NAMES.get(tid) or [f"Step {i + 1}" for i in range(n_subs)]
        progress = _PROGRESS.get(tid, [])
        for i in range(n_subs):
            status, pct = progress[i] if i < len(progress) else ("TODO", 0)
            rows.append({
                "parent": tid,
                "name": names[i] if i < len(names) else f"Step {i + 1}",
                "complexity": _CYCLE[i % len(_CYCLE)],
                # One override, to show a sub-task differing from its parent's
                # default assignee.
                "assignee": "Alice" if (tid == "T-04" and i == 0) else "",
                "status": status,
                "pct_done": pct,
            })
    return rows
