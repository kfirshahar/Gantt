"""Demo dataset shipped inside the template.

Sized to make input ergonomics visible: five parent tasks spanning the full
1..20 sub-task range, three assignees at differing proficiency, two equipment
types, and a 12-week horizon starting WW33 2026.
"""

YEAR = 2026
START_WEEK = 33
HORIZON = 12

COMPLEXITY = [("Simple", 1.0), ("Medium", 2.5), ("Complex", 5.0)]
PRIORITIES = [("P1", 1), ("P2", 2), ("P3", 3)]

ASSIGNEES = [
    ("Alice", 1.2),
    ("Bob", 1.0),
    ("Carol", 0.8),
]

# Available days per assignee per week, WW33..WW44. Deliberately uneven:
# Bob takes WW37 off entirely, Carol runs at four days a week.
CAPACITY = {
    "Alice": [5, 5, 4.5, 5, 5, 5, 3, 5, 5, 5, 4, 5],
    "Bob":   [5, 5, 5,   5, 0, 5, 5, 5, 5, 4, 5, 5],
    "Carol": [4, 4, 4,   4, 4, 4, 4, 3, 4, 4, 4, 4],
}

# Units available per equipment type per week. Rig-A dips to one unit in WW35.
EQUIPMENT = {
    "Rig-A": [2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2],
    "Rig-B": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
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


def subtasks():
    """Yield (parent_id, name, complexity, assignee_override) rows."""
    rows = []
    for task in TASKS:
        tid, _, _, _, _, _, _, _, n_subs = task
        names = _SUB_NAMES.get(tid) or [f"Step {i + 1}" for i in range(n_subs)]
        for i in range(n_subs):
            name = names[i] if i < len(names) else f"Step {i + 1}"
            complexity = _CYCLE[i % len(_CYCLE)]
            # One override, to demonstrate that a sub-task can differ from its
            # parent's default assignee.
            override = "Alice" if (tid == "T-04" and i == 0) else ""
            rows.append((tid, name, complexity, override))
    return rows
