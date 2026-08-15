"""Export of a workbook's input tabs to JSON.

The demo dataset is the oracle: `gantt/demo.py` is what the generator writes in,
so a faithful export has to give it back unchanged.
"""

import json
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from gantt import demo, exchange, layout as L  # noqa: E402
from gantt.build import build  # noqa: E402


@pytest.fixture(scope="module")
def generated(tmp_path_factory) -> Path:
    return build(tmp_path_factory.mktemp("exchange") / "plan.xlsx")


@pytest.fixture(scope="module")
def data(generated) -> dict:
    return exchange.export(generated)


def test_export_needs_no_recalculation(generated):
    """A freshly generated workbook has no cached formula results at all.

    openpyxl does not evaluate formulas, so a exporter that read a derived
    column would return None here and a stale value from an edited file. Export
    therefore reads only cells the user types into — this asserts the premise by
    confirming the derived cells really are empty of values.
    """
    ws = load_workbook(generated, data_only=True)[L.SUBTASKS]
    derived = ws.cell(row=L.SUB_FIRST_ROW, column=L.S_EFFORT).value
    assert derived is None, "premise broken: the file now carries cached values"

    exported = exchange.export(generated)
    assert exported["sub_tasks"][0]["id"] == "T-01.01"


def test_config_round_trips(data):
    config = data["config"]
    assert config["year"] == demo.YEAR
    assert config["start_week"] == demo.START_WEEK
    assert config["horizon_weeks"] == demo.HORIZON
    assert [(c["name"], c["base_days"]) for c in config["complexity"]] == \
        [(n, d) for n, d in demo.COMPLEXITY]
    assert [(p["name"], p["order"]) for p in config["priorities"]] == \
        [(n, o) for n, o in demo.PRIORITIES]


def test_assignees_round_trip(data):
    assert [(a["name"], a["proficiency"]) for a in data["assignees"]] == \
        [(n, p) for n, p in demo.ASSIGNEES]


def test_week_grids_round_trip(data):
    """Capacity and equipment come back as the same series that went in."""
    for row in data["capacity"]:
        expected = demo.CAPACITY[row["assignee"]][:L.WEEK_COLS]
        assert row["days_by_week_offset"] == expected, row["assignee"]
    for row in data["equipment"]:
        expected = demo.EQUIPMENT[row["type"]][:L.WEEK_COLS]
        assert row["units_by_week_offset"] == expected, row["type"]


def test_capacity_rows_follow_the_assignee_order(data):
    """The Capacity name column is a formula mirroring Assignees, so it cannot
    be read; row order is the only thing tying the two together."""
    assert [c["assignee"] for c in data["capacity"]] == \
        [a["name"] for a in data["assignees"]]


def test_holidays_round_trip(data):
    assert [(h["date"], h["name"]) for h in data["holidays"]] == list(demo.HOLIDAYS)


def test_tasks_round_trip(data):
    assert len(data["tasks"]) == len(demo.TASKS)
    for got, expected in zip(data["tasks"], demo.TASKS):
        tid, name, category, priority, complexity, equipment, assignee, start, _ = expected
        assert (got["id"], got["name"], got["category"], got["priority"],
                got["complexity"], got["equipment"], got["default_assignee"],
                got["earliest_start_week"]) == \
            (tid, name, category, priority, complexity, equipment, assignee, start)


def test_sub_tasks_round_trip(data):
    expected = demo.subtasks()
    assert len(data["sub_tasks"]) == len(expected)
    for got, (parent, name, complexity, override) in zip(data["sub_tasks"], expected):
        assert got["parent"] == parent
        assert got["name"] == name
        assert got["complexity"] == complexity
        assert got["assignee"] == (override or None)


def test_sub_task_ids_match_the_workbook_rule(data):
    """Recomputed here, so they must agree with what the formula would build."""
    counts: dict[str, int] = {}
    for row in data["sub_tasks"]:
        counts[row["parent"]] = counts.get(row["parent"], 0) + 1
        assert row["id"] == f"{row['parent']}.{counts[row['parent']]:02d}"


def test_no_derived_column_leaks_into_the_export(data):
    """Derived values would be stale the moment anything upstream changed."""
    for row in data["sub_tasks"]:
        assert set(row) == {"id", "parent", "name", "complexity", "assignee"}
    for row in data["tasks"]:
        assert "effort" not in row and "check" not in row and "start_ww" not in row


def test_output_is_json_serialisable(data):
    """Dates in particular have to survive; openpyxl hands them back as datetimes."""
    text = exchange.to_json(data)
    assert json.loads(text) == data


def test_blank_rows_are_skipped(data):
    """The template pre-builds 30 task and 600 sub-task rows."""
    assert len(data["tasks"]) == 5
    assert len(data["sub_tasks"]) == 33
    assert all(row["parent"] for row in data["sub_tasks"])


def test_cli_writes_a_file(generated, tmp_path):
    out = tmp_path / "plan.json"
    assert exchange._main(["export", str(generated), "-o", str(out)]) == 0
    assert json.loads(out.read_text(encoding="utf-8"))["schema_version"] == 1
