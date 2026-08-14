"""Single source of truth for where everything lives in the workbook.

Every sheet name, pre-built row count and column position is declared here so
that formulas built in other modules never hard-code an address.
"""

from openpyxl.utils import get_column_letter

# --- Sheet names -----------------------------------------------------------

GUIDE = "Guide"
CONFIG = "Config"
ASSIGNEES = "Assignees"
CAPACITY = "Capacity"
EQUIPMENT = "Equipment"
HOLIDAYS = "Holidays"
TASKS = "Tasks"
SUBTASKS = "Sub-Tasks"
GANTT_HIGH = "Gantt-High"
GANTT_DEEP = "Gantt-Deep"
CALC_WEEK = "CalcWeek"
CALC_DAY = "CalcDay"

INPUT_SHEETS = [GUIDE, CONFIG, ASSIGNEES, CAPACITY, EQUIPMENT, HOLIDAYS, TASKS, SUBTASKS]
OUTPUT_SHEETS = [GANTT_HIGH, GANTT_DEEP]
HIDDEN_SHEETS = [CALC_WEEK, CALC_DAY]

# --- Capacities (pre-built row counts) -------------------------------------

MAX_ASSIGNEES = 10
MAX_EQUIPMENT = 10
MAX_TASKS = 30
MAX_SUBTASKS = 600
MAX_HOLIDAYS = 50

# Week columns physically built into every grid. This is the ceiling; how many
# are actually *active* is set by the Horizon cell on Config at runtime, so
# lengthening a plan up to this many weeks needs no regeneration.
WEEK_COLS = 26
MAX_DEEP_WEEKS = 8          # soft cap on the deep-dive window
WORKDAYS_PER_WEEK = 5       # Sunday..Thursday
DEEP_DAY_COLS = MAX_DEEP_WEEKS * WORKDAYS_PER_WEEK

DAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu"]

# --- Config sheet ----------------------------------------------------------

CFG_YEAR_ROW = 3
CFG_START_WEEK_ROW = 4
CFG_HORIZON_ROW = 5
CFG_COMPLEXITY_HDR = 8
CFG_COMPLEXITY_FIRST = 9
CFG_COMPLEXITY_COUNT = 3
CFG_PRIORITY_HDR = 13
CFG_PRIORITY_FIRST = 14
CFG_PRIORITY_COUNT = 3

# --- Grid sheets (Capacity / Equipment): week columns start at B -----------

GRID_FIRST_WEEK_COL = 2      # column B
GRID_FIRST_DATA_ROW = 2

# --- Tasks sheet columns ---------------------------------------------------

T_ID = 1
T_NAME = 2
T_CATEGORY = 3
T_PRIORITY = 4
T_COMPLEXITY = 5
T_EQUIPMENT = 6
T_DEF_ASSIGNEE = 7
T_START_WW = 8
T_N_SUBS = 9
T_EFFORT = 10
T_CALC_START = 11
T_CALC_END = 12
T_CHECK = 13
TASK_FIRST_ROW = 2

TASK_HEADERS = {
    T_ID: "Task ID",
    T_NAME: "Task name",
    T_CATEGORY: "Category",
    T_PRIORITY: "Priority",
    T_COMPLEXITY: "Complexity",
    T_EQUIPMENT: "Equipment type",
    T_DEF_ASSIGNEE: "Default assignee",
    T_START_WW: "Earliest start WW",
    T_N_SUBS: "# Sub-tasks",
    T_EFFORT: "Effort (days)",
    T_CALC_START: "Start WW",
    T_CALC_END: "End WW",
    T_CHECK: "Check",
}
TASK_INPUT_COLS = [T_ID, T_NAME, T_CATEGORY, T_PRIORITY, T_COMPLEXITY,
                   T_EQUIPMENT, T_DEF_ASSIGNEE, T_START_WW]
TASK_DERIVED_COLS = [T_N_SUBS, T_EFFORT, T_CALC_START, T_CALC_END, T_CHECK]

# --- Sub-Tasks sheet columns ----------------------------------------------

S_PARENT = 1
S_ID = 2
S_NAME = 3
S_COMPLEXITY = 4
S_ASSIGNEE = 5
S_EFF_ASSIGNEE = 6
S_EFFORT = 7
S_KEY = 8
S_RANK = 9
S_CHECK = 10
SUB_FIRST_ROW = 2

SUB_HEADERS = {
    S_PARENT: "Parent task",
    S_ID: "Sub-task ID",
    S_NAME: "Sub-task name",
    S_COMPLEXITY: "Complexity",
    S_ASSIGNEE: "Assignee (blank = inherit)",
    S_EFF_ASSIGNEE: "Effective assignee",
    S_EFFORT: "Effort (days)",
    S_KEY: "Sort key",
    S_RANK: "Rank",
    S_CHECK: "Check",
}
SUB_INPUT_COLS = [S_PARENT, S_NAME, S_COMPLEXITY, S_ASSIGNEE]
SUB_DERIVED_COLS = [S_ID, S_EFF_ASSIGNEE, S_EFFORT, S_KEY, S_RANK, S_CHECK]

# --- CalcWeek sheet --------------------------------------------------------

# Header rows. Three different numbers describe a week column and conflating
# them is what produced "WW58": POS is the internal key used by every lookup and
# comparison, WEEK is the real calendar week the user types and reads, and ABS
# is a continuous count used only to get from a column to a date.
CW_POS_ROW = 1               # 1..WEEK_COLS
CW_ABS_ROW = 2               # CfgStartWeek + pos - 1, may exceed 52
CW_SUN_ROW = 3               # Sunday opening the week
CW_YEAR_ROW = 4              # calendar year the week belongs to
CW_WEEK_ROW = 5              # calendar week number, wraps at the year boundary
CW_WORKDAYS_ROW = 6          # Sun..Thu days left after company holidays
CW_FACTOR_ROW = 7            # workdays / 5, the pro-rata holiday factor
CW_ACTIVE_ROW = 8            # 1 where the week is inside the configured horizon
CW_LABEL_ROW = 9             # display text, e.g. WW01 '27
CW_HDR_ROW = 10
CW_FIRST_ROW = 11
CW_RANK = 1
CW_SUBID = 2
CW_ASSIGNEE = 3
CW_EFFORT = 4
CW_START_WW = 5
CW_PARENT = 6
CW_FIRST_WEEK_COL = 7        # column G

# Rollup blocks, all aligned to the same week columns.
CW_TASKWEEK_HDR = CW_FIRST_ROW + MAX_SUBTASKS + 2
CW_TASKWEEK_FIRST = CW_TASKWEEK_HDR + 1
CW_AWWEEK_HDR = CW_TASKWEEK_FIRST + MAX_TASKS + 2      # assignee load per week
CW_AWWEEK_FIRST = CW_AWWEEK_HDR + 1
CW_EQPWEEK_HDR = CW_AWWEEK_FIRST + MAX_ASSIGNEES + 2   # equipment demand per week
CW_EQPWEEK_FIRST = CW_EQPWEEK_HDR + 1

# --- CalcDay sheet ---------------------------------------------------------

CD_POS_ROW = 1               # position into the week grid, not a week number
CD_DATE_ROW = 2
CD_HOLIDAY_ROW = 3
CD_WORKDAYS_ROW = 4
CD_INWINDOW_ROW = 5
CD_LABEL_ROW = 6
CD_HDR_ROW = 7
CD_FIRST_ROW = 8
CD_RANK = 1
CD_SUBID = 2
CD_ASSIGNEE = 3
CD_REMAINING = 4
CD_START_WW = 5
CD_PARENT = 6
CD_FIRST_DAY_COL = 7         # column G

# --- Gantt-Deep window inputs ---------------------------------------------

GD_WINDOW_START_CELL = "B3"
GD_WINDOW_WEEKS_CELL = "B4"


def col(idx: int) -> str:
    """1-based column index -> letter."""
    return get_column_letter(idx)


def sheet_ref(name: str) -> str:
    """Quote a sheet name for use in a formula when it needs it."""
    return f"'{name}'" if not name.isalnum() else name


def abs_range(sheet: str, c1: int, r1: int, c2: int, r2: int) -> str:
    return f"{sheet_ref(sheet)}!${col(c1)}${r1}:${col(c2)}${r2}"


def abs_cell(sheet: str, c: int, r: int) -> str:
    return f"{sheet_ref(sheet)}!${col(c)}${r}"


def week_col(first_col: int, index: int) -> int:
    """Column index of the `index`-th week column (0-based)."""
    return first_col + index
