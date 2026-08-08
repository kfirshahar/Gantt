"""Workbook defined names.

Formulas throughout the template reference these rather than raw addresses, so
moving a block means editing one entry here.
"""

from openpyxl.workbook.defined_name import DefinedName

from . import layout as L

_LAST_ASG = L.GRID_FIRST_DATA_ROW + L.MAX_ASSIGNEES - 1
_LAST_EQP = L.GRID_FIRST_DATA_ROW + L.MAX_EQUIPMENT - 1
_LAST_TASK = L.TASK_FIRST_ROW + L.MAX_TASKS - 1
_LAST_SUB = L.SUB_FIRST_ROW + L.MAX_SUBTASKS - 1
_LAST_HOL = L.GRID_FIRST_DATA_ROW + L.MAX_HOLIDAYS - 1
_LAST_WEEK_COL = L.GRID_FIRST_WEEK_COL + L.HORIZON_WEEKS - 1
_CW_LAST_WEEK_COL = L.CW_FIRST_WEEK_COL + L.HORIZON_WEEKS - 1
_CW_LAST_ROW = L.CW_FIRST_ROW + L.MAX_SUBTASKS - 1
_TW_LAST_ROW = L.CW_TASKWEEK_FIRST + L.MAX_TASKS - 1
_AW_LAST_ROW = L.CW_AWWEEK_FIRST + L.MAX_ASSIGNEES - 1
_EQD_LAST_ROW = L.CW_EQPWEEK_FIRST + L.MAX_EQUIPMENT - 1

NAMES = {
    # Config
    "CfgYear": L.abs_cell(L.CONFIG, 2, L.CFG_YEAR_ROW),
    "CfgStartWeek": L.abs_cell(L.CONFIG, 2, L.CFG_START_WEEK_ROW),
    "CfgHorizon": L.abs_cell(L.CONFIG, 2, L.CFG_HORIZON_ROW),
    "CplxNames": L.abs_range(L.CONFIG, 1, L.CFG_COMPLEXITY_FIRST,
                             1, L.CFG_COMPLEXITY_FIRST + L.CFG_COMPLEXITY_COUNT - 1),
    "CplxDays": L.abs_range(L.CONFIG, 2, L.CFG_COMPLEXITY_FIRST,
                            2, L.CFG_COMPLEXITY_FIRST + L.CFG_COMPLEXITY_COUNT - 1),
    "PrioNames": L.abs_range(L.CONFIG, 1, L.CFG_PRIORITY_FIRST,
                             1, L.CFG_PRIORITY_FIRST + L.CFG_PRIORITY_COUNT - 1),
    "PrioRank": L.abs_range(L.CONFIG, 2, L.CFG_PRIORITY_FIRST,
                            2, L.CFG_PRIORITY_FIRST + L.CFG_PRIORITY_COUNT - 1),
    # Assignees
    "AsgNames": L.abs_range(L.ASSIGNEES, 1, L.GRID_FIRST_DATA_ROW, 1, _LAST_ASG),
    "AsgProf": L.abs_range(L.ASSIGNEES, 2, L.GRID_FIRST_DATA_ROW, 2, _LAST_ASG),
    # Capacity
    "CapNames": L.abs_range(L.CAPACITY, 1, L.GRID_FIRST_DATA_ROW, 1, _LAST_ASG),
    "CapWeeks": L.abs_range(L.CAPACITY, L.GRID_FIRST_WEEK_COL, 1, _LAST_WEEK_COL, 1),
    "CapGrid": L.abs_range(L.CAPACITY, L.GRID_FIRST_WEEK_COL, L.GRID_FIRST_DATA_ROW,
                           _LAST_WEEK_COL, _LAST_ASG),
    "CapTotals": L.abs_range(L.CAPACITY, _LAST_WEEK_COL + 1, L.GRID_FIRST_DATA_ROW,
                             _LAST_WEEK_COL + 1, _LAST_ASG),
    # Equipment
    "EqpNames": L.abs_range(L.EQUIPMENT, 1, L.GRID_FIRST_DATA_ROW, 1, _LAST_EQP),
    "EqpWeeks": L.abs_range(L.EQUIPMENT, L.GRID_FIRST_WEEK_COL, 1, _LAST_WEEK_COL, 1),
    "EqpGrid": L.abs_range(L.EQUIPMENT, L.GRID_FIRST_WEEK_COL, L.GRID_FIRST_DATA_ROW,
                           _LAST_WEEK_COL, _LAST_EQP),
    # Holidays
    "HolDates": L.abs_range(L.HOLIDAYS, 1, L.GRID_FIRST_DATA_ROW, 1, _LAST_HOL),
    # Tasks
    "TaskIDs": L.abs_range(L.TASKS, L.T_ID, L.TASK_FIRST_ROW, L.T_ID, _LAST_TASK),
    "TaskPrio": L.abs_range(L.TASKS, L.T_PRIORITY, L.TASK_FIRST_ROW, L.T_PRIORITY, _LAST_TASK),
    "TaskEquip": L.abs_range(L.TASKS, L.T_EQUIPMENT, L.TASK_FIRST_ROW, L.T_EQUIPMENT, _LAST_TASK),
    "TaskDefAsg": L.abs_range(L.TASKS, L.T_DEF_ASSIGNEE, L.TASK_FIRST_ROW,
                              L.T_DEF_ASSIGNEE, _LAST_TASK),
    "TaskStartWW": L.abs_range(L.TASKS, L.T_START_WW, L.TASK_FIRST_ROW, L.T_START_WW, _LAST_TASK),
    # Sub-tasks
    "SubParent": L.abs_range(L.SUBTASKS, L.S_PARENT, L.SUB_FIRST_ROW, L.S_PARENT, _LAST_SUB),
    "SubID": L.abs_range(L.SUBTASKS, L.S_ID, L.SUB_FIRST_ROW, L.S_ID, _LAST_SUB),
    "SubName": L.abs_range(L.SUBTASKS, L.S_NAME, L.SUB_FIRST_ROW, L.S_NAME, _LAST_SUB),
    "SubAsgEff": L.abs_range(L.SUBTASKS, L.S_EFF_ASSIGNEE, L.SUB_FIRST_ROW,
                             L.S_EFF_ASSIGNEE, _LAST_SUB),
    "SubEffort": L.abs_range(L.SUBTASKS, L.S_EFFORT, L.SUB_FIRST_ROW, L.S_EFFORT, _LAST_SUB),
    "SubKey": L.abs_range(L.SUBTASKS, L.S_KEY, L.SUB_FIRST_ROW, L.S_KEY, _LAST_SUB),
    "SubRank": L.abs_range(L.SUBTASKS, L.S_RANK, L.SUB_FIRST_ROW, L.S_RANK, _LAST_SUB),
    # CalcWeek
    "CwWeeks": L.abs_range(L.CALC_WEEK, L.CW_FIRST_WEEK_COL, L.CW_HDR_ROW,
                           _CW_LAST_WEEK_COL, L.CW_HDR_ROW),
    "CwParent": L.abs_range(L.CALC_WEEK, L.CW_PARENT, L.CW_FIRST_ROW, L.CW_PARENT, _CW_LAST_ROW),
    "CwAssignee": L.abs_range(L.CALC_WEEK, L.CW_ASSIGNEE, L.CW_FIRST_ROW,
                              L.CW_ASSIGNEE, _CW_LAST_ROW),
    "CwGrid": L.abs_range(L.CALC_WEEK, L.CW_FIRST_WEEK_COL, L.CW_FIRST_ROW,
                          _CW_LAST_WEEK_COL, _CW_LAST_ROW),
    "TwIDs": L.abs_range(L.CALC_WEEK, 1, L.CW_TASKWEEK_FIRST, 1, _TW_LAST_ROW),
    "TwGrid": L.abs_range(L.CALC_WEEK, L.CW_FIRST_WEEK_COL, L.CW_TASKWEEK_FIRST,
                          _CW_LAST_WEEK_COL, _TW_LAST_ROW),
    "AwGrid": L.abs_range(L.CALC_WEEK, L.CW_FIRST_WEEK_COL, L.CW_AWWEEK_FIRST,
                          _CW_LAST_WEEK_COL, _AW_LAST_ROW),
    "EqDemand": L.abs_range(L.CALC_WEEK, L.CW_FIRST_WEEK_COL, L.CW_EQPWEEK_FIRST,
                            _CW_LAST_WEEK_COL, _EQD_LAST_ROW),
    # Gantt-Deep window inputs
    "DeepStartWeek": f"{L.sheet_ref(L.GANTT_DEEP)}!${L.GD_WINDOW_START_CELL[0]}${L.GD_WINDOW_START_CELL[1:]}",
    "DeepWeeks": f"{L.sheet_ref(L.GANTT_DEEP)}!${L.GD_WINDOW_WEEKS_CELL[0]}${L.GD_WINDOW_WEEKS_CELL[1:]}",
}

# Convenience: last row/col constants other modules need.
LAST_ASG_ROW = _LAST_ASG
LAST_EQP_ROW = _LAST_EQP
LAST_TASK_ROW = _LAST_TASK
LAST_SUB_ROW = _LAST_SUB
LAST_HOL_ROW = _LAST_HOL
LAST_WEEK_COL = _LAST_WEEK_COL
CW_LAST_WEEK_COL = _CW_LAST_WEEK_COL
CW_LAST_ROW = _CW_LAST_ROW
TW_LAST_ROW = _TW_LAST_ROW
AW_LAST_ROW = _AW_LAST_ROW
EQD_LAST_ROW = _EQD_LAST_ROW


def register(wb) -> None:
    for name, ref in NAMES.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))
