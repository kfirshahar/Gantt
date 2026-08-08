"""The Tasks and Sub-Tasks input tabs, including the derived columns."""

from openpyxl.formatting.rule import Rule
from openpyxl.styles import Border, Font, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.datavalidation import DataValidation

from . import demo, layout as L, names as N, styles as S


def _list_validation(ws, name: str, col: int, first_row: int, last_row: int,
                     allow_blank: bool = True) -> None:
    dv = DataValidation(type="list", formula1=f"={name}", allow_blank=allow_blank,
                        showDropDown=False)
    dv.error = f"Pick a value from the {name} list."
    dv.errorTitle = "Not a known value"
    ws.add_data_validation(dv)
    dv.add(f"{L.col(col)}{first_row}:{L.col(col)}{last_row}")


def _check_formatting(ws, col: int, first_row: int, last_row: int) -> None:
    rng = f"{L.col(col)}{first_row}:{L.col(col)}{last_row}"
    ws.conditional_formatting.add(rng, Rule(
        type="containsText", operator="containsText", text="⚠",
        dxf=DifferentialStyle(fill=S.FILL_WARN),
        formula=[f'NOT(ISERROR(SEARCH("⚠",{L.col(col)}{first_row})))']))
    ws.conditional_formatting.add(rng, Rule(
        type="containsText", operator="containsText", text="ok",
        dxf=DifferentialStyle(fill=S.FILL_OK),
        formula=[f'NOT(ISERROR(SEARCH("ok",{L.col(col)}{first_row})))']))


# --- Tasks -----------------------------------------------------------------

def build_tasks(ws) -> None:
    S.header_row(ws, 1, 1, [L.TASK_HEADERS[c] for c in sorted(L.TASK_HEADERS)])
    for c in L.TASK_DERIVED_COLS:
        cell = ws.cell(row=1, column=c)
        cell.fill = S.FILL_OUTPUT_HDR

    for i, task in enumerate(demo.TASKS):
        r = L.TASK_FIRST_ROW + i
        tid, name, category, priority, complexity, equipment, assignee, start_ww, _ = task
        ws.cell(row=r, column=L.T_ID, value=tid)
        ws.cell(row=r, column=L.T_NAME, value=name)
        ws.cell(row=r, column=L.T_CATEGORY, value=category)
        ws.cell(row=r, column=L.T_PRIORITY, value=priority)
        ws.cell(row=r, column=L.T_COMPLEXITY, value=complexity)
        ws.cell(row=r, column=L.T_EQUIPMENT, value=equipment)
        ws.cell(row=r, column=L.T_DEF_ASSIGNEE, value=assignee)
        ws.cell(row=r, column=L.T_START_WW, value=start_ww)

    for r in range(L.TASK_FIRST_ROW, N.LAST_TASK_ROW + 1):
        tw_row = L.CW_TASKWEEK_FIRST + (r - L.TASK_FIRST_ROW)
        span = (f"{L.sheet_ref(L.CALC_WEEK)}!"
                f"${L.col(L.CW_FIRST_WEEK_COL)}${tw_row}:${L.col(N.CW_LAST_WEEK_COL)}${tw_row}")
        effort = f"SUMIF(SubParent,$A{r},SubEffort)"

        ws.cell(row=r, column=L.T_N_SUBS,
                value=f'=IF($A{r}="","",COUNTIF(SubParent,$A{r}))')
        ws.cell(row=r, column=L.T_EFFORT,
                value=f'=IF($A{r}="","",ROUND({effort},2))').number_format = "0.00"
        ws.cell(row=r, column=L.T_CALC_START,
                value=(f'=IF($A{r}="","",IF(SUM({span})=0,"",'
                       f'SUMPRODUCT(MIN(({span}>0)*CwWeeks+({span}<=0)*9999))))')
                ).number_format = '"WW"0'
        ws.cell(row=r, column=L.T_CALC_END,
                value=(f'=IF($A{r}="","",IF(SUM({span})=0,"",'
                       f'SUMPRODUCT(MAX(({span}>0)*CwWeeks))))')
                ).number_format = '"WW"0'
        ws.cell(row=r, column=L.T_CHECK, value=(
            f'=IF($A{r}="","",'
            f'IF(COUNTIF(SubParent,$A{r})=0,"⚠ no sub-tasks",'
            f'IF(OR($H{r}<CfgStartWeek,$H{r}>CfgStartWeek+CfgHorizon-1),"⚠ start outside horizon",'
            f'IF(SUM({span})=0,"⚠ not scheduled",'
            f'IF(ROUND(SUM({span}),4)<ROUND({effort},4),"⚠ overruns horizon","ok")))))'))

        for c in L.TASK_DERIVED_COLS:
            S.mark_derived(ws.cell(row=r, column=c))
        for c in range(1, L.T_CHECK + 1):
            ws.cell(row=r, column=c).border = S.BORDER_ALL
        ws.cell(row=r, column=L.T_START_WW).number_format = '"WW"0'

    _list_validation(ws, "PrioNames", L.T_PRIORITY, L.TASK_FIRST_ROW, N.LAST_TASK_ROW)
    _list_validation(ws, "CplxNames", L.T_COMPLEXITY, L.TASK_FIRST_ROW, N.LAST_TASK_ROW)
    _list_validation(ws, "EqpNames", L.T_EQUIPMENT, L.TASK_FIRST_ROW, N.LAST_TASK_ROW)
    _list_validation(ws, "AsgNames", L.T_DEF_ASSIGNEE, L.TASK_FIRST_ROW, N.LAST_TASK_ROW)
    _check_formatting(ws, L.T_CHECK, L.TASK_FIRST_ROW, N.LAST_TASK_ROW)

    S.widths(ws, {"A": 10, "B": 24, "C": 12, "D": 9, "E": 12, "F": 15, "G": 17,
                  "H": 17, "I": 12, "J": 13, "K": 10, "L": 10, "M": 22})
    ws.freeze_panes = "C2"
    r = N.LAST_TASK_ROW + 2
    ws.cell(row=r, column=1,
            value="Grey italic columns are computed — do not type in them. "
                  "Effort is the sum of the task's sub-tasks; task complexity is metadata.").font = S.FONT_NOTE


# --- Sub-Tasks -------------------------------------------------------------

def build_subtasks(ws) -> None:
    S.header_row(ws, 1, 1, [L.SUB_HEADERS[c] for c in sorted(L.SUB_HEADERS)])
    for c in L.SUB_DERIVED_COLS:
        ws.cell(row=1, column=c).fill = S.FILL_OUTPUT_HDR

    for i, (parent, name, complexity, override) in enumerate(demo.subtasks()):
        r = L.SUB_FIRST_ROW + i
        ws.cell(row=r, column=L.S_PARENT, value=parent)
        ws.cell(row=r, column=L.S_NAME, value=name)
        ws.cell(row=r, column=L.S_COMPLEXITY, value=complexity)
        if override:
            ws.cell(row=r, column=L.S_ASSIGNEE, value=override)

    for r in range(L.SUB_FIRST_ROW, N.LAST_SUB_ROW + 1):
        idx = f"COUNTIF($A${L.SUB_FIRST_ROW}:$A{r},$A{r})"
        parent_row = f"MATCH($A{r},TaskIDs,0)"
        ws.cell(row=r, column=L.S_ID,
                value=f'=IF($A{r}="","",$A{r}&"."&TEXT({idx},"00"))')
        ws.cell(row=r, column=L.S_EFF_ASSIGNEE, value=(
            f'=IF($A{r}="","",IF($E{r}<>"",$E{r},'
            f'IFERROR(INDEX(TaskDefAsg,{parent_row}),"")))'))
        ws.cell(row=r, column=L.S_EFFORT, value=(
            f'=IF(OR($A{r}="",$D{r}="",$F{r}=""),"",'
            f'IFERROR(INDEX(CplxDays,MATCH($D{r},CplxNames,0))'
            f'/INDEX(AsgProf,MATCH($F{r},AsgNames,0)),""))')).number_format = "0.00"
        ws.cell(row=r, column=L.S_KEY, value=(
            f'=IF($G{r}="","",IFERROR('
            f'INDEX(PrioRank,MATCH(INDEX(TaskPrio,{parent_row}),PrioNames,0))*1000000000'
            f'+INDEX(TaskStartWW,{parent_row})*1000000'
            f'+{parent_row}*1000+{idx},""))')).number_format = "0"
        ws.cell(row=r, column=L.S_RANK,
                value=f'=IF($H{r}="","",RANK($H{r},SubKey,1))')
        ws.cell(row=r, column=L.S_CHECK, value=(
            f'=IF($A{r}="","",'
            f'IF(ISNA({parent_row}),"⚠ unknown parent",'
            f'IF($D{r}="","⚠ no complexity",'
            f'IF($F{r}="","⚠ no assignee",'
            f'IF($G{r}="","⚠ check proficiency","ok")))))'))

        for c in L.SUB_DERIVED_COLS:
            S.mark_derived(ws.cell(row=r, column=c))
        for c in range(1, L.S_CHECK + 1):
            ws.cell(row=r, column=c).border = S.BORDER_ALL

    _list_validation(ws, "TaskIDs", L.S_PARENT, L.SUB_FIRST_ROW, N.LAST_SUB_ROW)
    _list_validation(ws, "CplxNames", L.S_COMPLEXITY, L.SUB_FIRST_ROW, N.LAST_SUB_ROW)
    _list_validation(ws, "AsgNames", L.S_ASSIGNEE, L.SUB_FIRST_ROW, N.LAST_SUB_ROW)
    _check_formatting(ws, L.S_CHECK, L.SUB_FIRST_ROW, N.LAST_SUB_ROW)
    _parent_looks_merged(ws)

    S.widths(ws, {"A": 12, "B": 12, "C": 26, "D": 12, "E": 22, "F": 18,
                  "G": 12, "H": 16, "I": 8, "J": 22})
    ws.freeze_panes = "C2"
    r = N.LAST_SUB_ROW + 2
    ws.cell(row=r, column=1,
            value="Leave Assignee blank to inherit the parent task's default assignee. "
                  "Repeated parent IDs are hidden for readability — the value is still "
                  "in every cell, so sorting and filtering keep working.").font = S.FONT_NOTE


def _parent_looks_merged(ws) -> None:
    """Render repeated parent IDs as one block without actually merging.

    A real merge would store the value only in the top-left cell and would break
    sort and AutoFilter on a table the user edits. Instead every row keeps its
    real parent ID and the repeats are drawn in white.
    """
    first, last = L.SUB_FIRST_ROW, N.LAST_SUB_ROW
    parent_col = L.col(L.S_PARENT)

    repeats = f"{parent_col}{first}:{parent_col}{last}"
    ws.conditional_formatting.add(repeats, Rule(
        type="expression",
        dxf=DifferentialStyle(font=Font(color="FFFFFF")),
        formula=[f'AND(${parent_col}{first}<>"",${parent_col}{first}=${parent_col}{first - 1})']))

    whole_row = f"A{first}:{L.col(L.S_CHECK)}{last}"
    ws.conditional_formatting.add(whole_row, Rule(
        type="expression",
        dxf=DifferentialStyle(border=Border(top=Side(style="medium", color=S.INK))),
        formula=[f'AND(${parent_col}{first}<>"",${parent_col}{first}<>${parent_col}{first - 1})']))
