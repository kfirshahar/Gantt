"""The hidden scheduling grids.

Both grids are re-ordered projections of Sub-Tasks: grid row k holds whichever
sub-task currently has rank k. That ordering is what makes the "capacity already
claimed by higher-ranked work" term a SUMIF over rows *above only*, which is the
only form Excel will accept. A cumulative that scanned the whole column would
include its own cell, and referencing rows below is equally circular because
those rows reference this one.
"""

from . import calendar_utils as C, layout as L, names as N, styles as S


def _pull(rank_cell: str, source: str) -> str:
    """Value of `source` for the sub-task holding this rank."""
    return f"INDEX({source},MATCH({rank_cell},SubRank,0))"


def _row_identity(ws, r: int, rank_col: int, first_row: int) -> None:
    """Columns A..F: which sub-task sits in this grid row, and its facts."""
    rank = f"$A{r}"
    ws.cell(row=r, column=rank_col, value=r - first_row + 1)
    ws.cell(row=r, column=rank_col + 1, value=f'=IFERROR({_pull(rank, "SubID")},"")')
    ws.cell(row=r, column=rank_col + 2,
            value=f'=IF($B{r}="","",{_pull(rank, "SubAsgEff")})')
    ws.cell(row=r, column=rank_col + 5,
            value=f'=IF($B{r}="","",{_pull(rank, "SubParent")})')
    ws.cell(row=r, column=rank_col + 4,
            value=f'=IF($B{r}="",0,IFERROR(INDEX(TaskStartWW,MATCH($F{r},TaskIDs,0)),0))')


def _effort_expr(r: int) -> str:
    return f'IFERROR({_pull(f"$A{r}", "SubEffort")}*1,0)'


# --- CalcWeek --------------------------------------------------------------

def build_calc_week(ws) -> None:
    hdr, first, last = L.CW_HDR_ROW, L.CW_FIRST_ROW, N.CW_LAST_ROW
    S.header_row(ws, hdr, 1,
                 ["Rank", "Sub ID", "Assignee", "Effort", "Start WW", "Parent"])
    for i in range(L.HORIZON_WEEKS):
        c = ws.cell(row=hdr, column=L.CW_FIRST_WEEK_COL + i, value=f"=CfgStartWeek+{i}")
        c.number_format = '"WW"0'
        c.fill = S.FILL_INPUT_HDR
        c.font = S.FONT_HDR

    for r in range(first, last + 1):
        _row_identity(ws, r, L.CW_RANK, first)
        ws.cell(row=r, column=L.CW_EFFORT,
                value=f'=IF($B{r}="",0,{_effort_expr(r)})').number_format = "0.00"

        for i in range(L.HORIZON_WEEKS):
            wc = L.col(L.CW_FIRST_WEEK_COL + i)
            prior = "0" if i == 0 else f"SUM($G{r}:{L.col(L.CW_FIRST_WEEK_COL + i - 1)}{r})"
            cap = (f"IFERROR(INDEX(CapGrid,MATCH($C{r},CapNames,0),"
                   f"MATCH({wc}${hdr},CapWeeks,0)),0)")
            claimed = ("0" if r == first else
                       f"SUMIF($C${first}:$C{r - 1},$C{r},{wc}${first}:{wc}{r - 1})")
            ws.cell(row=r, column=L.CW_FIRST_WEEK_COL + i, value=(
                f'=IF($B{r}="",0,IF({wc}${hdr}<$E{r},0,'
                f'MAX(0,MIN($D{r}-{prior},{cap}-{claimed}))))')).number_format = "0.00"

    _task_rollup(ws)
    _assignee_rollup(ws)
    _equipment_demand(ws)
    ws.sheet_state = "hidden"


def _task_rollup(ws) -> None:
    """One row per task: that task's allocated days in each week."""
    hdr, first = L.CW_TASKWEEK_HDR, L.CW_TASKWEEK_FIRST
    S.header_row(ws, hdr, 1, ["Task ID (rollup)"])
    for i in range(L.HORIZON_WEEKS):
        ws.cell(row=hdr, column=L.CW_FIRST_WEEK_COL + i, value=f"=CfgStartWeek+{i}")

    for i in range(L.MAX_TASKS):
        r = first + i
        task_row = L.TASK_FIRST_ROW + i
        ws.cell(row=r, column=1, value=(
            f'=IF({L.sheet_ref(L.TASKS)}!A{task_row}="","",'
            f'{L.sheet_ref(L.TASKS)}!A{task_row})'))
        for w in range(L.HORIZON_WEEKS):
            wc = L.col(L.CW_FIRST_WEEK_COL + w)
            ws.cell(row=r, column=L.CW_FIRST_WEEK_COL + w, value=(
                f'=IF($A{r}="",0,SUMIF(CwParent,$A{r},'
                f'{wc}${L.CW_FIRST_ROW}:{wc}${N.CW_LAST_ROW}))')).number_format = "0.00"


def _assignee_rollup(ws) -> None:
    """One row per assignee: days that person is loaded with each week."""
    hdr, first = L.CW_AWWEEK_HDR, L.CW_AWWEEK_FIRST
    S.header_row(ws, hdr, 1, ["Assignee (load)"])
    for i in range(L.HORIZON_WEEKS):
        ws.cell(row=hdr, column=L.CW_FIRST_WEEK_COL + i, value=f"=CfgStartWeek+{i}")

    for i in range(L.MAX_ASSIGNEES):
        r = first + i
        src = L.GRID_FIRST_DATA_ROW + i
        ws.cell(row=r, column=1, value=(
            f'=IF({L.sheet_ref(L.ASSIGNEES)}!A{src}="","",'
            f'{L.sheet_ref(L.ASSIGNEES)}!A{src})'))
        for w in range(L.HORIZON_WEEKS):
            wc = L.col(L.CW_FIRST_WEEK_COL + w)
            ws.cell(row=r, column=L.CW_FIRST_WEEK_COL + w, value=(
                f'=IF($A{r}="",0,SUMIF(CwAssignee,$A{r},'
                f'{wc}${L.CW_FIRST_ROW}:{wc}${N.CW_LAST_ROW}))')).number_format = "0.00"


def _equipment_demand(ws) -> None:
    """Units of each equipment type demanded per week.

    Demand is counted per *parent task*, since a task holds one unit while it is
    in progress regardless of how many of its sub-tasks are active.
    """
    hdr, first = L.CW_EQPWEEK_HDR, L.CW_EQPWEEK_FIRST
    S.header_row(ws, hdr, 1, ["Equipment (demand)"])
    for i in range(L.HORIZON_WEEKS):
        ws.cell(row=hdr, column=L.CW_FIRST_WEEK_COL + i, value=f"=CfgStartWeek+{i}")

    for i in range(L.MAX_EQUIPMENT):
        r = first + i
        eq_row = L.GRID_FIRST_DATA_ROW + i
        ws.cell(row=r, column=1, value=(
            f'=IF({L.sheet_ref(L.EQUIPMENT)}!A{eq_row}="","",'
            f'{L.sheet_ref(L.EQUIPMENT)}!A{eq_row})'))
        for w in range(L.HORIZON_WEEKS):
            wc = L.col(L.CW_FIRST_WEEK_COL + w)
            tw_col = (f"{wc}${L.CW_TASKWEEK_FIRST}:{wc}${N.TW_LAST_ROW}")
            ws.cell(row=r, column=L.CW_FIRST_WEEK_COL + w, value=(
                f'=IF($A{r}="",0,SUMPRODUCT((TaskEquip=$A{r})*({tw_col}>0)))'))


# --- CalcDay ---------------------------------------------------------------

def build_calc_day(ws) -> None:
    hdr, first, last = L.CD_HDR_ROW, L.CD_FIRST_ROW, L.CD_FIRST_ROW + L.MAX_SUBTASKS - 1
    S.header_row(ws, hdr, 1,
                 ["Rank", "Sub ID", "Assignee", "Remaining", "Start WW", "Parent"])
    _day_headers(ws)

    for r in range(first, last + 1):
        _row_identity(ws, r, L.CD_RANK, first)
        cw_row = L.CW_FIRST_ROW + (r - first)
        cw_span = (f"{L.sheet_ref(L.CALC_WEEK)}!"
                   f"${L.col(L.CW_FIRST_WEEK_COL)}${cw_row}:${L.col(N.CW_LAST_WEEK_COL)}${cw_row}")
        # Effort already burned before the window opens, so a window opened
        # mid-project shows correct residual work.
        ws.cell(row=r, column=L.CD_REMAINING, value=(
            f'=IF($B{r}="",0,MAX(0,{_effort_expr(r)}'
            f'-SUMPRODUCT((CwWeeks<DeepStartWeek)*{cw_span})))')).number_format = "0.00"

        for i in range(L.DEEP_DAY_COLS):
            dc = L.col(L.CD_FIRST_DAY_COL + i)
            prior = ("0" if i == 0 else
                     f"SUM($G{r}:{L.col(L.CD_FIRST_DAY_COL + i - 1)}{r})")
            claimed = ("0" if r == first else
                       f"SUMIF($C${first}:$C{r - 1},$C{r},{dc}${first}:{dc}{r - 1})")
            week_cap = (f"IFERROR(INDEX(CapGrid,MATCH($C{r},CapNames,0),"
                        f"MATCH({dc}${L.CD_WEEKNO_ROW},CapWeeks,0)),0)")
            day_cap = (f'IF(OR({dc}${L.CD_INWINDOW_ROW}=0,{dc}${L.CD_HOLIDAY_ROW}=1,'
                       f'{dc}${L.CD_WORKDAYS_ROW}<=0),0,'
                       f'{week_cap}/{dc}${L.CD_WORKDAYS_ROW})')
            ws.cell(row=r, column=L.CD_FIRST_DAY_COL + i, value=(
                f'=IF($B{r}="",0,IF({dc}${L.CD_INWINDOW_ROW}=0,0,'
                f'IF({dc}${L.CD_WEEKNO_ROW}<$E{r},0,'
                f'MAX(0,MIN($D{r}-{prior},{day_cap}-{claimed})))))')).number_format = "0.00"

    ws.sheet_state = "hidden"


def _day_headers(ws) -> None:
    """Five helper rows describing each day column: week number, date, holiday
    flag, that week's working-day count, and whether it is inside the window."""
    sunday = C.excel_week_sunday_formula("CfgYear", f"{{c}}${L.CD_WEEKNO_ROW}")
    for i in range(L.DEEP_DAY_COLS):
        c = L.col(L.CD_FIRST_DAY_COL + i)
        wk, day = divmod(i, L.WORKDAYS_PER_WEEK)
        sun = sunday.format(c=c)

        ws.cell(row=L.CD_WEEKNO_ROW, column=L.CD_FIRST_DAY_COL + i,
                value=f"=DeepStartWeek+{wk}")
        d = ws.cell(row=L.CD_DATE_ROW, column=L.CD_FIRST_DAY_COL + i,
                    value=f"={sun}+{day}")
        d.number_format = "yyyy-mm-dd"
        ws.cell(row=L.CD_HOLIDAY_ROW, column=L.CD_FIRST_DAY_COL + i,
                value=f"=IF(COUNTIF(HolDates,{c}${L.CD_DATE_ROW})>0,1,0)")
        ws.cell(row=L.CD_WORKDAYS_ROW, column=L.CD_FIRST_DAY_COL + i, value=(
            f'={L.WORKDAYS_PER_WEEK}-COUNTIFS(HolDates,">="&{sun},'
            f'HolDates,"<="&{sun}+{L.WORKDAYS_PER_WEEK - 1})'))
        ws.cell(row=L.CD_INWINDOW_ROW, column=L.CD_FIRST_DAY_COL + i, value=(
            f'=IF(AND({wk}<DeepWeeks,{c}${L.CD_WEEKNO_ROW}>=CfgStartWeek,'
            f'{c}${L.CD_WEEKNO_ROW}<=CfgStartWeek+CfgHorizon-1),1,0)'))

    for row, label in [(L.CD_WEEKNO_ROW, "Week"), (L.CD_DATE_ROW, "Date"),
                       (L.CD_HOLIDAY_ROW, "Holiday?"),
                       (L.CD_WORKDAYS_ROW, "Work days in week"),
                       (L.CD_INWINDOW_ROW, "In window?")]:
        ws.cell(row=row, column=1, value=label).font = S.FONT_NOTE
