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
    # Column E is the position work may first be planned into: the later of the
    # task's earliest start and the current week. Nothing is ever scheduled into
    # the past, and an earliest start that has already gone by simply means "as
    # soon as possible" rather than an error — which is why the horizon no
    # longer has to move as the project runs.
    #
    # An unknown week still yields 9999, so a start naming no column in the grid
    # never schedules and the Tasks check reports it.
    ws.cell(row=r, column=rank_col + 4, value=(
        f'=IF($B{r}="",9999,IFERROR(MAX(current_pos,MATCH('
        f'INDEX(TaskStartWW,MATCH($F{r},TaskIDs,0)),CwWeeks,0)),9999))'
        .replace("current_pos", _current_pos())))


def _current_pos() -> str:
    """Grid position of the current week, or the first column if it is behind
    the project start — planning cannot begin before the plan does."""
    return "IFERROR(MATCH(CfgCurrentWeek,CwWeeks,0),1)"


def _effort_expr(r: int) -> str:
    """The budget scheduling draws against: what's left, not what the task took.

    Done work claims nothing and part-done work claims only its balance, so
    finished sub-tasks stop competing for capacity the moment they're marked.
    """
    return f'IFERROR({_pull(f"$A{r}", "SubRemaining")}*1,0)'


# --- CalcWeek --------------------------------------------------------------

def build_calc_week(ws) -> None:
    hdr, first, last = L.CW_HDR_ROW, L.CW_FIRST_ROW, N.CW_LAST_ROW
    S.header_row(ws, hdr, 1,
                 ["Rank", "Sub ID", "Assignee", "Remaining", "Start WW", "Parent"])
    _week_headers(ws)

    for r in range(first, last + 1):
        _row_identity(ws, r, L.CW_RANK, first)
        ws.cell(row=r, column=L.CW_EFFORT,
                value=f'=IF($B{r}="",0,{_effort_expr(r)})').number_format = "0.00"

        for i in range(L.WEEK_COLS):
            wc = L.col(L.CW_FIRST_WEEK_COL + i)
            pos = i + 1
            prior = "0" if i == 0 else f"SUM($G{r}:{L.col(L.CW_FIRST_WEEK_COL + i - 1)}{r})"
            # Column position is known at generation time, so the capacity
            # lookup indexes straight into the grid instead of matching on a
            # header. Nothing here depends on the week's calendar number.
            cap = (f"IFERROR(INDEX(CapGrid,MATCH($C{r},CapNames,0),{pos}),0)"
                   f"*INDEX(CwFactor,{pos})")
            claimed = ("0" if r == first else
                       f"SUMIF($C${first}:$C{r - 1},$C{r},{wc}${first}:{wc}{r - 1})")
            ws.cell(row=r, column=L.CW_FIRST_WEEK_COL + i, value=(
                f'=IF($B{r}="",0,IF({wc}${L.CW_ACTIVE_ROW}=0,0,'
                f'IF({pos}<$E{r},0,'
                f'MAX(0,MIN($D{r}-{prior},{cap}-{claimed})))))')).number_format = "0.00"

    _task_rollup(ws)
    _assignee_rollup(ws)
    _equipment_demand(ws)
    _actuals(ws)
    ws.sheet_state = "hidden"


def _week_headers(ws) -> None:
    """Seven helper rows describing each week column.

    Position drives the machinery; the calendar week and year are derived from
    the column's actual Sunday, so a plan crossing New Year labels itself
    correctly instead of counting on to WW58.
    """
    labels = [(L.CW_POS_ROW, "Position"), (L.CW_ABS_ROW, "Absolute"),
              (L.CW_SUN_ROW, "Sunday"), (L.CW_YEAR_ROW, "Year"),
              (L.CW_WEEK_ROW, "Calendar WW"), (L.CW_WORKDAYS_ROW, "Work days"),
              (L.CW_FACTOR_ROW, "Holiday factor"), (L.CW_ACTIVE_ROW, "Active?"),
              (L.CW_LABEL_ROW, "Label")]
    for row, text in labels:
        ws.cell(row=row, column=1, value=text).font = S.FONT_NOTE

    for i in range(L.WEEK_COLS):
        col = L.CW_FIRST_WEEK_COL + i
        wc = L.col(col)
        pos, abs_, sun = f"{wc}${L.CW_POS_ROW}", f"{wc}${L.CW_ABS_ROW}", f"{wc}${L.CW_SUN_ROW}"
        year, week = f"{wc}${L.CW_YEAR_ROW}", f"{wc}${L.CW_WEEK_ROW}"

        ws.cell(row=L.CW_POS_ROW, column=col, value=i + 1)
        ws.cell(row=L.CW_ABS_ROW, column=col, value=f"=CfgStartWeek+{pos}-1")
        d = ws.cell(row=L.CW_SUN_ROW, column=col,
                    value=f"={C.excel_week_sunday_formula('CfgYear', abs_)}")
        d.number_format = "yyyy-mm-dd"
        ws.cell(row=L.CW_YEAR_ROW, column=col,
                value=f"={C.excel_calendar_year_formula(sun)}")
        ws.cell(row=L.CW_WEEK_ROW, column=col,
                value=f"={C.excel_calendar_week_formula(sun, year)}")
        # Company holidays shorten the week, and capacity is reduced in
        # proportion. Entered capacity is what someone has in a normal five-day
        # week; this is the only place holidays touch the weekly totals.
        ws.cell(row=L.CW_WORKDAYS_ROW, column=col, value=(
            f'={L.WORKDAYS_PER_WEEK}-COUNTIFS(HolDates,">="&{sun},'
            f'HolDates,"<="&{sun}+{L.WORKDAYS_PER_WEEK - 1})'))
        ws.cell(row=L.CW_FACTOR_ROW, column=col,
                value=f"={wc}${L.CW_WORKDAYS_ROW}/{L.WORKDAYS_PER_WEEK}")
        # Weeks past the configured horizon are switched off here, which is what
        # lets Config's Horizon cell lengthen or shorten the plan without the
        # workbook being regenerated.
        ws.cell(row=L.CW_ACTIVE_ROW, column=col, value=f"=IF({pos}<=CfgHorizon,1,0)")
        ws.cell(row=L.CW_LABEL_ROW, column=col,
                value=f"={C.excel_week_label_formula(week, year, 'CfgYear')}")

        h = ws.cell(row=L.CW_HDR_ROW, column=col, value=f"={wc}${L.CW_LABEL_ROW}")
        h.fill = S.FILL_INPUT_HDR
        h.font = S.FONT_HDR


def _actuals(ws) -> None:
    """Consumed effort laid out across the weeks it was actually spent in.

    History, not a schedule. It is spread evenly across the recorded span, which
    is approximate on purpose — the exact shape of work already done is not
    worth modelling, and pretending otherwise would invite the number to be read
    as a measurement.

    Clamped to end before the current week so it can never collide with the
    plan: everything behind the pointer is actual, everything from it forward is
    planned, and the two are simply added when a view wants a total.
    """
    hdr, first = L.CW_ACTUAL_HDR, L.CW_ACTUAL_FIRST
    S.header_row(ws, hdr, 1, ["Rank", "Sub ID", "Consumed", "From", "To"])
    for i in range(L.WEEK_COLS):
        ws.cell(row=hdr, column=L.CW_FIRST_WEEK_COL + i, value=f"=INDEX(CwLabel,{i + 1})")

    current = _current_pos()
    for i in range(L.MAX_SUBTASKS):
        r = first + i
        rank = f"$A{r}"
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=f'=IFERROR({_pull(rank, "SubID")},"")')
        ws.cell(row=r, column=3,
                value=f'=IF($B{r}="",0,IFERROR({_pull(rank, "SubConsumed")}*1,0))')
        # A missing end means still running, so the band reaches the last
        # finished week. Both are clamped so a band never crosses into the plan.
        ws.cell(row=r, column=4, value=(
            f'=IF($B{r}="",0,IFERROR(MATCH({_pull(rank, "SubActStart")},CwWeeks,0),0))'))
        ws.cell(row=r, column=5, value=(
            f'=IF($D{r}=0,0,MIN({current}-1,'
            f'IFERROR(MATCH({_pull(rank, "SubActEnd")},CwWeeks,0),{current}-1)))'))

        for w in range(L.WEEK_COLS):
            pos = w + 1
            ws.cell(row=r, column=L.CW_FIRST_WEEK_COL + w, value=(
                f'=IF(OR($D{r}=0,$E{r}<$D{r},$C{r}=0),0,'
                f'IF(AND({pos}>=$D{r},{pos}<=$E{r}),'
                f'$C{r}/($E{r}-$D{r}+1),0))')).number_format = "0.00"


def _task_rollup(ws) -> None:
    """One row per task: that task's allocated days in each week."""
    hdr, first = L.CW_TASKWEEK_HDR, L.CW_TASKWEEK_FIRST
    S.header_row(ws, hdr, 1, ["Task ID (rollup)"])
    for i in range(L.WEEK_COLS):
        ws.cell(row=hdr, column=L.CW_FIRST_WEEK_COL + i, value=f"=INDEX(CwLabel,{i + 1})")

    for i in range(L.MAX_TASKS):
        r = first + i
        task_row = L.TASK_FIRST_ROW + i
        ws.cell(row=r, column=1, value=(
            f'=IF({L.sheet_ref(L.TASKS)}!A{task_row}="","",'
            f'{L.sheet_ref(L.TASKS)}!A{task_row})'))
        for w in range(L.WEEK_COLS):
            wc = L.col(L.CW_FIRST_WEEK_COL + w)
            ws.cell(row=r, column=L.CW_FIRST_WEEK_COL + w, value=(
                f'=IF($A{r}="",0,SUMIF(CwParent,$A{r},'
                f'{wc}${L.CW_FIRST_ROW}:{wc}${N.CW_LAST_ROW})'
                f'+SUMIF(CwParent,$A{r},'
                f'{wc}${L.CW_ACTUAL_FIRST}:{wc}${N.ACT_LAST_ROW}))')).number_format = "0.00"


def _assignee_rollup(ws) -> None:
    """One row per assignee: days that person is loaded with each week."""
    hdr, first = L.CW_AWWEEK_HDR, L.CW_AWWEEK_FIRST
    S.header_row(ws, hdr, 1, ["Assignee (load)"])
    for i in range(L.WEEK_COLS):
        ws.cell(row=hdr, column=L.CW_FIRST_WEEK_COL + i, value=f"=INDEX(CwLabel,{i + 1})")

    for i in range(L.MAX_ASSIGNEES):
        r = first + i
        src = L.GRID_FIRST_DATA_ROW + i
        ws.cell(row=r, column=1, value=(
            f'=IF({L.sheet_ref(L.ASSIGNEES)}!A{src}="","",'
            f'{L.sheet_ref(L.ASSIGNEES)}!A{src})'))
        for w in range(L.WEEK_COLS):
            wc = L.col(L.CW_FIRST_WEEK_COL + w)
            ws.cell(row=r, column=L.CW_FIRST_WEEK_COL + w, value=(
                f'=IF($A{r}="",0,SUMIF(CwAssignee,$A{r},'
                f'{wc}${L.CW_FIRST_ROW}:{wc}${N.CW_LAST_ROW})'
                f'+SUMIF(CwAssignee,$A{r},'
                f'{wc}${L.CW_ACTUAL_FIRST}:{wc}${N.ACT_LAST_ROW}))')).number_format = "0.00"


def _equipment_demand(ws) -> None:
    """Units of each equipment type demanded per week.

    Demand is counted per *parent task*, since a task holds one unit while it is
    in progress regardless of how many of its sub-tasks are active.
    """
    hdr, first = L.CW_EQPWEEK_HDR, L.CW_EQPWEEK_FIRST
    S.header_row(ws, hdr, 1, ["Equipment (demand)"])
    for i in range(L.WEEK_COLS):
        ws.cell(row=hdr, column=L.CW_FIRST_WEEK_COL + i, value=f"=INDEX(CwLabel,{i + 1})")

    for i in range(L.MAX_EQUIPMENT):
        r = first + i
        eq_row = L.GRID_FIRST_DATA_ROW + i
        ws.cell(row=r, column=1, value=(
            f'=IF({L.sheet_ref(L.EQUIPMENT)}!A{eq_row}="","",'
            f'{L.sheet_ref(L.EQUIPMENT)}!A{eq_row})'))
        for w in range(L.WEEK_COLS):
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
        # mid-project shows correct residual work. Compared by position, not by
        # week number, which would misorder across a year boundary.
        ws.cell(row=r, column=L.CD_REMAINING, value=(
            f'=IF($B{r}="",0,MAX(0,{_effort_expr(r)}'
            f'-SUMPRODUCT((CwPos<{_deep_start_pos()})*{cw_span})))')).number_format = "0.00"

        for i in range(L.DEEP_DAY_COLS):
            dc = L.col(L.CD_FIRST_DAY_COL + i)
            pos = f"{dc}${L.CD_POS_ROW}"
            # Only planned days count towards what this sub-task has already
            # been given; the history columns to the left are not allocations.
            prior = ("0" if i == 0 else (
                f"SUMIF(${L.col(L.CD_FIRST_DAY_COL)}${L.CD_POS_ROW}:"
                f"${L.col(L.CD_FIRST_DAY_COL + i - 1)}${L.CD_POS_ROW},"
                f'">="&{_current_pos()},'
                f"${L.col(L.CD_FIRST_DAY_COL)}{r}:"
                f"${L.col(L.CD_FIRST_DAY_COL + i - 1)}{r})"))
            claimed = ("0" if r == first else
                       f"SUMIF($C${first}:$C{r - 1},$C{r},{dc}${first}:{dc}{r - 1})")
            # Weekly capacity is already reduced pro-rata, so dividing it by the
            # shortened week cancels out and leaves the person's ordinary daily
            # rate. Dividing by the remaining days instead would compress a full
            # week's work into fewer days and allow more than a day per day.
            rate = (f"IFERROR(INDEX(CapGrid,MATCH($C{r},CapNames,0),{pos}),0)"
                    f"/{L.WORKDAYS_PER_WEEK}")
            day_cap = (f'IF(OR({dc}${L.CD_INWINDOW_ROW}=0,{dc}${L.CD_HOLIDAY_ROW}=1),'
                       f'0,{rate})')
            # Days behind the current week show what was actually spent, taken
            # from the same weekly history the high-level view uses and divided
            # across that week's working days. Approximate, and labelled as
            # history — the point is that both views tell one story.
            actual = (f'IFERROR(INDEX(ActGrid,$A{r},{pos}),0)'
                      f'/MAX(1,{dc}${L.CD_WORKDAYS_ROW})')
            ws.cell(row=r, column=L.CD_FIRST_DAY_COL + i, value=(
                f'=IF($B{r}="",0,IF({dc}${L.CD_INWINDOW_ROW}=0,0,'
                f'IF({dc}${L.CD_HOLIDAY_ROW}=1,0,'
                f'IF({pos}<{_current_pos()},{actual},'
                f'IF({pos}<$E{r},0,'
                f'MAX(0,MIN($D{r}-{prior},{day_cap}-{claimed})))))))')
            ).number_format = "0.00"

    ws.sheet_state = "hidden"


def _deep_start_pos() -> str:
    """Position of the window's first week, from the calendar week the user typed."""
    return "IFERROR(MATCH(DeepStartWeek,CwWeeks,0),1)"


def _day_headers(ws) -> None:
    """Six helper rows describing each day column.

    The window start is a calendar week the user types; everything downstream
    works from the position it resolves to, so the window behaves the same
    either side of New Year.
    """
    for i in range(L.DEEP_DAY_COLS):
        c = L.col(L.CD_FIRST_DAY_COL + i)
        col = L.CD_FIRST_DAY_COL + i
        wk, day = divmod(i, L.WORKDAYS_PER_WEEK)
        pos = f"{c}${L.CD_POS_ROW}"
        # A position maps back to a date through the continuous absolute count.
        sun = C.excel_week_sunday_formula("CfgYear", f"(CfgStartWeek+{pos}-1)")

        ws.cell(row=L.CD_POS_ROW, column=col, value=f"={_deep_start_pos()}+{wk}")
        d = ws.cell(row=L.CD_DATE_ROW, column=col, value=f"={sun}+{day}")
        d.number_format = "yyyy-mm-dd"
        ws.cell(row=L.CD_HOLIDAY_ROW, column=col,
                value=f"=IF(COUNTIF(HolDates,{c}${L.CD_DATE_ROW})>0,1,0)")
        ws.cell(row=L.CD_WORKDAYS_ROW, column=col,
                value=f'=IFERROR(INDEX(CwWorkdays,{pos}),0)')
        ws.cell(row=L.CD_INWINDOW_ROW, column=col, value=(
            f'=IF(AND({wk}<DeepWeeks,{pos}>=1,{pos}<=CfgHorizon),1,0)'))
        ws.cell(row=L.CD_LABEL_ROW, column=col,
                value=f'=IFERROR(INDEX(CwLabel,{pos}),"")')

    for row, label in [(L.CD_POS_ROW, "Position"), (L.CD_DATE_ROW, "Date"),
                       (L.CD_HOLIDAY_ROW, "Holiday?"),
                       (L.CD_WORKDAYS_ROW, "Work days in week"),
                       (L.CD_INWINDOW_ROW, "In window?"), (L.CD_LABEL_ROW, "Label")]:
        ws.cell(row=row, column=1, value=label).font = S.FONT_NOTE
