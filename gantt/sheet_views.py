"""The read-only Gantt views."""

from openpyxl.formatting.rule import Rule
from openpyxl.styles import Border, Font, Side
from openpyxl.styles.differential import DifferentialStyle

from . import demo, layout as L, names as N, styles as S

# --- Gantt-High ------------------------------------------------------------

GH_WEEK_COL = 3          # weeks start at column C
GH_LOAD_HDR = 4
GH_LOAD_FIRST = 5


def build_gantt_high(ws) -> None:
    S.title(ws, "Gantt — high level",
            "Read-only. Weeks run across; every figure is in work days.")

    last_week_col = GH_WEEK_COL + L.WEEK_COLS - 1
    load_last = _assignee_load(ws, last_week_col)
    timeline_last = _task_timeline(ws, load_last + 2, last_week_col)
    equip_last = _equipment_block(ws, timeline_last + 2, last_week_col)
    _checks(ws, equip_last + 2)

    S.widths(ws, {"A": 38, "B": 9})
    for i in range(L.WEEK_COLS):
        ws.column_dimensions[L.col(GH_WEEK_COL + i)].width = 8
    ws.freeze_panes = L.col(GH_WEEK_COL) + str(GH_LOAD_FIRST)


def _week_header(ws, row: int, last_col: int, first_labels: list[str]) -> None:
    S.header_row(ws, row, 1, first_labels, output=True)
    for i in range(L.WEEK_COLS):
        cell = ws.cell(row=row, column=GH_WEEK_COL + i,
                       value=f"=INDEX(CwLabel,{i + 1})")
        cell.fill = S.FILL_OUTPUT_HDR
        cell.font = S.FONT_HDR
        cell.alignment = S.CENTER
        cell.border = S.BORDER_ALL


def _assignee_load(ws, last_col: int) -> int:
    """Two rows per assignee: days used, days available. Red where used > available."""
    S.section(ws, GH_LOAD_HDR - 1, "Assignee load — days used vs days available",
              span=last_col)
    _week_header(ws, GH_LOAD_HDR, last_col, ["Assignee", "Metric"])

    for i in range(L.MAX_ASSIGNEES):
        used_row = GH_LOAD_FIRST + 2 * i
        avail_row = used_row + 1
        aw_row = L.CW_AWWEEK_FIRST + i
        src = L.GRID_FIRST_DATA_ROW + i

        ws.cell(row=used_row, column=1, value=(
            f'=IF({L.sheet_ref(L.ASSIGNEES)}!A{src}="","",'
            f'{L.sheet_ref(L.ASSIGNEES)}!A{src})')).font = S.FONT_SECTION
        # Labels vanish on unused rows so the pre-built capacity stays invisible
        # until someone actually adds that assignee.
        ws.cell(row=used_row, column=2,
                value=f'=IF($A{used_row}="","","Used")').font = S.FONT_BODY
        ws.cell(row=avail_row, column=2,
                value=f'=IF($A{used_row}="","","Avail")').font = S.FONT_NOTE

        for w in range(L.WEEK_COLS):
            c = GH_WEEK_COL + w
            wc = L.col(L.CW_FIRST_WEEK_COL + w)
            used = ws.cell(row=used_row, column=c, value=(
                f'=IF($A{used_row}="","",{L.sheet_ref(L.CALC_WEEK)}!{wc}{aw_row})'))
            avail = ws.cell(row=avail_row, column=c, value=(
                f'=IF($A{used_row}="","",IFERROR(INDEX(CapGrid,'
                f'MATCH($A{used_row},CapNames,0),{w + 1}),0))'))
            for cell in (used, avail):
                cell.number_format = "0.0"
                cell.alignment = S.CENTER
                cell.border = S.BORDER_ALL
            avail.font = S.FONT_NOTE

        # Allocation is capped at capacity, so "used > avail" can never happen.
        # The useful signal is saturation: a week with no slack is a week where
        # this person is the bottleneck holding the plan back.
        span = f"{L.col(GH_WEEK_COL)}{used_row}:{L.col(last_col)}{used_row}"
        ws.conditional_formatting.add(span, Rule(
            type="expression", dxf=DifferentialStyle(fill=S.FILL_OVER, font=Font(bold=True)),
            formula=[f'AND({L.col(GH_WEEK_COL)}{used_row}<>"",'
                     f'{L.col(GH_WEEK_COL)}{avail_row}>0,'
                     f'{L.col(GH_WEEK_COL)}{used_row}>='
                     f'{L.col(GH_WEEK_COL)}{avail_row})']))
        ws.conditional_formatting.add(span, Rule(
            type="expression", dxf=DifferentialStyle(fill=S.FILL_OK),
            formula=[f'AND({L.col(GH_WEEK_COL)}{used_row}>0,'
                     f'{L.col(GH_WEEK_COL)}{used_row}<{L.col(GH_WEEK_COL)}{avail_row})']))

        ws.cell(row=used_row, column=1).border = S.BORDER_ALL
        ws.cell(row=avail_row, column=1).border = Border(
            top=Side(style="thin", color="D2D6DC"),
            bottom=Side(style="medium", color=S.INK))

    last = GH_LOAD_FIRST + 2 * L.MAX_ASSIGNEES - 1
    S.grey_inactive_weeks(ws, GH_WEEK_COL, last_col, GH_LOAD_HDR, last)
    return last


def _task_timeline(ws, start_row: int, last_col: int) -> int:
    """One row per task: allocated days per week, shaded where non-zero."""
    S.section(ws, start_row, "Task timeline — allocated days per week", span=last_col)
    hdr = start_row + 1
    _week_header(ws, hdr, last_col, ["Task", "Priority"])
    first = hdr + 1

    for i in range(L.MAX_TASKS):
        r = first + i
        task_row = L.TASK_FIRST_ROW + i
        tw_row = L.CW_TASKWEEK_FIRST + i
        t = L.sheet_ref(L.TASKS)
        ws.cell(row=r, column=1, value=(
            f'=IF({t}!A{task_row}="","",{t}!A{task_row}&"  "&{t}!B{task_row})'))
        ws.cell(row=r, column=2, value=(
            f'=IF({t}!A{task_row}="","",{t}!D{task_row})')).alignment = S.CENTER

        for w in range(L.WEEK_COLS):
            wc = L.col(L.CW_FIRST_WEEK_COL + w)
            cell = ws.cell(row=r, column=GH_WEEK_COL + w, value=(
                f'=IF($A{r}="","",IF({L.sheet_ref(L.CALC_WEEK)}!{wc}{tw_row}=0,"",'
                f'{L.sheet_ref(L.CALC_WEEK)}!{wc}{tw_row}))'))
            cell.number_format = "0.0"
            cell.alignment = S.CENTER
            cell.border = S.BORDER_ALL
        for c in (1, 2):
            ws.cell(row=r, column=c).border = S.BORDER_ALL

    last = first + L.MAX_TASKS - 1
    ws.conditional_formatting.add(
        f"{L.col(GH_WEEK_COL)}{first}:{L.col(last_col)}{last}",
        Rule(type="cellIs", operator="greaterThan", formula=["0"],
             dxf=DifferentialStyle(fill=S.FILL_BAR)))
    S.grey_inactive_weeks(ws, GH_WEEK_COL, last_col, hdr, last)
    return last


def _equipment_block(ws, start_row: int, last_col: int) -> int:
    """Two rows per equipment type: units demanded, units available."""
    S.section(ws, start_row, "Equipment — units demanded vs units available",
              span=last_col)
    hdr = start_row + 1
    _week_header(ws, hdr, last_col, ["Equipment", "Metric"])
    first = hdr + 1

    for i in range(L.MAX_EQUIPMENT):
        dem_row = first + 2 * i
        sup_row = dem_row + 1
        eq_row = L.CW_EQPWEEK_FIRST + i
        src = L.GRID_FIRST_DATA_ROW + i

        ws.cell(row=dem_row, column=1, value=(
            f'=IF({L.sheet_ref(L.EQUIPMENT)}!A{src}="","",'
            f'{L.sheet_ref(L.EQUIPMENT)}!A{src})')).font = S.FONT_SECTION
        ws.cell(row=dem_row, column=2,
                value=f'=IF($A{dem_row}="","","Need")').font = S.FONT_BODY
        ws.cell(row=sup_row, column=2,
                value=f'=IF($A{dem_row}="","","Have")').font = S.FONT_NOTE

        for w in range(L.WEEK_COLS):
            c = GH_WEEK_COL + w
            wc = L.col(L.CW_FIRST_WEEK_COL + w)
            need = ws.cell(row=dem_row, column=c, value=(
                f'=IF($A{dem_row}="","",{L.sheet_ref(L.CALC_WEEK)}!{wc}{eq_row})'))
            have = ws.cell(row=sup_row, column=c, value=(
                f'=IF($A{dem_row}="","",IFERROR(INDEX(EqpGrid,'
                f'MATCH($A{dem_row},EqpNames,0),{w + 1}),0))'))
            for cell in (need, have):
                cell.number_format = "0"
                cell.alignment = S.CENTER
                cell.border = S.BORDER_ALL
            have.font = S.FONT_NOTE

        span = f"{L.col(GH_WEEK_COL)}{dem_row}:{L.col(last_col)}{dem_row}"
        ws.conditional_formatting.add(span, Rule(
            type="expression", dxf=DifferentialStyle(fill=S.FILL_OVER, font=Font(bold=True)),
            formula=[f'AND({L.col(GH_WEEK_COL)}{dem_row}<>"",'
                     f'{L.col(GH_WEEK_COL)}{dem_row}>{L.col(GH_WEEK_COL)}{sup_row})']))
        ws.cell(row=dem_row, column=1).border = S.BORDER_ALL

    last = first + 2 * L.MAX_EQUIPMENT - 1
    S.grey_inactive_weeks(ws, GH_WEEK_COL, last_col, hdr, last)
    return last


def _checks(ws, start_row: int) -> None:
    """Conditions that make a plan wrong, as distinct from merely tight."""
    S.section(ws, start_row, "Checks", span=6)
    S.header_row(ws, start_row + 1, 1, ["Check", "Count"], output=True)

    sub_check = L.abs_range(L.SUBTASKS, L.S_CHECK, L.SUB_FIRST_ROW, L.S_CHECK, N.LAST_SUB_ROW)
    task_check = L.abs_range(L.TASKS, L.T_CHECK, L.TASK_FIRST_ROW, L.T_CHECK, N.LAST_TASK_ROW)

    # Note there is deliberately no "over capacity" check: the spill-over engine
    # caps each week's allocation at what the assignee has, so demand never
    # exceeds supply. Excess work shows up as work that misses the horizon.
    checks = [
        ("Sub-task rows with a problem", f'=COUNTIF({sub_check},"⚠*")'),
        ("Tasks with a problem", f'=COUNTIF({task_check},"⚠*")'),
        ("Assignees with zero capacity all horizon", '=COUNTIFS(AsgNames,"<>",CapTotals,0)'),
        ("Work days that do not fit in the horizon",
         "=ROUND(SUM(SubEffort)-SUM(CwGrid),2)"),
        ("Assignee-weeks with no slack left",
         "=SUMPRODUCT((AwGrid>=CapGrid)*(CapGrid>0))"),
        ("Equipment-weeks short of units", "=SUMPRODUCT((EqDemand>EqpGrid)*1)"),
    ]
    for i, (label, formula) in enumerate(checks):
        r = start_row + 2 + i
        ws.cell(row=r, column=1, value=label).border = S.BORDER_ALL
        cell = ws.cell(row=r, column=2, value=formula)
        cell.alignment = S.CENTER
        cell.border = S.BORDER_ALL

    span = f"B{start_row + 2}:B{start_row + 1 + len(checks)}"
    ws.conditional_formatting.add(span, Rule(
        type="cellIs", operator="greaterThan", formula=["0"],
        dxf=DifferentialStyle(fill=S.FILL_WARN, font=Font(bold=True))))
    ws.conditional_formatting.add(span, Rule(
        type="cellIs", operator="equal", formula=["0"],
        dxf=DifferentialStyle(fill=S.FILL_OK)))

    note = ws.cell(row=start_row + 2 + len(checks) + 1, column=1, value=(
        "Saturation and equipment shortage are information, not errors — they show "
        "where the plan is tight. Work that does not fit the horizon is the row to act on."))
    note.font = S.FONT_NOTE


# --- Gantt-Deep ------------------------------------------------------------

GD_WEEK_ROW = 7
GD_DAY_ROW = 8
GD_DATE_ROW = 9
GD_FIRST_ROW = 10
GD_DAY_COL = 5           # day columns start at column E


def build_gantt_deep(ws) -> None:
    S.title(ws, "Gantt — deep dive",
            "Read-only except the two window cells. Day columns, sub-task rows.")

    ws.cell(row=3, column=1, value="Window start week").font = S.FONT_SECTION
    start = ws.cell(row=3, column=2, value="=CfgStartWeek")
    start.number_format = '"WW"0'
    ws.cell(row=4, column=1, value="Weeks to show").font = S.FONT_SECTION
    weeks = ws.cell(row=4, column=2, value=4)
    for cell in (start, weeks):
        cell.border = S.BORDER_ALL
        cell.alignment = S.CENTER
        cell.fill = S.FILL_WARN
    ws.cell(row=4, column=3,
            value=f"Type over either cell. Maximum {L.MAX_DEEP_WEEKS} weeks — "
                  f"wider windows slow recalculation.").font = S.FONT_NOTE

    _deep_headers(ws)
    _deep_rows(ws)

    S.widths(ws, {"A": 12, "B": 12, "C": 26, "D": 14})
    for i in range(L.DEEP_DAY_COLS):
        ws.column_dimensions[L.col(GD_DAY_COL + i)].width = 5.5
    ws.freeze_panes = L.col(GD_DAY_COL) + str(GD_FIRST_ROW)


def _deep_headers(ws) -> None:
    S.header_row(ws, GD_WEEK_ROW, 1,
                 ["Parent", "Sub-task", "Name", "Assignee"], output=True)
    for row in (GD_DAY_ROW, GD_DATE_ROW):
        for c in range(1, GD_DAY_COL):
            ws.cell(row=row, column=c).fill = S.FILL_OUTPUT_HDR

    for i in range(L.DEEP_DAY_COLS):
        c = GD_DAY_COL + i
        src = L.col(L.CD_FIRST_DAY_COL + i)
        cd = L.sheet_ref(L.CALC_DAY)

        week = ws.cell(row=GD_WEEK_ROW, column=c, value=f"={cd}!{src}${L.CD_LABEL_ROW}")
        day = ws.cell(row=GD_DAY_ROW, column=c,
                      value=f'="{L.DAY_NAMES[i % L.WORKDAYS_PER_WEEK]}"')
        date = ws.cell(row=GD_DATE_ROW, column=c, value=f"={cd}!{src}${L.CD_DATE_ROW}")
        date.number_format = "dd/mm"

        for cell in (week, day, date):
            cell.fill = S.FILL_OUTPUT_HDR
            cell.font = S.FONT_HDR
            cell.alignment = S.CENTER
            cell.border = S.BORDER_ALL

    last = GD_DAY_COL + L.DEEP_DAY_COLS - 1
    cd = L.sheet_ref(L.CALC_DAY)
    # Grey out columns outside the window and holidays; they cannot be hidden
    # dynamically without macros.
    for row in (GD_WEEK_ROW, GD_DAY_ROW, GD_DATE_ROW):
        ws.conditional_formatting.add(
            f"{L.col(GD_DAY_COL)}{row}:{L.col(last)}{row}",
            Rule(type="expression", dxf=DifferentialStyle(fill=S.FILL_WEEKEND,
                                                          font=Font(color=S.MUTED)),
                 formula=[f"{cd}!{L.col(L.CD_FIRST_DAY_COL)}${L.CD_INWINDOW_ROW}=0"]))


def _deep_rows(ws) -> None:
    cd = L.sheet_ref(L.CALC_DAY)
    last_col = GD_DAY_COL + L.DEEP_DAY_COLS - 1

    for i in range(L.MAX_SUBTASKS):
        r = GD_FIRST_ROW + i
        cdr = L.CD_FIRST_ROW + i

        ws.cell(row=r, column=1, value=f'=IF({cd}!$F{cdr}="","",{cd}!$F{cdr})')
        ws.cell(row=r, column=2, value=f'=IF({cd}!$B{cdr}="","",{cd}!$B{cdr})')
        ws.cell(row=r, column=3, value=(
            f'=IF({cd}!$B{cdr}="","",IFERROR(INDEX(SubName,'
            f'MATCH({cd}!$A{cdr},SubRank,0)),""))'))
        ws.cell(row=r, column=4, value=f'=IF({cd}!$C{cdr}="","",{cd}!$C{cdr})')

        for j in range(L.DEEP_DAY_COLS):
            src = L.col(L.CD_FIRST_DAY_COL + j)
            cell = ws.cell(row=r, column=GD_DAY_COL + j, value=(
                f'=IF({cd}!$B{cdr}="","",IF({cd}!{src}{cdr}=0,"",{cd}!{src}{cdr}))'))
            cell.number_format = "0.0"
            cell.alignment = S.CENTER
            cell.border = S.BORDER_ALL
        for c in range(1, GD_DAY_COL):
            ws.cell(row=r, column=c).border = S.BORDER_ALL

    last_row = GD_FIRST_ROW + L.MAX_SUBTASKS - 1
    body = f"{L.col(GD_DAY_COL)}{GD_FIRST_ROW}:{L.col(last_col)}{last_row}"
    ws.conditional_formatting.add(body, Rule(
        type="cellIs", operator="greaterThan", formula=["0"],
        dxf=DifferentialStyle(fill=S.FILL_BAR)))
    ws.conditional_formatting.add(body, Rule(
        type="expression", dxf=DifferentialStyle(fill=S.FILL_WEEKEND),
        formula=[f"{cd}!{L.col(L.CD_FIRST_DAY_COL)}${L.CD_INWINDOW_ROW}=0"]))

    # Same hide-the-repeats treatment as the Sub-Tasks tab.
    ws.conditional_formatting.add(
        f"A{GD_FIRST_ROW}:A{last_row}",
        Rule(type="expression", dxf=DifferentialStyle(font=Font(color="FFFFFF")),
             formula=[f'AND($A{GD_FIRST_ROW}<>"",$A{GD_FIRST_ROW}=$A{GD_FIRST_ROW - 1})']))
    ws.conditional_formatting.add(
        f"A{GD_FIRST_ROW}:{L.col(last_col)}{last_row}",
        Rule(type="expression",
             dxf=DifferentialStyle(border=Border(top=Side(style="medium", color=S.INK))),
             formula=[f'AND($A{GD_FIRST_ROW}<>"",$A{GD_FIRST_ROW}<>$A{GD_FIRST_ROW - 1})']))
