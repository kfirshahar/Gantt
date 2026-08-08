"""Sunday-start week arithmetic, shared by the generator and the reference model.

Convention: week 1 is the Sunday-start week containing 1 January. The work week
runs Sunday..Thursday, so the five working days sit contiguously at the front of
each numbered week and Friday/Saturday close it.
"""

from datetime import date, timedelta

WORKDAYS_PER_WEEK = 5


def week1_sunday(year: int) -> date:
    """Sunday that opens week 1 of `year`."""
    jan1 = date(year, 1, 1)
    # weekday(): Mon=0..Sun=6; shift so Sunday maps to 0.
    return jan1 - timedelta(days=(jan1.weekday() + 1) % 7)


def week_sunday(year: int, week: int) -> date:
    """Sunday that opens week `week` of `year`."""
    return week1_sunday(year) + timedelta(days=7 * (week - 1))


def workdays(year: int, week: int) -> list[date]:
    """The five Sunday..Thursday dates of a week."""
    start = week_sunday(year, week)
    return [start + timedelta(days=i) for i in range(WORKDAYS_PER_WEEK)]


def day_of(year: int, week: int, day_index: int) -> date:
    """`day_index` 0..4 maps to Sunday..Thursday of `week`."""
    return week_sunday(year, week) + timedelta(days=day_index)


def excel_week_sunday_formula(year_ref: str, week_expr: str) -> str:
    """Excel expression for the Sunday opening a given week.

    Mirrors `week_sunday`: DATE(y,1,1) minus its Sunday offset, plus whole
    weeks. WEEKDAY(d,1) returns 1 for Sunday, so the offset is WEEKDAY()-1.
    """
    jan1 = f"DATE({year_ref},1,1)"
    return f"({jan1}-(WEEKDAY({jan1},1)-1)+({week_expr}-1)*7)"
