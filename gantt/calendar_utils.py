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


def calendar_week(d: date) -> tuple[int, int]:
    """The (year, week) a Sunday-start week belongs to.

    A week belongs to the year containing its *last* day, because week 1 is
    defined as the week containing 1 January: if a week straddles New Year it
    contains that January the 1st and therefore belongs to the later year. This
    also makes 52- and 53-week years fall out automatically.
    """
    year = (d + timedelta(days=6)).year
    return year, (d - week1_sunday(year)).days // 7 + 1


def excel_calendar_year_formula(sunday_ref: str) -> str:
    """Excel expression for the year a week belongs to. Mirrors `calendar_week`."""
    return f"YEAR({sunday_ref}+6)"


def excel_calendar_week_formula(sunday_ref: str, year_ref: str) -> str:
    """Excel expression for the calendar week number. Mirrors `calendar_week`."""
    jan1 = f"DATE({year_ref},1,1)"
    return f"(({sunday_ref}-({jan1}-(WEEKDAY({jan1},1)-1)))/7+1)"


def excel_week_label_formula(week_ref: str, year_ref: str, project_year_ref: str) -> str:
    """Display text for a week: WW04, or WW01 '27 once the year changes."""
    return (f'"WW"&TEXT({week_ref},"00")'
            f'&IF({year_ref}<>{project_year_ref}," \'"&TEXT({year_ref}-2000,"00"),"")')


def excel_week_sunday_formula(year_ref: str, week_expr: str) -> str:
    """Excel expression for the Sunday opening a given week.

    Mirrors `week_sunday`: DATE(y,1,1) minus its Sunday offset, plus whole
    weeks. WEEKDAY(d,1) returns 1 for Sunday, so the offset is WEEKDAY()-1.
    """
    jan1 = f"DATE({year_ref},1,1)"
    return f"({jan1}-(WEEKDAY({jan1},1)-1)+({week_expr}-1)*7)"
