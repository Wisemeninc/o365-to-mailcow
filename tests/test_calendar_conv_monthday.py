"""Outlook books the 29th-31st on the last day of shorter months; the RRULE must too."""

from __future__ import annotations

from datetime import UTC, datetime

from dateutil.rrule import rrulestr
from icalendar import Calendar

from o365_to_mailcow.calendar_conv import convert_event

WINDOW = (datetime(2026, 1, 1, tzinfo=UTC), datetime(2027, 1, 1, tzinfo=UTC))


def _master(pattern: dict, start: str, end: str) -> dict:
    return {
        "id": "AAMk-monthday",
        "iCalUId": "monthday-uid",
        "type": "seriesMaster",
        "subject": "Invoice run",
        "isAllDay": False,
        "start": {"dateTime": start, "timeZone": "UTC"},
        "end": {"dateTime": end, "timeZone": "UTC"},
        "originalStartTimeZone": "UTC",
        "originalEndTimeZone": "UTC",
        "recurrence": {
            "pattern": pattern,
            "range": {"type": "noEnd", "startDate": start[:10]},
        },
    }


def _rrule_and_dtstart(master: dict) -> tuple[str, datetime]:
    converted = convert_event(master, [], window=WINDOW)
    cal = Calendar.from_ical(converted.ics)
    vevent = next(c for c in cal.walk("VEVENT"))
    return vevent["RRULE"].to_ical().decode(), vevent["DTSTART"].dt


def _occurrence_days(rrule_text: str, dtstart: datetime, count: int) -> list[tuple[int, int]]:
    rule = rrulestr(rrule_text, dtstart=dtstart)
    return [(d.month, d.day) for d in list(rule)[:count]]


def test_absolute_monthly_31st_lands_on_last_day_of_short_months() -> None:
    pattern = {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 31}
    rrule_text, dtstart = _rrule_and_dtstart(
        _master(pattern, "2026-01-31T09:00:00.0000000", "2026-01-31T10:00:00.0000000")
    )
    assert "BYSETPOS=-1" in rrule_text
    assert _occurrence_days(rrule_text, dtstart, 5) == [
        (1, 31), (2, 28), (3, 31), (4, 30), (5, 31),
    ]


def test_absolute_monthly_15th_is_a_plain_rule() -> None:
    pattern = {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 15}
    rrule_text, _ = _rrule_and_dtstart(
        _master(pattern, "2026-01-15T09:00:00.0000000", "2026-01-15T10:00:00.0000000")
    )
    assert "BYMONTHDAY=15" in rrule_text
    assert "BYSETPOS" not in rrule_text


def test_absolute_yearly_feb_29_falls_back_to_feb_28() -> None:
    pattern = {"type": "absoluteYearly", "interval": 1, "month": 2, "dayOfMonth": 29}
    rrule_text, dtstart = _rrule_and_dtstart(
        _master(pattern, "2024-02-29T09:00:00.0000000", "2024-02-29T10:00:00.0000000")
    )
    rule = rrulestr(rrule_text, dtstart=dtstart)
    years = [(d.year, d.month, d.day) for d in list(rule)[:4]]
    assert years == [(2024, 2, 29), (2025, 2, 28), (2026, 2, 28), (2027, 2, 28)]
