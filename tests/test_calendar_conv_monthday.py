"""Outlook books the 29th-31st on the last day of shorter months; SOGo only expands
BYMONTHDAY=-1 (it ignores BYSETPOS without BYDAY), so that is the form emitted for the
31st and for 29 February, and shorter months for the 29th/30th become RDATEs."""

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


def _vevent(master: dict, instances: list[dict] | None = None):
    converted = convert_event(master, instances or [], window=WINDOW)
    cal = Calendar.from_ical(converted.ics)
    return next(c for c in cal.walk("VEVENT")), converted


def test_absolute_monthly_31st_becomes_last_day_rule() -> None:
    pattern = {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 31}
    vevent, _ = _vevent(
        _master(pattern, "2026-01-31T09:00:00.0000000", "2026-01-31T10:00:00.0000000")
    )
    rrule_text = vevent["RRULE"].to_ical().decode()
    assert "BYMONTHDAY=-1" in rrule_text and "BYSETPOS" not in rrule_text
    days = [(d.month, d.day) for d in list(rrulestr(rrule_text, dtstart=vevent["DTSTART"].dt))[:4]]
    assert days == [(1, 31), (2, 28), (3, 31), (4, 30)]


def test_absolute_monthly_15th_is_a_plain_rule() -> None:
    pattern = {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 15}
    vevent, conv = _vevent(
        _master(pattern, "2026-01-15T09:00:00.0000000", "2026-01-15T10:00:00.0000000")
    )
    text = vevent["RRULE"].to_ical().decode()
    assert "BYMONTHDAY=15" in text and "BYSETPOS" not in text
    assert not any("day-of-month" in w for w in conv.warnings)


def test_absolute_yearly_feb_29_becomes_last_day_of_february() -> None:
    pattern = {"type": "absoluteYearly", "interval": 1, "month": 2, "dayOfMonth": 29}
    vevent, _ = _vevent(
        _master(pattern, "2024-02-29T09:00:00.0000000", "2024-02-29T10:00:00.0000000")
    )
    text = vevent["RRULE"].to_ical().decode()
    assert "BYMONTH=2" in text and "BYMONTHDAY=-1" in text
    rule = rrulestr(text, dtstart=vevent["DTSTART"].dt)
    years = [(d.year, d.month, d.day) for d in list(rule)[:3]]
    assert years == [(2024, 2, 29), (2025, 2, 28), (2026, 2, 28)]


def test_absolute_monthly_30th_short_month_occurrence_becomes_rdate() -> None:
    pattern = {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 30}
    master = _master(pattern, "2026-01-30T09:00:00.0000000", "2026-01-30T10:00:00.0000000")
    # Graph (Outlook semantics) reports the February occurrence on the 28th
    starts = [f"2026-{m:02d}-{28 if m == 2 else 30}T09:00:00.0000000" for m in range(1, 13)]
    instances = [
        {"type": "occurrence", "seriesMasterId": master["id"], "originalStart": start,
         "start": {"dateTime": start, "timeZone": "UTC"},
         "end": {"dateTime": start[:11] + "10:00:00.0000000", "timeZone": "UTC"}}
        for start in starts
    ]
    vevent, conv = _vevent(master, instances)
    assert "BYMONTHDAY=30" in vevent["RRULE"].to_ical().decode()
    rdates = [d.dt for d in vevent["RDATE"].dts]
    assert [(d.month, d.day) for d in rdates] == [(2, 28)]
    assert "EXDATE" not in vevent
    assert any("RDATE" in w for w in conv.warnings)


def test_overlapping_instance_at_window_start_is_not_an_rdate() -> None:
    """Graph returns occurrences that overlap the window; one in progress at the window
    start (or on its partial first day) must not be duplicated as an RDATE."""
    pattern = {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"],
               "firstDayOfWeek": "monday"}
    master = _master(pattern, "2025-12-29T09:00:00.0000000", "2025-12-29T10:00:00.0000000")
    window = (datetime(2026, 1, 5, 9, 30, tzinfo=UTC), datetime(2026, 2, 1, tzinfo=UTC))
    starts = ["2026-01-05T09:00:00.0000000", "2026-01-12T09:00:00.0000000",
              "2026-01-19T09:00:00.0000000", "2026-01-26T09:00:00.0000000"]
    instances = [
        {"type": "occurrence", "seriesMasterId": master["id"], "originalStart": s,
         "start": {"dateTime": s, "timeZone": "UTC"},
         "end": {"dateTime": s[:11] + "10:00:00.0000000", "timeZone": "UTC"}}
        for s in starts
    ]
    converted = convert_event(master, instances, window=window)
    vevent = next(c for c in Calendar.from_ical(converted.ics).walk("VEVENT"))
    assert "RDATE" not in vevent and "EXDATE" not in vevent
