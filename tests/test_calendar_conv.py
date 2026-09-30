"""Graph event -> iCalendar (ISC-58..79, ISC-105, ISC-106, ISC-122, ISC-123).

Every generated calendar is parsed back with icalendar; recurrence semantics are checked
by expanding the emitted RRULE with dateutil, the same way a CalDAV client would.
"""

from __future__ import annotations

import copy
import json
import re
import uuid
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import pytest
from dateutil.rrule import rrulestr
from icalendar import Calendar

from o365_to_mailcow.calendar_conv import ConvertedEvent, convert_event

FIXTURES = Path(__file__).parent / "fixtures" / "graph"
WINDOW = (datetime(2026, 1, 1, tzinfo=UTC), datetime(2027, 1, 1, tzinfo=UTC))
BERLIN = ZoneInfo("Europe/Berlin")
PACIFIC = ZoneInfo("America/Los_Angeles")


# -- helpers -------------------------------------------------------------------------------


def load(name: str) -> dict[str, Any]:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def weekly() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    data = load("event_weekly_series.json")
    return data["master"], data["instances"]


def convert(
    master: dict[str, Any],
    instances: list[dict[str, Any]] | None = None,
    *,
    window: tuple[datetime, datetime] = WINDOW,
    attendees: str = "keep",
) -> ConvertedEvent:
    result = convert_event(master, instances or [], window=window, attendees=attendees)
    assert_round_trip(result)
    return result


def assert_round_trip(result: ConvertedEvent) -> None:
    """ISC-79: parses back, and every TZID used has a VTIMEZONE."""
    cal = Calendar.from_ical(result.ics)
    defined = {str(tz["TZID"]) for tz in cal.walk("VTIMEZONE")}
    used = set(re.findall(r";TZID=([^:;]+)", unfolded(result)))
    assert used <= defined, f"TZIDs without VTIMEZONE: {used - defined}"
    assert defined <= used, f"unused VTIMEZONEs: {defined - used}"


def unfolded(result: ConvertedEvent) -> str:
    return result.ics.decode("utf-8").replace("\r\n ", "")


def vevents(result: ConvertedEvent) -> list[Any]:
    return Calendar.from_ical(result.ics).walk("VEVENT")


def master_of(result: ConvertedEvent) -> Any:
    masters = [ev for ev in vevents(result) if "RECURRENCE-ID" not in ev]
    assert len(masters) == 1
    return masters[0]


def exceptions_of(result: ConvertedEvent) -> list[Any]:
    return [ev for ev in vevents(result) if "RECURRENCE-ID" in ev]


def rrule_text(result: ConvertedEvent) -> str:
    lines = [ln for ln in unfolded(result).split("\r\n") if ln.startswith("RRULE:")]
    assert len(lines) == 1
    return lines[0].removeprefix("RRULE:")


def expand(result: ConvertedEvent) -> list[datetime]:
    """Expand the emitted master exactly as a client would (RRULE from the output)."""
    dtstart = master_of(result)["DTSTART"].dt
    if not isinstance(dtstart, datetime):
        dtstart = datetime.combine(dtstart, datetime.min.time())
    return list(rrulestr(rrule_text(result), dtstart=dtstart))


def exdates(event: Any) -> list[datetime | date]:
    raw = event.get("EXDATE")
    if raw is None:
        return []
    groups = raw if isinstance(raw, list) else [raw]
    return [item.dt for group in groups for item in group.dts]


def recurring(pattern: dict[str, Any], range_: dict[str, Any] | None = None) -> dict[str, Any]:
    """Weekly fixture master with its pattern/range replaced."""
    master, _ = weekly()
    master = copy.deepcopy(master)
    master["recurrence"] = {
        "pattern": {"interval": 1, "firstDayOfWeek": "sunday", **pattern},
        "range": range_ or {"type": "noEnd", "startDate": "2026-03-16"},
    }
    return master


# -- ISC-58, ISC-106: UID ------------------------------------------------------------------


def test_isc58_uid_is_ical_uid() -> None:
    master = load("event_single_timed.json")
    result = convert(master)
    assert result.uid == master["iCalUId"]
    assert str(master_of(result)["UID"]) == master["iCalUId"]


def test_isc106_missing_ical_uid_uses_uuid5_of_graph_id() -> None:
    master = load("event_single_timed.json")
    del master["iCalUId"]
    expected = str(uuid.uuid5(uuid.NAMESPACE_URL, "o365-to-mailcow:event:" + master["id"]))
    result = convert(master)
    assert result.uid == expected
    assert str(master_of(result)["UID"]) == expected
    assert convert(master).uid == expected  # deterministic


def test_last_modified_is_the_masters_graph_value() -> None:
    master = load("event_single_timed.json")
    assert convert(master).last_modified == master["lastModifiedDateTime"]


def test_calendar_envelope() -> None:
    cal = Calendar.from_ical(convert(load("event_single_timed.json")).ics)
    assert str(cal["VERSION"]) == "2.0"
    assert str(cal["PRODID"]) == "-//o365-to-mailcow//EN"
    assert str(cal["CALSCALE"]) == "GREGORIAN"


# -- ISC-59..62: times and zones -----------------------------------------------------------


def test_isc59_timed_event_uses_original_zone() -> None:
    result = convert(load("event_single_timed.json"))
    raw = unfolded(result)
    assert "DTSTART;TZID=Europe/Berlin:20260330T090000\r\n" in raw
    assert "DTEND;TZID=Europe/Berlin:20260330T100000\r\n" in raw
    event = master_of(result)
    assert event["DTSTART"].dt == datetime(2026, 3, 30, 9, 0, tzinfo=BERLIN)
    assert event["DTSTART"].dt.astimezone(UTC) == datetime(2026, 3, 30, 7, 0, tzinfo=UTC)
    assert event["DTEND"].dt == datetime(2026, 3, 30, 10, 0, tzinfo=BERLIN)
    assert result.warnings == ()


def test_isc59_end_zone_can_differ_from_start_zone() -> None:
    master = load("event_single_timed.json")
    master["originalEndTimeZone"] = "GMT Standard Time"
    raw = unfolded(convert(master))
    assert "DTSTART;TZID=Europe/Berlin:20260330T090000\r\n" in raw
    assert "DTEND;TZID=Europe/London:20260330T090000\r\n" in raw


def test_isc60_all_day_event_uses_exclusive_dates() -> None:
    result = convert(load("event_all_day.json"))
    raw = unfolded(result)
    assert "DTSTART;VALUE=DATE:20260501\r\n" in raw
    assert "DTEND;VALUE=DATE:20260502\r\n" in raw
    event = master_of(result)
    assert event["DTSTART"].dt == date(2026, 5, 1)
    assert event["DTEND"].dt == date(2026, 5, 2)
    assert "BEGIN:VTIMEZONE" not in raw


def test_isc60_all_day_end_not_after_start_is_forced_to_next_day() -> None:
    master = load("event_all_day.json")
    master["end"] = dict(master["start"])
    result = convert(master)
    assert master_of(result)["DTEND"].dt == date(2026, 5, 2)
    assert any("DTSTART + 1 day" in w for w in result.warnings)


def test_isc60_all_day_given_as_local_midnight_instant() -> None:
    master = load("event_all_day.json")  # W. Europe: 2026-05-01 00:00 CEST = 04-30 22:00Z
    master["start"] = {"dateTime": "2026-04-30T22:00:00.0000000", "timeZone": "UTC"}
    master["end"] = {"dateTime": "2026-05-01T22:00:00.0000000", "timeZone": "UTC"}
    event = master_of(convert(master))
    assert (event["DTSTART"].dt, event["DTEND"].dt) == (date(2026, 5, 1), date(2026, 5, 2))


@pytest.mark.parametrize(
    ("windows", "tzid"),
    [
        ("W. Europe Standard Time", "Europe/Berlin"),
        ("Romance Standard Time", "Europe/Paris"),
        ("GMT Standard Time", "Europe/London"),
        ("Pacific Standard Time", "America/Los_Angeles"),
    ],
)
def test_isc61_windows_zone_becomes_iana_tzid(windows: str, tzid: str) -> None:
    master = load("event_single_timed.json")
    master["originalStartTimeZone"] = master["originalEndTimeZone"] = windows
    result = convert(master)
    local = datetime(2026, 3, 30, 7, 0, tzinfo=UTC).astimezone(ZoneInfo(tzid))
    assert f"DTSTART;TZID={tzid}:{local:%Y%m%dT%H%M%S}\r\n" in unfolded(result)


@pytest.mark.parametrize("utc_name", ["UTC", "tzone://Microsoft/Utc"])
def test_isc61_utc_zone_uses_z_form_without_vtimezone(utc_name: str) -> None:
    master = load("event_single_timed.json")
    master["originalStartTimeZone"] = master["originalEndTimeZone"] = utc_name
    result = convert(master)
    raw = unfolded(result)
    assert "DTSTART:20260330T070000Z\r\n" in raw
    assert "BEGIN:VTIMEZONE" not in raw
    assert result.warnings == ()


def test_isc62_unknown_zone_falls_back_to_utc_and_warns_once() -> None:
    result = convert(load("event_html_body.json"))  # tzone://Microsoft/Custom start and end
    raw = unfolded(result)
    assert "DTSTART:20260901T120000Z\r\n" in raw
    assert "DTEND:20260901T130000Z\r\n" in raw
    assert "TZID=" not in raw
    assert result.warnings == ("unknown time zone 'tzone://Microsoft/Custom' -> UTC",)


def test_isc62_missing_zone_falls_back_to_utc() -> None:
    master = load("event_single_timed.json")
    del master["originalStartTimeZone"], master["originalEndTimeZone"]
    result = convert(master)
    assert "DTSTART:20260330T070000Z\r\n" in unfolded(result)
    assert "unknown time zone '<missing>' -> UTC" in result.warnings


# -- ISC-63..68: recurrence rules ----------------------------------------------------------


def test_isc64_weekly_fixture_rrule() -> None:
    result = convert(*weekly())
    assert rrule_text(result) == "FREQ=WEEKLY;UNTIL=20260427T070000Z;INTERVAL=1;BYDAY=MO;WKST=SU"


def test_isc66_relative_monthly_last_friday_numbered() -> None:
    result = convert(load("event_relative_monthly_last_friday.json"))
    assert rrule_text(result) == "FREQ=MONTHLY;COUNT=6;INTERVAL=1;BYDAY=-1FR"
    starts = expand(result)
    assert [d.date() for d in starts] == [
        date(2026, 1, 30), date(2026, 2, 27), date(2026, 3, 27),
        date(2026, 4, 24), date(2026, 5, 29), date(2026, 6, 26),
    ]
    assert {(d.hour, d.minute) for d in starts} == {(10, 0)}
    assert all(d.tzinfo == PACIFIC for d in starts)


def test_isc67_absolute_yearly_all_day_no_end() -> None:
    result = convert(load("event_absolute_yearly.json"))
    assert rrule_text(result) == "FREQ=YEARLY;INTERVAL=1;BYMONTHDAY=14;BYMONTH=7"
    rule = master_of(result)["RRULE"]
    assert "UNTIL" not in rule and "COUNT" not in rule  # ISC-68 noEnd


@pytest.mark.parametrize(
    ("pattern", "expected"),
    [
        ({"type": "daily", "interval": 3}, {"FREQ": ["DAILY"], "INTERVAL": [3]}),
        (
            {"type": "weekly", "interval": 2, "daysOfWeek": ["monday", "wednesday"],
             "firstDayOfWeek": "monday"},
            {"FREQ": ["WEEKLY"], "INTERVAL": [2], "BYDAY": ["MO", "WE"], "WKST": ["MO"]},
        ),
        (
            {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 15},
            {"FREQ": ["MONTHLY"], "INTERVAL": [1], "BYMONTHDAY": [15]},
        ),
        (
            {"type": "relativeMonthly", "interval": 1, "daysOfWeek": ["tuesday"],
             "index": "second"},
            {"FREQ": ["MONTHLY"], "INTERVAL": [1], "BYDAY": ["2TU"]},
        ),
        (
            {"type": "relativeMonthly", "interval": 1, "index": "first",
             "daysOfWeek": ["monday", "tuesday", "wednesday", "thursday", "friday"]},
            {"FREQ": ["MONTHLY"], "INTERVAL": [1], "BYDAY": ["MO", "TU", "WE", "TH", "FR"],
             "BYSETPOS": [1]},
        ),
        (
            {"type": "absoluteYearly", "interval": 1, "month": 12, "dayOfMonth": 24},
            {"FREQ": ["YEARLY"], "INTERVAL": [1], "BYMONTH": [12], "BYMONTHDAY": [24]},
        ),
        (
            {"type": "relativeYearly", "interval": 1, "month": 11, "daysOfWeek": ["thursday"],
             "index": "fourth"},
            {"FREQ": ["YEARLY"], "INTERVAL": [1], "BYMONTH": [11], "BYDAY": ["4TH"]},
        ),
        (
            {"type": "relativeYearly", "interval": 1, "month": 3, "daysOfWeek": ["sunday"],
             "index": "last"},
            {"FREQ": ["YEARLY"], "INTERVAL": [1], "BYMONTH": [3], "BYDAY": ["-1SU"]},
        ),
    ],
)
def test_isc63_to_67_pattern_mapping(pattern: dict[str, Any], expected: dict[str, Any]) -> None:
    rule = master_of(convert(recurring(pattern)))["RRULE"]
    got = {key: [v if isinstance(v, int) else str(v) for v in values]
           for key, values in rule.items()}
    assert got == expected


def test_isc64_weekly_default_wkst_is_monday() -> None:
    master = recurring({"type": "weekly", "daysOfWeek": ["friday"]})
    del master["recurrence"]["pattern"]["firstDayOfWeek"]
    assert master_of(convert(master))["RRULE"]["WKST"] == ["MO"]


def test_isc63_missing_interval_defaults_to_one() -> None:
    master = recurring({"type": "daily"})
    del master["recurrence"]["pattern"]["interval"]
    assert master_of(convert(master))["RRULE"]["INTERVAL"] == [1]


def test_isc68_until_is_inclusive_utc_start_of_last_occurrence() -> None:
    result = convert(*weekly())
    until = master_of(result)["RRULE"]["UNTIL"][0]
    assert until == datetime(2026, 4, 27, 7, 0, tzinfo=UTC)  # 09:00 CEST on endDate
    starts = expand(result)
    assert starts[-1] == datetime(2026, 4, 27, 9, 0, tzinfo=BERLIN)  # endDate itself included
    assert len(starts) == 7


def test_isc68_until_for_utc_fallback_series() -> None:
    master = recurring(
        {"type": "daily"}, {"type": "endDate", "startDate": "2026-03-16", "endDate": "2026-03-20"}
    )
    master["originalStartTimeZone"] = master["originalEndTimeZone"] = "tzone://Microsoft/Custom"
    result = convert(master, [])
    assert master_of(result)["RRULE"]["UNTIL"][0] == datetime(2026, 3, 20, 8, 0, tzinfo=UTC)
    assert len(expand(result)) == 5


def test_isc68_all_day_until_is_a_date() -> None:
    master = load("event_absolute_yearly.json")
    master["recurrence"]["range"] = {"type": "endDate", "startDate": "2026-07-14",
                                     "endDate": "2028-07-14"}
    instances = [{"type": "occurrence", "originalStart": "2026-07-14T00:00:00Z"}]
    result = convert(master, instances)
    assert "UNTIL=20280714" in rrule_text(result).split(";")
    assert master_of(result)["RRULE"]["UNTIL"][0] == date(2028, 7, 14)
    assert [d.date() for d in expand(result)] == [
        date(2026, 7, 14), date(2027, 7, 14), date(2028, 7, 14),
    ]


def test_isc68_numbered_is_count() -> None:
    master = recurring({"type": "daily"}, {"type": "numbered", "numberOfOccurrences": 4})
    result = convert(master)
    assert master_of(result)["RRULE"]["COUNT"] == [4]
    assert len(expand(result)) == 4


def test_unsupported_pattern_emits_master_without_rrule() -> None:
    master = recurring({"type": "hourly"})
    result = convert(master, weekly()[1])
    event = master_of(result)
    assert "RRULE" not in event and "EXDATE" not in event
    assert "unsupported recurrence pattern type 'hourly'; RRULE omitted" in result.warnings


# -- DST -----------------------------------------------------------------------------------


def test_weekly_series_keeps_local_wall_time_across_dst() -> None:
    master, _ = weekly()
    master = copy.deepcopy(master)
    master["recurrence"]["range"] = {"type": "numbered", "numberOfOccurrences": 12}
    starts = expand(convert(master))
    assert {(d.hour, d.minute) for d in starts} == {(9, 0)}
    utc_hours = [d.astimezone(UTC).hour for d in starts]
    assert utc_hours[:2] == [8, 8]  # CET, before 2026-03-29
    assert set(utc_hours[2:]) == {7}  # CEST, after 2026-03-29


# -- ISC-69: exceptions --------------------------------------------------------------------


def test_isc69_exception_becomes_recurrence_id_vevent() -> None:
    master, instances = weekly()
    result = convert(master, instances)
    excs = exceptions_of(result)
    assert len(excs) == 1
    exc = excs[0]
    assert str(exc["UID"]) == master["iCalUId"]
    assert exc["RECURRENCE-ID"].dt == datetime(2026, 4, 6, 9, 0, tzinfo=BERLIN)
    assert "RECURRENCE-ID;TZID=Europe/Berlin:20260406T090000\r\n" in unfolded(result)
    assert exc["DTSTART"].dt == datetime(2026, 4, 6, 10, 0, tzinfo=BERLIN)
    assert exc["DTEND"].dt == datetime(2026, 4, 6, 11, 0, tzinfo=BERLIN)
    assert str(exc["SUMMARY"]) == "Team sync (moved, extended)"
    assert str(exc["LOCATION"]) == "Room 2.01"
    assert "RRULE" not in exc and "EXDATE" not in exc


def test_isc69_exception_without_original_start_is_skipped() -> None:
    master, instances = weekly()
    for inst in instances:
        if inst["type"] == "exception":
            del inst["originalStart"]
    result = convert(master, instances)
    assert exceptions_of(result) == []
    assert "exception instance missing originalStart; skipped" in result.warnings


def test_isc69_recurrence_id_uses_utc_form_for_utc_series() -> None:
    master, instances = weekly()
    master["originalStartTimeZone"] = master["originalEndTimeZone"] = "UTC"
    result = convert(master, instances)
    assert "RECURRENCE-ID:20260406T070000Z\r\n" in unfolded(result)


def test_isc69_all_day_exception_recurrence_id_is_a_date() -> None:
    master = load("event_absolute_yearly.json")
    instances = [
        {"type": "occurrence", "originalStart": "2026-07-14T00:00:00Z"},
        {
            **copy.deepcopy(master),
            "type": "exception",
            "recurrence": None,
            "subject": "Bastille Day (observed)",
            # true instant of Paris midnight, as Graph reports originalStart
            "originalStart": "2027-07-13T22:00:00Z",
            "start": {"dateTime": "2027-07-15T00:00:00.0000000", "timeZone": "UTC"},
            "end": {"dateTime": "2027-07-16T00:00:00.0000000", "timeZone": "UTC"},
        },
    ]
    result = convert(master, instances, window=(WINDOW[0], datetime(2028, 1, 1, tzinfo=UTC)))
    exc = exceptions_of(result)[0]
    assert "RECURRENCE-ID;VALUE=DATE:20270714\r\n" in unfolded(result)
    assert exc["DTSTART"].dt == date(2027, 7, 15)
    assert exdates(master_of(result)) == []  # both occurrences accounted for


# -- ISC-70: cancelled occurrences ---------------------------------------------------------


def test_isc70_cancelled_occurrence_becomes_exdate() -> None:
    result = convert(*weekly())
    event = master_of(result)
    assert exdates(event) == [datetime(2026, 4, 13, 9, 0, tzinfo=BERLIN)]
    assert "EXDATE;TZID=Europe/Berlin:20260413T090000\r\n" in unfolded(result)
    assert result.warnings == ()


def test_isc70_exdate_only_inside_window() -> None:
    master, instances = weekly()
    window = (datetime(2026, 3, 1, tzinfo=UTC), datetime(2026, 4, 10, tzinfo=UTC))
    in_window = [i for i in instances if i["originalStart"] < "2026-04-10"]
    result = convert(master, in_window, window=window)
    assert exdates(master_of(result)) == []


def test_isc70_occurrence_matched_by_start_when_original_start_missing() -> None:
    master, instances = weekly()
    for inst in instances:
        if inst["type"] == "occurrence":
            del inst["originalStart"]
    assert exdates(master_of(convert(master, instances))) == [
        datetime(2026, 4, 13, 9, 0, tzinfo=BERLIN)
    ]


def test_isc70_empty_instances_emits_no_exdate_and_warns() -> None:
    master, _ = weekly()
    result = convert(master, [])
    assert exdates(master_of(result)) == []
    assert "series master has empty instances list; EXDATE derivation skipped" in result.warnings


def test_isc70_series_outside_window_does_not_warn() -> None:
    master, _ = weekly()
    window = (datetime(2030, 1, 1, tzinfo=UTC), datetime(2031, 1, 1, tzinfo=UTC))
    result = convert(master, [], window=window)
    assert exdates(master_of(result)) == []
    assert result.warnings == ()


def test_isc70_all_day_cancelled_year_becomes_date_exdate() -> None:
    master = load("event_absolute_yearly.json")
    window = (WINDOW[0], datetime(2029, 1, 1, tzinfo=UTC))
    instances = [
        {"type": "occurrence", "originalStart": "2026-07-14T00:00:00Z"},
        {"type": "occurrence", "originalStart": "2028-07-13T22:00:00Z"},  # Paris midnight
        {"type": "seriesMaster", "originalStart": "2027-07-14T00:00:00Z"},  # ignored type
    ]
    result = convert(master, instances, window=window)
    assert exdates(master_of(result)) == [date(2027, 7, 14)]
    assert "EXDATE;VALUE=DATE:20270714\r\n" in unfolded(result)


def test_isc62_unknown_series_zone_uses_recurrence_time_zone() -> None:
    master, instances = weekly()
    master["originalStartTimeZone"] = master["originalEndTimeZone"] = "tzone://Microsoft/Custom"
    result = convert(master, instances)
    assert "DTSTART;TZID=Europe/Berlin:20260316T090000\r\n" in unfolded(result)
    assert exdates(master_of(result)) == [datetime(2026, 4, 13, 9, 0, tzinfo=BERLIN)]
    assert result.warnings == (
        "time zone 'tzone://Microsoft/Custom' unknown; "
        "using recurrenceTimeZone 'W. Europe Standard Time'",
    )


def test_isc70_utc_fallback_series_only_excludes_really_missing_occurrence() -> None:
    master, instances = weekly()
    master["originalStartTimeZone"] = master["originalEndTimeZone"] = "tzone://Microsoft/Custom"
    del master["recurrence"]["range"]["recurrenceTimeZone"]
    result = convert(master, instances)
    # The series is expanded at 08:00Z; after DST Graph's instances sit at 07:00Z. Matching
    # within +-3h keeps those occurrences instead of hiding them behind EXDATEs.
    assert "DTSTART:20260316T080000Z\r\n" in unfolded(result)
    excluded = exdates(master_of(result))
    assert excluded == [datetime(2026, 4, 13, 8, 0, tzinfo=UTC)]
    assert all(d.utcoffset() == timedelta(0) for d in excluded)
    assert "EXDATE:20260413T080000Z\r\n" in unfolded(result)
    assert "unknown time zone 'tzone://Microsoft/Custom' -> UTC" in result.warnings


# -- ISC-72, ISC-73, ISC-122, ISC-123: people ----------------------------------------------


def people(result: ConvertedEvent) -> dict[str, Any]:
    event = master_of(result)
    attendees = event.get("ATTENDEE", [])
    attendees = attendees if isinstance(attendees, list) else [attendees]
    return {str(a).removeprefix("mailto:"): a.params for a in attendees}


def test_isc72_attendee_roles_and_partstats() -> None:
    result = convert(load("event_attendees.json"))
    got = people(result)
    expected = {
        "alice@example.com": ("Alice Accepted", "REQ-PARTICIPANT", "ACCEPTED"),
        "dan@example.com": ("Dan Declined", "REQ-PARTICIPANT", "DECLINED"),
        "tina@example.com": ("Tina Tentative", "OPT-PARTICIPANT", "TENTATIVE"),
        "nora@example.com": ("Nora None", "REQ-PARTICIPANT", "NEEDS-ACTION"),
        "nick@example.com": ("Nick NotResponded", "OPT-PARTICIPANT", "NEEDS-ACTION"),
        "room-aurora@example.com": ("Room Aurora", "NON-PARTICIPANT", "ACCEPTED"),
        "olga@example.com": ("Olga Organiser", "REQ-PARTICIPANT", "ACCEPTED"),
    }
    assert set(got) == set(expected)
    for address, (cn, role, partstat) in expected.items():
        params = got[address]
        assert (params["CN"], params["ROLE"], params["PARTSTAT"]) == (cn, role, partstat)
        assert "RSVP" not in params
    assert got["room-aurora@example.com"]["CUTYPE"] == "RESOURCE"
    assert all("CUTYPE" not in p for a, p in got.items() if a != "room-aurora@example.com")


def test_isc72_legacy_dn_attendee_is_skipped_with_warning() -> None:
    result = convert(load("event_attendees.json"))
    assert "/o=ExchangeLabs" not in result.ics.decode()
    assert any(w.startswith("attendee address '/o=ExchangeLabs") for w in result.warnings)


def test_isc73_organizer_with_cn() -> None:
    event = master_of(convert(load("event_attendees.json")))
    organizer = event["ORGANIZER"]
    assert str(organizer) == "mailto:olga@example.com"
    assert organizer.params["CN"] == "Olga Organiser"


@pytest.mark.parametrize(
    "fixture", ["event_attendees.json", "event_weekly_series.json", "event_single_timed.json"]
)
def test_isc122_every_attendee_and_organizer_has_schedule_agent_client(fixture: str) -> None:
    data = load(fixture)
    master, instances = (data["master"], data["instances"]) if "master" in data else (data, [])
    raw = unfolded(convert(master, instances))
    lines = [ln for ln in raw.split("\r\n") if ln.startswith(("ATTENDEE", "ORGANIZER"))]
    assert lines, "fixture must produce people"
    for line in lines:
        params = line.split(":", 1)[0]
        assert ";SCHEDULE-AGENT=CLIENT" in params, line
        assert "RSVP" not in params


def test_isc122_exception_attendees_also_carry_schedule_agent() -> None:
    exc = exceptions_of(convert(*weekly()))[0]
    assert exc["ATTENDEE"].params["SCHEDULE-AGENT"] == "CLIENT"
    assert exc["ORGANIZER"].params["SCHEDULE-AGENT"] == "CLIENT"


@pytest.mark.parametrize("fixture", ["event_attendees.json", "event_weekly_series.json"])
def test_isc123_strip_removes_attendees_and_organizer_everywhere(fixture: str) -> None:
    data = load(fixture)
    master, instances = (data["master"], data["instances"]) if "master" in data else (data, [])
    result = convert(master, instances, attendees="strip")
    raw = unfolded(result)
    assert "ATTENDEE" not in raw
    assert "ORGANIZER" not in raw
    assert "mailto:" not in raw


def test_isc123_invalid_attendees_mode_raises() -> None:
    with pytest.raises(ValueError, match="attendees"):
        convert_event(load("event_single_timed.json"), [], window=WINDOW, attendees="drop")


# -- ISC-74..78: remaining fields ----------------------------------------------------------


def test_isc74_reminder_becomes_valarm() -> None:
    result = convert(load("event_single_timed.json"))
    alarms = master_of(result).walk("VALARM")
    assert len(alarms) == 1
    alarm = alarms[0]
    assert str(alarm["ACTION"]) == "DISPLAY"
    assert str(alarm["DESCRIPTION"]) == "Quarterly review"
    assert alarm["TRIGGER"].dt == timedelta(minutes=-15)
    assert "TRIGGER:-PT15M\r\n" in unfolded(result)


def test_isc74_no_valarm_when_reminder_off() -> None:
    master = load("event_single_timed.json")
    master["isReminderOn"] = False
    assert master_of(convert(master)).walk("VALARM") == []


@pytest.mark.parametrize(
    ("sensitivity", "klass"),
    [("normal", "PUBLIC"), ("personal", "PUBLIC"), ("private", "PRIVATE"),
     ("confidential", "CONFIDENTIAL")],
)
def test_isc75_sensitivity_to_class(sensitivity: str, klass: str) -> None:
    master = load("event_single_timed.json")
    master["sensitivity"] = sensitivity
    assert str(master_of(convert(master))["CLASS"]) == klass


@pytest.mark.parametrize(
    ("show_as", "transp", "busy"),
    [
        ("free", "TRANSPARENT", "FREE"),
        ("tentative", "OPAQUE", "TENTATIVE"),
        ("busy", "OPAQUE", "BUSY"),
        ("oof", "OPAQUE", "OOF"),
        ("workingElsewhere", "OPAQUE", "WORKINGELSEWHERE"),
        ("unknown", "OPAQUE", None),
    ],
)
def test_isc76_show_as_to_transp_and_busystatus(
    show_as: str, transp: str, busy: str | None
) -> None:
    master = load("event_single_timed.json")
    master["showAs"] = show_as
    event = master_of(convert(master))
    assert str(event["TRANSP"]) == transp
    if busy is None:
        assert "X-MICROSOFT-CDO-BUSYSTATUS" not in event
    else:
        assert str(event["X-MICROSOFT-CDO-BUSYSTATUS"]) == busy


def test_isc77_location_text_description_join_url_and_categories() -> None:
    master = load("event_single_timed.json")
    event = master_of(convert(master))
    assert str(event["LOCATION"]) == "Room 4.12, HQ"
    assert str(event["DESCRIPTION"]) == (
        "Agenda: numbers, plans; risks\nBring laptop.\n\nJoin: "
        + master["onlineMeeting"]["joinUrl"]
    )
    assert "X-ALT-DESC" not in event
    assert list(event["CATEGORIES"].cats) == ["Blue category", "Finance"]
    assert master["webLink"] not in str(event["DESCRIPTION"])


def test_isc77_html_body_rendered_to_text_and_preserved() -> None:
    master = load("event_html_body.json")
    event = master_of(convert(master))
    text = str(event["DESCRIPTION"])
    assert "<" not in text and ">" not in text
    assert "alert" not in text and "margin" not in text  # script/style content dropped
    assert "Hello team," in text
    assert "Fish & chips; café at 12:00." in text
    assert "Second line" in text  # whitespace collapsed
    assert "One\n" in text and "Two" in text
    alt = event["X-ALT-DESC"]
    assert alt.params["FMTTYPE"] == "text/html"
    assert str(alt) == master["body"]["content"]


def test_isc77_x_alt_desc_capped_at_64_kib() -> None:
    master = load("event_html_body.json")
    master["body"]["content"] = "<p>" + "ä" * 40_000 + "</p>"  # 80 kB of UTF-8
    result = convert(master)
    alt = str(master_of(result)["X-ALT-DESC"])
    assert len(alt.encode("utf-8")) <= 64 * 1024
    assert alt.startswith("<p>ää")
    assert "X-ALT-DESC exceeded 64KiB and was truncated" in result.warnings


def test_isc77_body_preview_used_when_body_missing() -> None:
    master = load("event_single_timed.json")
    del master["body"], master["onlineMeeting"]
    assert str(master_of(convert(master))["DESCRIPTION"]) == master["bodyPreview"]


def test_isc77_no_description_when_body_empty() -> None:
    event = master_of(convert(load("event_attendees.json")))
    assert "DESCRIPTION" not in event


def test_isc78_cancelled_event_status() -> None:
    assert str(master_of(convert(load("event_html_body.json")))["STATUS"]) == "CANCELLED"
    assert "STATUS" not in master_of(convert(load("event_single_timed.json")))


def test_summary_fallback_and_timestamps() -> None:
    master = load("event_single_timed.json")
    master["subject"] = ""
    event = master_of(convert(master))
    assert str(event["SUMMARY"]) == "(no subject)"
    assert event["CREATED"].dt == datetime(2026, 2, 1, 10, 15, 30, tzinfo=UTC)  # 7-digit fraction
    assert event["LAST-MODIFIED"].dt.date() == date(2026, 2, 10)
    assert event["DTSTAMP"].dt.tzinfo is not None


# -- ISC-79 and serialisation --------------------------------------------------------------


@pytest.mark.parametrize(
    "fixture",
    ["event_single_timed.json", "event_all_day.json", "event_weekly_series.json",
     "event_relative_monthly_last_friday.json", "event_absolute_yearly.json",
     "event_attendees.json", "event_html_body.json"],
)
def test_isc79_every_fixture_round_trips_with_vtimezones(fixture: str) -> None:
    data = load(fixture)
    master, instances = (data["master"], data["instances"]) if "master" in data else (data, [])
    result = convert(master, instances)  # convert() asserts parse + VTIMEZONE coverage
    cal = Calendar.from_ical(result.ics)
    assert cal.walk("VTIMEZONE") == [] or cal.subcomponents[0].name == "VTIMEZONE"
    assert Calendar.from_ical(cal.to_ical()).to_ical() == cal.to_ical()


def test_isc79_vtimezone_covers_series_start_and_until() -> None:
    master = load("event_relative_monthly_last_friday.json")
    master["start"]["dateTime"] = "2019-01-25T18:00:00.0000000"
    master["end"]["dateTime"] = "2019-01-25T20:00:00.0000000"
    cal = Calendar.from_ical(convert(master).ics)
    tz = cal.walk("VTIMEZONE")[0]
    first = min(sub["DTSTART"].dt for sub in tz.subcomponents)
    assert first.year <= 2018


def test_text_escaping_round_trips() -> None:
    master = load("event_single_timed.json")
    master["subject"] = "Budget, Q3; final\\draft"
    master["body"] = {"contentType": "text", "content": "line1\nline2, with; stuff"}
    master["onlineMeeting"] = None
    result = convert(master)
    raw = unfolded(result)
    assert "SUMMARY:Budget\\, Q3\\; final\\\\draft\r\n" in raw
    event = master_of(result)
    assert str(event["SUMMARY"]) == "Budget, Q3; final\\draft"
    assert str(event["DESCRIPTION"]) == "line1\nline2, with; stuff"


def test_long_lines_are_folded_to_75_octets() -> None:
    master = load("event_single_timed.json")
    master["body"] = {"contentType": "text", "content": "Grüße aus Köln € 😀 " * 40}
    result = convert(master)
    physical = result.ics.split(b"\r\n")
    assert any(line.startswith(b" ") for line in physical)
    assert all(len(line) <= 75 for line in physical)
    assert str(master_of(result)["DESCRIPTION"]).startswith("Grüße aus Köln € 😀 ")


# -- input validation and robustness -------------------------------------------------------


def test_malformed_optional_fields_become_warnings() -> None:
    master, instances = weekly()
    master = copy.deepcopy(master)
    master["attendees"] = "x"
    master["categories"] = None
    master["reminderMinutesBeforeStart"] = "abc"
    master["isReminderOn"] = True
    master["createdDateTime"] = "garbage"
    master["location"] = ["not", "a", "dict"]
    master["organizer"] = {"emailAddress": {"address": None}}
    master["recurrence"]["pattern"]["daysOfWeek"] = 5
    result = convert(master, instances)
    event = master_of(result)
    assert "RRULE" not in event and "EXDATE" not in event
    assert "ATTENDEE" not in event and "ORGANIZER" not in event
    assert "CREATED" not in event and "LOCATION" not in event
    assert event.walk("VALARM") == []
    for expected in (
        "invalid attendees; expected list",
        "invalid categories; expected list[str]",
        "invalid reminderMinutesBeforeStart; VALARM omitted",
        "invalid createdDateTime; expected datetime string",
        "invalid recurrence.pattern.daysOfWeek; RRULE omitted",
        "organizer address 'None' is invalid; skipped",
    ):
        assert expected in result.warnings


def test_malformed_instances_are_tolerated() -> None:
    master, instances = weekly()
    instances = [*instances, "junk", {"type": "exception", "originalStart": "garbage"}]
    result = convert(master, instances)  # type: ignore[list-item]
    assert len(exceptions_of(result)) == 1
    assert "instance entry is not an object; skipped" in result.warnings
    assert "exception instance has invalid originalStart 'garbage'; skipped" in result.warnings


def test_exception_with_unparseable_start_is_skipped() -> None:
    master, instances = weekly()
    for inst in instances:
        if inst["type"] == "exception":
            inst["start"] = {"dateTime": "garbage"}
    result = convert(master, instances)
    assert exceptions_of(result) == []
    assert "exception instance has invalid start/end; skipped" in result.warnings


@pytest.mark.parametrize("field", ["start", "end"])
def test_missing_start_or_end_raises(field: str) -> None:
    master = load("event_single_timed.json")
    del master[field]
    with pytest.raises(ValueError, match="start/end"):
        convert_event(master, [], window=WINDOW)


def test_unparseable_start_raises() -> None:
    master = load("event_single_timed.json")
    master["start"]["dateTime"] = "garbage"
    with pytest.raises(ValueError):
        convert_event(master, [], window=WINDOW)


@pytest.mark.parametrize("kind", ["occurrence", "exception"])
def test_isc105_occurrence_or_exception_as_master_raises(kind: str) -> None:
    master = load("event_single_timed.json")
    master["type"] = kind
    with pytest.raises(ValueError, match="singleInstance or seriesMaster"):
        convert_event(master, [], window=WINDOW)


def test_naive_window_raises() -> None:
    with pytest.raises(ValueError, match="timezone-aware"):
        convert_event(
            load("event_single_timed.json"), [],
            window=(datetime(2026, 1, 1), datetime(2027, 1, 1)),  # noqa: DTZ001
        )


def test_master_without_any_id_raises() -> None:
    master = load("event_single_timed.json")
    del master["iCalUId"], master["id"]
    with pytest.raises(ValueError, match="iCalUId or id"):
        convert_event(master, [], window=WINDOW)
