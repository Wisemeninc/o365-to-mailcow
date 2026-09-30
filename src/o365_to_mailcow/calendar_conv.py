"""Graph event -> iCalendar conversion.

Graph event payloads are fetched with ``Prefer: outlook.timezone=\"UTC\"``. That means
``start.dateTime`` and ``end.dateTime`` are UTC wall-time strings while
``originalStartTimeZone`` / ``originalEndTimeZone`` keep the organizer's intended zone.
This converter therefore parses the UTC instant first and then re-zones DTSTART/DTEND to
the original zone so recurrence expansion keeps the original wall time across DST.

Recurrence range ``endDate`` is converted to an inclusive iCalendar ``UNTIL``. For timed
events, ``UNTIL`` is computed from ``endDate`` combined with the DTSTART local wall time,
then converted to UTC. That includes an occurrence on ``endDate`` itself.

Cancelled occurrences are inferred by expanding the master's RRULE inside the caller's
window and comparing expected starts to Graph ``/instances`` rows (``originalStart``,
within +-3h). Missing starts become ``EXDATE`` in the same DTSTART form. If Graph returns
no instances for a series master that should have occurrences in the window, EXDATE
derivation is skipped: an all-cancelled series is otherwise indistinguishable from a
failed fetch.

A series whose own zone is unknown ("tzone://Microsoft/Custom") is anchored to
``recurrence.range.recurrenceTimeZone`` when that names a real zone; only if neither is
usable does it fall back to UTC (with a warning).

All ATTENDEE and ORGANIZER properties carry ``SCHEDULE-AGENT=CLIENT`` so SOGo treats PUTs
as passive imports and does not send invitations.

Graph ``relativeMonthly`` / ``relativeYearly`` with multiple ``daysOfWeek`` are expressed
as ``BYDAY=<days>;BYSETPOS=<index>``. With one day we emit ``BYDAY=<index><day>``.

ISC-105 note: callers must pass only ``singleInstance`` or ``seriesMaster`` masters.
Passing ``occurrence`` / ``exception`` as ``master`` is a ValueError.
"""

from __future__ import annotations

import re
from bisect import bisect_left
from dataclasses import dataclass
from datetime import UTC, date, datetime, time, timedelta
from html.parser import HTMLParser
from typing import Any
from uuid import NAMESPACE_URL, uuid5
from zoneinfo import ZoneInfoNotFoundError

from dateutil.rrule import rrulestr
from icalendar import Alarm, Calendar, Event, Timezone, vCalAddress

from .tz import windows_to_iana, zone

_PRODID = "-//o365-to-mailcow//EN"
_MATCH_TOLERANCE = timedelta(hours=3)
_DAY_TOKEN = {
    "sunday": "SU",
    "monday": "MO",
    "tuesday": "TU",
    "wednesday": "WE",
    "thursday": "TH",
    "friday": "FR",
    "saturday": "SA",
}
_INDEX_TOKEN = {"first": 1, "second": 2, "third": 3, "fourth": 4, "last": -1}
_SHOW_AS_BUSY = {
    "free": "FREE",
    "tentative": "TENTATIVE",
    "busy": "BUSY",
    "oof": "OOF",
    "workingelsewhere": "WORKINGELSEWHERE",
}
_SENSITIVITY_CLASS = {
    "normal": "PUBLIC",
    "personal": "PUBLIC",
    "private": "PRIVATE",
    "confidential": "CONFIDENTIAL",
}
_PARTSTAT = {
    "none": "NEEDS-ACTION",
    "notresponded": "NEEDS-ACTION",
    "accepted": "ACCEPTED",
    "declined": "DECLINED",
    "tentativelyaccepted": "TENTATIVE",
    "organizer": "ACCEPTED",
}


@dataclass(frozen=True)
class ConvertedEvent:
    """One ``.ics`` resource: master VEVENT, exception VEVENTs and their VTIMEZONEs."""

    uid: str
    ics: bytes
    last_modified: str | None
    warnings: tuple[str, ...]


@dataclass(frozen=True)
class _EventSpan:
    is_all_day: bool
    dtstart_value: datetime | date
    dtend_value: datetime | date
    dtstart_utc: datetime
    dtend_utc: datetime
    source_zone: str | None  # raw Graph originalStartTimeZone, used to date all-day instants


class _WarningSink:
    def __init__(self) -> None:
        self._items: list[str] = []
        self._unknown_tz_once: set[str] = set()

    def warn(self, message: str) -> None:
        self._items.append(message)

    def warn_unknown_tz(self, raw_name: object) -> None:
        if isinstance(raw_name, str) and raw_name.strip():
            shown = raw_name.strip()
        else:
            shown = "<missing>"
        if shown in self._unknown_tz_once:
            return
        self._unknown_tz_once.add(shown)
        self.warn(f"unknown time zone '{shown}' -> UTC")

    def as_tuple(self) -> tuple[str, ...]:
        return tuple(self._items)


class _GraphHtmlToText(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._out: list[str] = []
        self._skip_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        lowered = tag.lower()
        if lowered in {"script", "style"}:
            self._skip_depth += 1
            return
        if self._skip_depth:
            return
        if lowered in {"br", "p", "div", "li", "tr"}:
            self._out.append("\n")

    def handle_endtag(self, tag: str) -> None:
        lowered = tag.lower()
        if lowered in {"script", "style"} and self._skip_depth:
            self._skip_depth -= 1
            return
        if self._skip_depth:
            return
        if lowered in {"p", "div", "li", "tr"}:
            self._out.append("\n")

    def handle_data(self, data: str) -> None:
        if self._skip_depth:
            return
        self._out.append(data)

    def text(self) -> str:
        value = "".join(self._out)
        value = value.replace("\r\n", "\n").replace("\r", "\n").replace("\xa0", " ")
        value = re.sub(r"[\t ]+", " ", value)
        value = re.sub(r"\n{3,}", "\n\n", value)
        return value.strip()


def convert_event(
    master: dict,
    instances: list[dict],
    *,
    window: tuple[datetime, datetime],
    attendees: str = "keep",
) -> ConvertedEvent:
    """Convert one Graph master event and its instances into one VCALENDAR payload.

    ``instances`` must come from ``/events/{id}/instances`` over the same ``window``; it is
    the only source of exceptions and cancelled occurrences. ``attendees="strip"`` drops
    ATTENDEE and ORGANIZER entirely. Unknown-zone warnings are deduplicated per event;
    per-run deduplication (ISC-62) is the caller's job.

    Raises ``ValueError`` for an invalid ``attendees`` mode, a naive ``window``, a master
    of type occurrence/exception, a master without ``iCalUId`` and ``id``, or a master
    whose ``start``/``end`` is missing or unparseable. Other malformed fields become
    warnings.
    """
    if attendees not in {"keep", "strip"}:
        raise ValueError("attendees must be 'keep' or 'strip'")
    if not _is_aware(window[0]) or not _is_aware(window[1]):
        raise ValueError("window datetimes must be timezone-aware")
    if not isinstance(master, dict):
        raise ValueError("master must be a dict")

    master_type = _lower_str(master.get("type"))
    if master_type in {"occurrence", "exception"}:
        raise ValueError("master event type must be singleInstance or seriesMaster")
    if not _non_empty_str(master.get("iCalUId")) and not _non_empty_str(master.get("id")):
        raise ValueError("master must include iCalUId or id")
    if not isinstance(instances, list):
        raise ValueError("instances must be a list")

    warnings = _WarningSink()
    uid = _uid_for(master)

    master = _with_recurrence_zone_fallback(master, warnings=warnings)
    master_span = _span_from_event(master, warnings=warnings, strict=True, label="master")
    if master_span is None:  # strict=True raises instead; kept for the type checker
        raise ValueError("master event is missing parseable start/end")

    master_event = Event()
    _apply_event_fields(
        master_event,
        source=master,
        uid=uid,
        span=master_span,
        attendees_mode=attendees,
        warnings=warnings,
    )

    rrule_map, rrule_text, latest_until = _rrule_from_master(
        master, span=master_span, warnings=warnings
    )
    if rrule_map:
        master_event.add("rrule", rrule_map)

    exdate_values: list[datetime | date] = []
    if master_type == "seriesmaster" and rrule_map:
        exdate_values = _derive_exdates(
            span=master_span,
            rrule_text=rrule_text,
            instances=instances,
            window=window,
            warnings=warnings,
        )
    if exdate_values and master_span.is_all_day:
        # icalendar omits VALUE=DATE on date lists; RFC 5545 defaults EXDATE to DATE-TIME.
        master_event.add("exdate", exdate_values, parameters={"VALUE": "DATE"})
    elif exdate_values:
        master_event.add("exdate", exdate_values)

    exception_events: list[Event] = []
    if master_type == "seriesmaster":
        for entry in instances:
            if not isinstance(entry, dict):
                warnings.warn("instance entry is not an object; skipped")
                continue
            if _lower_str(entry.get("type")) != "exception":
                continue
            original_start = entry.get("originalStart")
            if not _non_empty_str(original_start):
                warnings.warn("exception instance missing originalStart; skipped")
                continue
            recurrence_id = _recurrence_id_from_original_start(
                str(original_start),
                span=master_span,
                warnings=warnings,
            )
            if recurrence_id is None:
                continue
            exc_span = _span_from_event(entry, warnings=warnings, strict=False, label="exception")
            if exc_span is None:
                warnings.warn("exception instance has invalid start/end; skipped")
                continue
            exc_event = Event()
            _apply_event_fields(
                exc_event,
                source=entry,
                uid=uid,
                span=exc_span,
                attendees_mode=attendees,
                warnings=warnings,
            )
            exc_event.add("recurrence-id", recurrence_id)
            exception_events.append(exc_event)

    all_events = [master_event, *exception_events]
    used_tzids = _collect_tzids(all_events)

    emitted_datetimes = _collect_emitted_datetimes(all_events)
    latest = max(
        [_as_utc(window[1]), *(_as_utc(value) for value in emitted_datetimes)],
        default=_as_utc(window[1]),
    )
    if latest_until is not None:
        latest = max(latest, _as_utc(latest_until))
    earliest = min((_as_utc(value) for value in emitted_datetimes), default=_as_utc(window[0]))

    calendar = Calendar()
    calendar.add("version", "2.0")
    calendar.add("prodid", _PRODID)
    calendar.add("calscale", "GREGORIAN")

    for tzid in sorted(used_tzids):
        tz_component = Timezone.from_tzinfo(
            zone(tzid),
            first_date=date(earliest.year - 1, 1, 1),
            last_date=date(latest.year + 2, 1, 1),
        )
        calendar.add_component(tz_component)

    for component in all_events:
        calendar.add_component(component)

    last_modified = master.get("lastModifiedDateTime")
    return ConvertedEvent(
        uid=uid,
        ics=calendar.to_ical(),
        last_modified=last_modified if isinstance(last_modified, str) else None,
        warnings=warnings.as_tuple(),
    )


def _with_recurrence_zone_fallback(
    master: dict[str, Any], *, warnings: _WarningSink
) -> dict[str, Any]:
    """Use ``recurrence.range.recurrenceTimeZone`` when a series' own zone is unknown.

    Organiser-defined zones arrive as "tzone://Microsoft/Custom", but the range usually
    still names a real zone. Expanding a series in UTC instead would shift every occurrence
    on the far side of a DST change by an hour, so the named zone is the better anchor.
    """
    if windows_to_iana(_non_empty_str(master.get("originalStartTimeZone"))) is not None:
        return master
    recurrence = master.get("recurrence")
    range_obj = recurrence.get("range") if isinstance(recurrence, dict) else None
    fallback = range_obj.get("recurrenceTimeZone") if isinstance(range_obj, dict) else None
    if not isinstance(fallback, str) or windows_to_iana(fallback) is None:
        return master
    original = master.get("originalStartTimeZone")
    warnings.warn(f"time zone '{original}' unknown; using recurrenceTimeZone '{fallback}'")
    patched = dict(master)
    patched["originalStartTimeZone"] = fallback
    if windows_to_iana(_non_empty_str(master.get("originalEndTimeZone"))) is None:
        patched["originalEndTimeZone"] = fallback
    return patched


def _uid_for(master: dict[str, Any]) -> str:
    uid = _non_empty_str(master.get("iCalUId"))
    if uid:
        return uid
    graph_id = _non_empty_str(master.get("id"))
    if not graph_id:
        raise ValueError("master must include id when iCalUId is missing")
    return str(uuid5(NAMESPACE_URL, f"o365-to-mailcow:event:{graph_id}"))


def _span_from_event(
    event: dict[str, Any],
    *,
    warnings: _WarningSink,
    strict: bool,
    label: str,
) -> _EventSpan | None:
    start = _parse_start_end_datetime(event.get("start"))
    end = _parse_start_end_datetime(event.get("end"))
    if start is None or end is None:
        if strict:
            raise ValueError(f"{label} event is missing parseable start/end")
        return None

    raw_zone = event.get("originalStartTimeZone")
    source_zone = raw_zone if isinstance(raw_zone, str) else None
    is_all_day = bool(event.get("isAllDay"))
    if is_all_day:
        start_date = _all_day_date(start, source_zone)
        end_date = _all_day_date(end, source_zone)
        if end_date <= start_date:
            end_date = start_date + timedelta(days=1)
            warnings.warn("all-day event had non-exclusive end; forced DTEND to DTSTART + 1 day")
        return _EventSpan(
            is_all_day=True,
            dtstart_value=start_date,
            dtend_value=end_date,
            dtstart_utc=datetime.combine(start_date, time.min, tzinfo=UTC),
            dtend_utc=datetime.combine(end_date, time.min, tzinfo=UTC),
            source_zone=source_zone,
        )

    start_tz = _resolve_iana(event.get("originalStartTimeZone"), warnings=warnings)
    end_source = event.get("originalEndTimeZone")
    if not _non_empty_str(end_source):
        end_source = event.get("originalStartTimeZone")
    end_tz = _resolve_iana(end_source, warnings=warnings)

    start_value = start.astimezone(zone(start_tz)) if start_tz != "UTC" else start
    end_value = end.astimezone(zone(end_tz)) if end_tz != "UTC" else end

    return _EventSpan(
        is_all_day=False,
        dtstart_value=start_value,
        dtend_value=end_value,
        dtstart_utc=start,
        dtend_utc=end,
        source_zone=source_zone,
    )


def _all_day_date(value_utc: datetime, zone_name: str | None) -> date:
    """Return the calendar date an all-day Graph instant stands for.

    With the UTC Prefer header Graph usually reports all-day boundaries as floating
    midnights ("2026-05-01T00:00:00" UTC), whose date part is the answer. Some fields
    (notably an instance's ``originalStart``) can instead carry the true instant of local
    midnight, e.g. "2026-04-30T22:00:00Z" for Berlin. Those are converted to the event's
    original zone when it is known and lands on midnight there; otherwise the instant is
    rounded to the nearest midnight, which is correct for every UTC offset within +-12h.
    """
    if value_utc.time() == time.min:
        return value_utc.date()
    iana = windows_to_iana(zone_name)
    if iana is not None and iana != "UTC":
        try:
            local = value_utc.astimezone(zone(iana))
        except (ZoneInfoNotFoundError, ValueError):
            local = None
        if local is not None and local.time() == time.min:
            return local.date()
    return (value_utc + timedelta(hours=12)).date()


def _resolve_iana(raw_name: object, *, warnings: _WarningSink) -> str:
    mapped = windows_to_iana(raw_name if isinstance(raw_name, str) else None)
    if mapped is None:
        warnings.warn_unknown_tz(raw_name)
        return "UTC"
    if mapped == "UTC":
        return "UTC"
    try:
        zone(mapped)
    except (ZoneInfoNotFoundError, ValueError):
        # The CLDR table can name a zone the installed tzdata lacks; treat as unknown.
        warnings.warn_unknown_tz(raw_name)
        return "UTC"
    return mapped


def _parse_start_end_datetime(value: object) -> datetime | None:
    if not isinstance(value, dict):
        return None
    raw = value.get("dateTime")
    if not isinstance(raw, str) or not raw.strip():
        return None
    try:
        return _parse_graph_datetime(raw)
    except ValueError:
        return None


def _parse_graph_datetime(raw: str) -> datetime:
    text = raw.strip()
    text = re.sub(r"\.(\d{6})\d+(?=(?:Z|[+-]\d\d:\d\d)?$)", r".\1", text)
    if text.endswith("Z"):
        text = f"{text[:-1]}+00:00"
    dt = datetime.fromisoformat(text)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt.astimezone(UTC)


def _apply_event_fields(
    target: Event,
    *,
    source: dict[str, Any],
    uid: str,
    span: _EventSpan,
    attendees_mode: str,
    warnings: _WarningSink,
) -> None:
    target.add("uid", uid)
    target.add("dtstart", span.dtstart_value)
    target.add("dtend", span.dtend_value)
    target.add("dtstamp", datetime.now(tz=UTC))

    created = _parse_optional_datetime(
        source.get("createdDateTime"), field="createdDateTime", warnings=warnings
    )
    if created is not None:
        target.add("created", created)
    last_modified = _parse_optional_datetime(
        source.get("lastModifiedDateTime"),
        field="lastModifiedDateTime",
        warnings=warnings,
    )
    if last_modified is not None:
        target.add("last-modified", last_modified)

    summary = source.get("subject") if isinstance(source.get("subject"), str) else None
    target.add("summary", summary if summary else "(no subject)")

    class_value = _SENSITIVITY_CLASS.get(_lower_str(source.get("sensitivity")), "PUBLIC")
    target.add("class", class_value)

    show_as = _lower_str(source.get("showAs"))
    target.add("transp", "TRANSPARENT" if show_as == "free" else "OPAQUE")
    if show_as in _SHOW_AS_BUSY:
        target.add("X-MICROSOFT-CDO-BUSYSTATUS", _SHOW_AS_BUSY[show_as])

    if bool(source.get("isCancelled")):
        target.add("status", "CANCELLED")

    location_name = _extract_location(source.get("location"))
    if location_name:
        target.add("location", location_name)

    body_text, body_html = _description_fields(source, warnings=warnings)
    join_url = _extract_join_url(source.get("onlineMeeting"))
    if join_url:
        body_text = f"{body_text}\n\nJoin: {join_url}" if body_text else f"Join: {join_url}"
    if body_text:
        target.add("description", body_text)
    if body_html:
        trimmed = _cap_utf8(body_html, cap=64 * 1024)
        if trimmed != body_html:
            warnings.warn("X-ALT-DESC exceeded 64KiB and was truncated")
        target.add("X-ALT-DESC", trimmed, parameters={"FMTTYPE": "text/html"})

    categories = source.get("categories")
    if categories is None:
        if "categories" in source:
            warnings.warn("invalid categories; expected list[str]")
    elif isinstance(categories, list):
        values = [item for item in categories if isinstance(item, str) and item]
        if values:
            target.add("categories", values)
    else:
        warnings.warn("invalid categories; expected list[str]")

    _add_alarm(
        target, source=source, summary=summary if summary else "(no subject)", warnings=warnings
    )
    _add_people(target, source=source, attendees_mode=attendees_mode, warnings=warnings)


def _extract_location(value: object) -> str | None:
    if not isinstance(value, dict):
        return None
    display = value.get("displayName")
    if isinstance(display, str) and display.strip():
        return display.strip()
    return None


def _extract_join_url(value: object) -> str | None:
    if not isinstance(value, dict):
        return None
    join = value.get("joinUrl")
    if isinstance(join, str) and join.strip():
        return join.strip()
    return None


def _description_fields(
    source: dict[str, Any], *, warnings: _WarningSink
) -> tuple[str | None, str | None]:
    body = source.get("body")
    body_preview = source.get("bodyPreview")
    if isinstance(body, dict):
        content = body.get("content")
        content_type = _lower_str(body.get("contentType"))
        if isinstance(content, str) and content:
            if content_type == "text":
                return content, None
            if content_type == "html":
                parser = _GraphHtmlToText()
                parser.feed(content)
                parser.close()
                text = parser.text()
                if not text and isinstance(body_preview, str) and body_preview:
                    text = body_preview
                return text if text else None, content
            warnings.warn("body.contentType is invalid; falling back to bodyPreview")
    if isinstance(body_preview, str) and body_preview:
        return body_preview, None
    return None, None


def _cap_utf8(value: str, *, cap: int) -> str:
    encoded = value.encode("utf-8")
    if len(encoded) <= cap:
        return value
    chunk = encoded[:cap]
    while chunk:
        try:
            return chunk.decode("utf-8")
        except UnicodeDecodeError as exc:
            chunk = chunk[:exc.start]
    return ""


def _add_alarm(
    target: Event, *, source: dict[str, Any], summary: str, warnings: _WarningSink
) -> None:
    if not bool(source.get("isReminderOn")):
        return
    minutes = source.get("reminderMinutesBeforeStart")
    if not isinstance(minutes, int) or isinstance(minutes, bool) or minutes < 0:
        warnings.warn("invalid reminderMinutesBeforeStart; VALARM omitted")
        return
    alarm = Alarm()
    alarm.add("action", "DISPLAY")
    alarm.add("description", summary)
    alarm.add("trigger", timedelta(minutes=-minutes))
    target.add_component(alarm)


def _add_people(
    target: Event,
    *,
    source: dict[str, Any],
    attendees_mode: str,
    warnings: _WarningSink,
) -> None:
    if attendees_mode == "strip":
        return

    organizer = source.get("organizer")
    if isinstance(organizer, dict):
        email = organizer.get("emailAddress")
        if isinstance(email, dict):
            _add_organizer(target, email, warnings=warnings)

    attendees = source.get("attendees")
    if attendees is None:
        return
    if not isinstance(attendees, list):
        warnings.warn("invalid attendees; expected list")
        return
    for attendee in attendees:
        if not isinstance(attendee, dict):
            warnings.warn("attendee entry is not an object; skipped")
            continue
        email = attendee.get("emailAddress")
        if not isinstance(email, dict):
            warnings.warn("attendee missing emailAddress; skipped")
            continue
        address = email.get("address")
        if not isinstance(address, str) or "@" not in address:
            warnings.warn(f"attendee address '{address}' is invalid; skipped")
            continue
        cal_address = vCalAddress(f"mailto:{address}")
        name = email.get("name")
        if isinstance(name, str) and name:
            cal_address.params["CN"] = name

        attendee_type = _lower_str(attendee.get("type"))
        role = "REQ-PARTICIPANT"
        if attendee_type == "optional":
            role = "OPT-PARTICIPANT"
        elif attendee_type == "resource":
            role = "NON-PARTICIPANT"
            cal_address.params["CUTYPE"] = "RESOURCE"
        cal_address.params["ROLE"] = role

        status = attendee.get("status")
        response = _lower_str(status.get("response") if isinstance(status, dict) else None)
        cal_address.params["PARTSTAT"] = _PARTSTAT.get(response, "NEEDS-ACTION")
        cal_address.params["SCHEDULE-AGENT"] = "CLIENT"
        target.add("attendee", cal_address)


def _add_organizer(target: Event, email: dict[str, Any], *, warnings: _WarningSink) -> None:
    address = email.get("address")
    if not isinstance(address, str) or "@" not in address:
        warnings.warn(f"organizer address '{address}' is invalid; skipped")
        return
    value = vCalAddress(f"mailto:{address}")
    name = email.get("name")
    if isinstance(name, str) and name:
        value.params["CN"] = name
    value.params["SCHEDULE-AGENT"] = "CLIENT"
    target.add("organizer", value)


def _rrule_from_master(
    master: dict[str, Any],
    *,
    span: _EventSpan,
    warnings: _WarningSink,
) -> tuple[dict[str, Any] | None, str | None, datetime | None]:
    recurrence = master.get("recurrence")
    if recurrence is None:
        return None, None, None
    if not isinstance(recurrence, dict):
        warnings.warn("invalid recurrence; RRULE omitted")
        return None, None, None
    pattern = recurrence.get("pattern")
    if not isinstance(pattern, dict):
        warnings.warn("invalid recurrence.pattern; RRULE omitted")
        return None, None, None

    interval = _positive_int(pattern.get("interval"), default=1)
    pattern_type = _lower_str(pattern.get("type"))
    rule: dict[str, Any] = {"FREQ": None, "INTERVAL": interval}

    if pattern_type == "daily":
        rule["FREQ"] = "DAILY"
    elif pattern_type == "weekly":
        days = _days_from_graph(pattern.get("daysOfWeek"))
        if not days:
            warnings.warn("invalid recurrence.pattern.daysOfWeek; RRULE omitted")
            return None, None, None
        rule["FREQ"] = "WEEKLY"
        rule["BYDAY"] = days
        rule["WKST"] = _DAY_TOKEN.get(_lower_str(pattern.get("firstDayOfWeek")), "MO")
    elif pattern_type == "absolutemonthly":
        day_of_month = _bounded_int(pattern.get("dayOfMonth"), min_value=1, max_value=31)
        if day_of_month is None:
            warnings.warn("invalid recurrence.pattern.dayOfMonth; RRULE omitted")
            return None, None, None
        rule["FREQ"] = "MONTHLY"
        rule.update(_monthday_rule(day_of_month))
    elif pattern_type == "relativemonthly":
        relative = _relative_byday(pattern=pattern, warnings=warnings)
        if relative is None:
            return None, None, None
        rule["FREQ"] = "MONTHLY"
        rule.update(relative)
    elif pattern_type == "absoluteyearly":
        month = _bounded_int(pattern.get("month"), min_value=1, max_value=12)
        day_of_month = _bounded_int(pattern.get("dayOfMonth"), min_value=1, max_value=31)
        if month is None or day_of_month is None:
            warnings.warn("invalid recurrence.pattern.month/dayOfMonth; RRULE omitted")
            return None, None, None
        rule["FREQ"] = "YEARLY"
        rule["BYMONTH"] = month
        rule.update(_monthday_rule(day_of_month))
    elif pattern_type == "relativeyearly":
        month = _bounded_int(pattern.get("month"), min_value=1, max_value=12)
        if month is None:
            warnings.warn("invalid recurrence.pattern.month; RRULE omitted")
            return None, None, None
        relative = _relative_byday(pattern=pattern, warnings=warnings)
        if relative is None:
            return None, None, None
        rule["FREQ"] = "YEARLY"
        rule["BYMONTH"] = month
        rule.update(relative)
    else:
        warnings.warn(f"unsupported recurrence pattern type '{pattern_type}'; RRULE omitted")
        return None, None, None

    latest_until: datetime | None = None
    range_obj = recurrence.get("range")
    if range_obj is not None:
        if not isinstance(range_obj, dict):
            warnings.warn("invalid recurrence.range; RRULE end omitted")
        else:
            range_type = _lower_str(range_obj.get("type"))
            if range_type == "enddate":
                until_date = _parse_iso_date(range_obj.get("endDate"))
                if until_date is None:
                    warnings.warn("invalid recurrence.range.endDate; RRULE end omitted")
                elif span.is_all_day:
                    rule["UNTIL"] = until_date
                elif isinstance(span.dtstart_value, datetime):
                    start_local = span.dtstart_value
                    until_local = datetime(
                        year=until_date.year,
                        month=until_date.month,
                        day=until_date.day,
                        hour=start_local.hour,
                        minute=start_local.minute,
                        second=start_local.second,
                        microsecond=start_local.microsecond,
                        tzinfo=start_local.tzinfo,
                    )
                    until_utc = until_local.astimezone(UTC)
                    latest_until = until_utc
                    rule["UNTIL"] = until_utc
            elif range_type == "numbered":
                count = _positive_int(range_obj.get("numberOfOccurrences"), default=0)
                if count <= 0:
                    warnings.warn("invalid recurrence.range.numberOfOccurrences; COUNT omitted")
                else:
                    rule["COUNT"] = count
            elif range_type == "noend":
                pass
            elif range_type:
                warnings.warn(
                    f"unsupported recurrence range type '{range_type}'; RRULE end omitted"
                )

    event = Event()
    event.add("rrule", rule)
    rrule_value = event["RRULE"].to_ical().decode("utf-8")
    return rule, rrule_value, latest_until


def _monthday_rule(day_of_month: int) -> dict[str, Any]:
    """BYMONTHDAY for a Graph ``dayOfMonth``, keeping Outlook's short-month behaviour.

    Outlook books a series on the 29th, 30th or 31st on the *last* day of months that are
    shorter, whereas RFC 5545 simply skips a month with no such day. ``BYMONTHDAY=28..d``
    with ``BYSETPOS=-1`` picks the latest existing candidate in every month (or, under
    FREQ=YEARLY with BYMONTH, in that month), which reproduces Outlook exactly.
    """
    if day_of_month <= 28:
        return {"BYMONTHDAY": day_of_month}
    return {"BYMONTHDAY": list(range(28, day_of_month + 1)), "BYSETPOS": -1}


def _relative_byday(
    *,
    pattern: dict[str, Any],
    warnings: _WarningSink,
) -> dict[str, Any] | None:
    index_token = _INDEX_TOKEN.get(_lower_str(pattern.get("index")))
    if index_token is None:
        warnings.warn("invalid recurrence.pattern.index; RRULE omitted")
        return None
    days = _days_from_graph(pattern.get("daysOfWeek"))
    if not days:
        warnings.warn("invalid recurrence.pattern.daysOfWeek; RRULE omitted")
        return None
    if len(days) == 1:
        return {"BYDAY": f"{index_token}{days[0]}"}
    return {"BYDAY": days, "BYSETPOS": index_token}


def _days_from_graph(value: object) -> list[str] | None:
    if not isinstance(value, list) or not value:
        return None
    out: list[str] = []
    for item in value:
        token = _DAY_TOKEN.get(_lower_str(item))
        if token is None:
            return None
        out.append(token)
    return out


def _derive_exdates(
    *,
    span: _EventSpan,
    rrule_text: str | None,
    instances: list[dict[str, Any]],
    window: tuple[datetime, datetime],
    warnings: _WarningSink,
) -> list[datetime | date]:
    """Return EXDATE values for rule occurrences in ``window`` that Graph did not return.

    Timed series are expanded from the tz-aware local DTSTART, so dateutil keeps the local
    wall time across DST changes, and compared as UTC instants. All-day series are expanded
    from a naive midnight (their UNTIL is a DATE, which dateutil rejects next to an aware
    DTSTART) and compared as dates. If the rule expects occurrences in the window but Graph
    returned no instances at all, nothing is excluded: that looks like a failed or filtered
    fetch far more often than a series whose every occurrence was cancelled.
    """
    if rrule_text is None:
        return []
    expected = _expected_occurrences(span, rrule_text, window, warnings=warnings)
    if not expected:
        return []
    if not instances:
        warnings.warn("series master has empty instances list; EXDATE derivation skipped")
        return []

    present: set[datetime | date] = set()
    for item in instances:
        if not isinstance(item, dict):
            continue
        if _lower_str(item.get("type")) not in {"occurrence", "exception"}:
            continue
        anchor = _instance_anchor(item, span=span)
        if anchor is not None:
            present.add(anchor)

    present_instants = sorted(a for a in present if isinstance(a, datetime))
    missing: list[datetime | date] = []
    for occurrence in expected:
        if isinstance(occurrence, datetime):
            if not _has_instant_near(present_instants, occurrence.astimezone(UTC)):
                missing.append(occurrence)
        elif occurrence not in present:
            missing.append(occurrence)
    return missing


def _has_instant_near(sorted_instants: list[datetime], target: datetime) -> bool:
    """True if an instance starts within ``_MATCH_TOLERANCE`` of ``target``.

    Graph's finest recurrence is daily, so rule occurrences are at least 24h apart and a
    3h tolerance can never match the wrong one. The tolerance absorbs DST offsets when
    the series had to be expanded in UTC (unknown zone), which exact matching would turn
    into spurious EXDATEs that hide real occurrences.
    """
    index = bisect_left(sorted_instants, target - _MATCH_TOLERANCE)
    return index < len(sorted_instants) and sorted_instants[index] <= target + _MATCH_TOLERANCE


def _expected_occurrences(
    span: _EventSpan,
    rrule_text: str,
    window: tuple[datetime, datetime],
    *,
    warnings: _WarningSink,
) -> list[datetime | date]:
    """Expand the emitted RRULE over ``[window_start, window_end)``.

    Timed occurrences come back in DTSTART's own zone (TZID or UTC), all-day ones as dates,
    i.e. already in the form an EXDATE must take.
    """
    window_start = _as_utc(window[0])
    window_end = _as_utc(window[1])
    out: list[datetime | date] = []
    try:
        if isinstance(span.dtstart_value, datetime):
            local_start = span.dtstart_value
            rule = rrulestr(rrule_text, dtstart=local_start)
            for occurrence in rule.between(window_start, window_end, inc=True):
                if window_start <= occurrence.astimezone(UTC) < window_end:
                    out.append(occurrence.astimezone(local_start.tzinfo))
            return out
        naive_start = datetime.combine(span.dtstart_value, time.min)
        rule = rrulestr(rrule_text, dtstart=naive_start)
        naive_bounds = (window_start.replace(tzinfo=None), window_end.replace(tzinfo=None))
        for occurrence in rule.between(*naive_bounds, inc=True):
            if naive_bounds[0] <= occurrence < naive_bounds[1]:
                out.append(occurrence.date())
        return out
    except (ValueError, TypeError) as exc:
        warnings.warn(f"RRULE could not be expanded ({exc}); EXDATE derivation skipped")
        return []


def _instance_anchor(item: dict[str, Any], *, span: _EventSpan) -> datetime | date | None:
    """Identify which rule occurrence an instance stands for (originalStart, else start)."""
    raw = item.get("originalStart")
    parsed: datetime | None = None
    if isinstance(raw, str) and raw.strip():
        try:
            parsed = _parse_graph_datetime(raw)
        except ValueError:
            parsed = None
    if parsed is None:
        parsed = _parse_start_end_datetime(item.get("start"))
    if parsed is None:
        return None
    return _all_day_date(parsed, span.source_zone) if span.is_all_day else parsed


def _recurrence_id_from_original_start(
    value: str,
    *,
    span: _EventSpan,
    warnings: _WarningSink,
) -> datetime | date | None:
    """Express an exception's ``originalStart`` in the same form as the master DTSTART."""
    try:
        as_utc = _parse_graph_datetime(value)
    except ValueError:
        warnings.warn(f"exception instance has invalid originalStart '{value}'; skipped")
        return None
    if isinstance(span.dtstart_value, datetime):
        return as_utc.astimezone(span.dtstart_value.tzinfo)
    return _all_day_date(as_utc, span.source_zone)


def _collect_tzids(events: list[Event]) -> set[str]:
    tzids: set[str] = set()
    for event in events:
        for _name, value in event.property_items():
            params = getattr(value, "params", None)
            if params and "TZID" in params:
                tzids.add(str(params["TZID"]))
    return tzids


def _collect_emitted_datetimes(events: list[Event]) -> list[datetime]:
    out: list[datetime] = []
    for event in events:
        for _name, value in event.property_items():
            dt_value = getattr(value, "dt", None)
            if isinstance(dt_value, datetime):
                out.append(dt_value)
            dts = getattr(value, "dts", None)
            if dts is None:
                continue
            for ddd in dts:
                dt_item = getattr(ddd, "dt", None)
                if isinstance(dt_item, datetime):
                    out.append(dt_item)
    return out


def _parse_optional_datetime(
    value: object,
    *,
    field: str,
    warnings: _WarningSink,
) -> datetime | None:
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip():
        warnings.warn(f"invalid {field}; expected datetime string")
        return None
    try:
        return _parse_graph_datetime(value)
    except ValueError:
        warnings.warn(f"invalid {field}; expected datetime string")
        return None


def _parse_iso_date(value: object) -> date | None:
    if not isinstance(value, str):
        return None
    try:
        return date.fromisoformat(value)
    except ValueError:
        return None


def _as_utc(value: datetime) -> datetime:
    if value.tzinfo is None:
        return value.replace(tzinfo=UTC)
    return value.astimezone(UTC)


def _is_aware(value: datetime) -> bool:
    return value.tzinfo is not None and value.utcoffset() is not None


def _lower_str(value: object) -> str:
    return value.strip().lower() if isinstance(value, str) else ""


def _non_empty_str(value: object) -> str | None:
    if isinstance(value, str) and value.strip():
        return value.strip()
    return None


def _positive_int(value: object, *, default: int) -> int:
    if isinstance(value, int) and not isinstance(value, bool) and value > 0:
        return value
    return default


def _bounded_int(value: object, *, min_value: int, max_value: int) -> int | None:
    if isinstance(value, int) and not isinstance(value, bool) and min_value <= value <= max_value:
        return value
    return None
