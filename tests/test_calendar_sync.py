"""Calendar migrator with a fake Graph, fake DAV and a monkeypatched converter."""

from __future__ import annotations

import sqlite3
from datetime import UTC, datetime

import pytest

from fakes_o365 import MAPPING, FakeDav, FakeGraph, make_config
from o365_to_mailcow import calendar_sync
from o365_to_mailcow.calendar_sync import CalendarMigrator
from o365_to_mailcow.dav import DavError
from o365_to_mailcow.graph import GraphError
from o365_to_mailcow.state import STATUS_FAILED, State

U = "/users/alice@contoso.com"
NOW = datetime(2026, 9, 30, 12, 0, tzinfo=UTC)


def routes() -> dict:
    return {
        f"{U}/calendars": [
            {"id": "c-def", "name": "Calendar", "isDefaultCalendar": True,
             "owner": {"address": "alice@contoso.com"}},
            {"id": "c-team", "name": "Team Events", "isDefaultCalendar": False,
             "owner": {"address": "ALICE@contoso.com"}},
            {"id": "c-shared", "name": "Bob's calendar", "isDefaultCalendar": False,
             "owner": {"address": "bob@contoso.com"}},
            {"id": "c-p2", "name": "Personal", "owner": {"address": "alice@contoso.com"}},
        ],
        f"{U}/calendars/c-def/events": [
            {"id": "e1", "iCalUId": "UID1", "type": "singleInstance",
             "lastModifiedDateTime": "2024-01-01T00:00:00Z"},
            {"id": "e2", "iCalUId": "UID2", "type": "seriesMaster",
             "lastModifiedDateTime": "2024-01-02T00:00:00Z"},
            {"id": "e3", "iCalUId": "UID2", "type": "occurrence"},
        ],
        f"{U}/events/e2/instances": [{"id": "e2-i1", "type": "exception"}],
        f"{U}/calendars/c-team/events": [
            {"id": "e4", "type": "singleInstance", "lastModifiedDateTime": "2024-03-01T00:00:00Z"},
        ],
        f"{U}/calendars/c-p2/events": [],
    }


class Converter:
    def __init__(self, warnings=("unknown time zone 'Mars Standard Time' mapped to UTC",)):
        self.calls: list[tuple[dict, list, dict]] = []
        self.warnings = tuple(warnings)
        self.fail_for: set[str] = set()

    def __call__(self, master, instances, *, window, attendees="keep"):
        self.calls.append((master, instances, {"window": window, "attendees": attendees}))
        if master["id"] in self.fail_for:
            raise ValueError("bad recurrence")
        mod = calendar_sync.calendar_conv
        uid = master.get("iCalUId") or f"uuid5-{master['id']}"
        return mod.ConvertedEvent(uid, f"ICS {uid}".encode(),
                                  master.get("lastModifiedDateTime"), self.warnings)


@pytest.fixture
def env(tmp_path, monkeypatch):
    conv = Converter()
    monkeypatch.setattr(calendar_sync.calendar_conv, "convert_event", conv)
    cfg = make_config(tmp_path, calendar_attendees="strip")
    state = State(tmp_path / "state.db")
    yield cfg, state, FakeGraph(routes()), FakeDav(), conv
    state.close()


def mig(cfg, state, graph, dav, dry_run=False):
    return CalendarMigrator(cfg, graph, state, dav, MAPPING, dry_run, now=NOW)


def test_plan_owned_only_with_slugs_isc_56(env):
    cfg, state, graph, dav, _ = env
    plan = mig(cfg, state, graph, None, dry_run=True).plan()
    assert [(c.name, c.slug, c.count) for c in plan.collections] == [
        ("Calendar", "personal", 2), ("Team Events", "team-events", 1),
        ("Personal", "personal-2", 0)]
    assert plan.skipped == ["Bob's calendar (shared by bob@contoso.com)"]


def test_migrate_puts_events_into_right_collections_isc_57_105(env):
    cfg, state, graph, dav, conv = env
    res = mig(cfg, state, graph, dav).migrate()
    assert ("PUT", "Calendar/personal/UID1.ics") in dav.writes
    assert ("PUT", "Calendar/personal/UID2.ics") in dav.writes
    assert ("PUT", "Calendar/team-events/uuid5-e4.ics") in dav.writes
    assert ("MKCALENDAR", "team-events") in dav.writes
    assert ("MKCALENDAR", "personal") not in dav.writes  # default is never created
    assert len([w for w in dav.writes if w[0] == "PUT"]) == 3  # occurrence ignored
    assert sum(c.put for c in res.collections) == 3 and res.failed == 0
    # ISC-83: nothing of the shared calendar was even listed
    assert not [p for p in graph.paths() if "c-shared" in p]


def test_converter_gets_instances_window_and_attendee_mode(env):
    cfg, state, graph, dav, conv = env
    mig(cfg, state, graph, dav).migrate()
    by_id = {m["id"]: (inst, kw) for m, inst, kw in conv.calls}
    assert by_id["e1"][0] == []
    assert by_id["e2"][0] == [{"id": "e2-i1", "type": "exception"}]
    start, end = by_id["e2"][1]["window"]
    assert (NOW - start).days == 730 and (end - NOW).days == 1095
    assert by_id["e2"][1]["attendees"] == "strip"
    inst_call = next(c for c in graph.calls if c[1] == f"{U}/events/e2/instances")
    assert inst_call[2]["startDateTime"] == "2024-09-30T12:00:00Z"
    assert inst_call[2]["endDateTime"] == "2029-09-29T12:00:00Z"
    assert "originalStart" in inst_call[2]["$select"]
    ev_call = next(c for c in graph.calls if c[1] == f"{U}/calendars/c-def/events")
    assert ev_call[3] == {"Prefer": 'outlook.timezone="UTC"'}
    assert "iCalUId" in ev_call[2]["$select"] and ev_call[2]["$top"] >= 100


def test_second_run_puts_only_changed_events_isc_81(env):
    cfg, state, graph, dav, _ = env
    mig(cfg, state, graph, dav).migrate()
    dav.writes.clear()
    res = mig(cfg, state, graph, dav).migrate()
    assert not [w for w in dav.writes if w[0] == "PUT"]
    assert sum(c.unchanged for c in res.collections) == 3
    graph.routes[f"{U}/calendars/c-def/events"][0]["lastModifiedDateTime"] = "2025-01-01T00:00:00Z"
    mig(cfg, state, graph, dav).migrate()
    assert [w for w in dav.writes if w[0] == "PUT"] == [("PUT", "Calendar/personal/UID1.ics")]


def test_warnings_reported_once_isc_62(env):
    cfg, state, graph, dav, _ = env
    res = mig(cfg, state, graph, dav).migrate()
    assert res.warnings == ["unknown time zone 'Mars Standard Time' mapped to UTC"]


def test_put_failure_recorded_and_continues_isc_108(env):
    cfg, state, graph, dav, _ = env
    dav.put_errors["UID1"] = DavError(403, "Forbidden: body")
    res = mig(cfg, state, graph, dav).migrate()
    assert res.failed == 1
    assert state.event_counts(MAPPING.source)["personal"][STATUS_FAILED] == 1
    assert ("PUT", "Calendar/personal/UID2.ics") in dav.writes
    # a failed event is retried on the next run
    dav.put_errors.clear()
    dav.writes.clear()
    mig(cfg, state, graph, dav).migrate()
    assert dav.writes == [("PUT", "Calendar/personal/UID1.ics")]


def test_converter_exception_is_item_failure(env):
    cfg, state, graph, dav, conv = env
    conv.fail_for.add("e2")
    res = mig(cfg, state, graph, dav).migrate()
    assert res.failed == 1 and sum(c.put for c in res.collections) == 2


def test_instances_graph_failure_is_item_failure(env):
    cfg, state, graph, dav, _ = env
    graph.routes[f"{U}/events/e2/instances"] = GraphError(500, "boom", "x")
    res = mig(cfg, state, graph, dav).migrate()
    assert res.failed == 1


def test_missing_calendar_home_skips_with_clear_error_isc_107(env):
    cfg, state, graph, _, _ = env
    dav = FakeDav(calendar_home=False)
    res = mig(cfg, state, graph, dav).migrate()
    assert dav.writes == []
    assert "alice@contoso.com" in res.errors[0] and "404" in res.errors[0]


def test_mkcalendar_refused_marks_calendar_error(env):
    cfg, state, graph, dav, _ = env

    def refuse(name, slug=None):
        raise DavError(405, "Method Not Allowed")

    dav.ensure_calendar = refuse
    res = mig(cfg, state, graph, dav).migrate()
    team = next(c for c in res.collections if c.slug == "team-events")
    assert team.error and "405" in team.error
    assert res.failed >= 1
    assert ("PUT", "Calendar/personal/UID1.ics") in dav.writes


def test_dry_run_needs_no_dav_and_counts_isc_96(env):
    cfg, state, graph, _, conv = env
    res = mig(cfg, state, graph, None, dry_run=True).migrate()
    assert sum(c.would_put for c in res.collections) == 3
    assert conv.calls == []
    assert state.event_counts(MAPPING.source) == {}


def test_verify_counts_isc_82(env):
    cfg, state, graph, dav, _ = env
    mig(cfg, state, graph, dav).migrate()
    v = mig(cfg, state, graph, dav).verify()
    rows = {c.slug: c for c in v.collections}
    assert (rows["personal"].graph_count, rows["personal"].dav_count) == (2, 2)
    assert not rows["personal"].mismatch
    assert rows["personal-2"].dav_count == 0 and not rows["personal-2"].mismatch
    assert v.skipped == ["Bob's calendar (shared by bob@contoso.com)"]
    dav.calendars["personal"].pop("UID1")
    v = mig(cfg, state, graph, dav).verify()
    assert {c.slug: c for c in v.collections}["personal"].mismatch


# -- labels of failed events --------------------------------------------------------------

def test_put_failure_is_listed_with_title_start_and_calendar(env):
    cfg, state, graph, dav, _ = env
    graph.routes[f"{U}/calendars/c-def/events"][0].update(
        subject="Board meeting", start={"dateTime": "2024-03-05T09:00:00.0000000",
                                        "timeZone": "UTC"})
    dav.put_errors["UID1"] = DavError(500, "")
    res = mig(cfg, state, graph, dav).migrate()
    expected = [{"place": "Calendar", "title": "Board meeting",
                 "hint": "starts 2024-03-05 09:00 UTC", "status": "failed",
                 "error": "DAV HTTP 500: "}]
    assert (res.failed_items, res.failed_items_total) == (expected, 1)
    assert res.failed == 1  # the listed items are not counted a second time
    v = mig(cfg, state, graph, dav).verify()
    assert (v.failed_items, v.failed_items_total) == (expected, 1)


def test_event_without_subject_or_start_is_still_listed(env):
    cfg, state, graph, dav, _ = env
    dav.put_errors["UID1"] = DavError(500, "")
    res = mig(cfg, state, graph, dav).migrate()
    assert [(i["title"], i["hint"]) for i in res.failed_items] == [("(no title)", "")]


def test_all_day_and_recurring_hints():
    from o365_to_mailcow.calendar_sync import event_label

    start = {"dateTime": "2024-03-05T00:00:00.0000000", "timeZone": "UTC"}
    assert event_label({"subject": "Holiday", "isAllDay": True, "start": start}) == (
        "Holiday", "all day 2024-03-05")
    assert event_label({"type": "seriesMaster", "start": start}) == (
        "(no title)", "starts 2024-03-05 00:00 UTC · recurring")
    assert event_label({"start": "garbage", "subject": 5}) == ("(no title)", "")


def test_verify_lists_stored_failures_when_listing_fails(env):
    cfg, state, graph, dav, _ = env
    dav.put_errors["UID1"] = DavError(500, "")
    mig(cfg, state, graph, dav).migrate()
    graph.routes[f"{U}/calendars"] = GraphError(503, "down", "x")
    v = mig(cfg, state, graph, dav).verify()
    assert v.errors and v.failed_items_total == 1


def test_dry_run_lists_no_failed_events(env):
    cfg, state, graph, _, _ = env
    state.mark_event(MAPPING.source, "personal", "UID1", None, STATUS_FAILED, "x")
    res = mig(cfg, state, graph, None, dry_run=True).migrate()
    assert (res.failed_items, res.failed_items_total) == ([], 0)


@pytest.mark.parametrize(("method", "error"), [
    ("set_label", sqlite3.OperationalError("database is locked")),
    ("failed_items", sqlite3.OperationalError("database is locked")),
    ("failed_items", MemoryError()),
])
def test_an_error_in_the_label_code_changes_no_outcome(env, monkeypatch, method, error):
    cfg, state, graph, dav, _ = env
    dav.put_errors["UID1"] = DavError(403, "Forbidden: body")

    def boom(*args, **kwargs):
        raise error

    monkeypatch.setattr(state, method, boom)
    res = mig(cfg, state, graph, dav).migrate()
    assert res.failed == 1 and res.errors == [] and res.duration_s >= 0
    assert ("PUT", "Calendar/personal/UID2.ics") in dav.writes
    assert state.event_counts(MAPPING.source)["personal"][STATUS_FAILED] == 1
    v = mig(cfg, state, graph, dav).verify()
    assert v.errors == [] and [c.failed for c in v.collections][0] == 1
    if method == "failed_items":
        assert res.failed_items == [] and v.failed_items == []
