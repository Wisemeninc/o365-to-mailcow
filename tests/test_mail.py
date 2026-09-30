"""Mail path: folder mapping, flags, dates, idempotency, dedupe, failures, delta, verify."""

from __future__ import annotations

import threading
from datetime import UTC, datetime

import pytest

from fakes_o365 import MAPPING, FakeGraph, FakeImap, ImapWorld, make_config
from o365_to_mailcow.graph import GraphError
from o365_to_mailcow.imap_dest import ImapError
from o365_to_mailcow.mail import (
    MESSAGE_SELECT,
    MailMigrator,
    category_keyword,
    imap_flags,
    internal_date,
    sanitize_folder_name,
)
from o365_to_mailcow.state import STATUS_DONE, STATUS_FAILED, STATUS_SKIPPED, State

U = "/users/alice@contoso.com"
DELTA_INBOX = "https://graph.microsoft.com/v1.0/delta?token=inbox-1"


def mime(mid: str | None, body: str = "hello") -> bytes:
    head = f"Message-ID: {mid}\r\n" if mid else ""
    return f"{head}Subject: {body}\r\n\r\n{body}\r\n".encode()


def folder(fid: str, name: str, total: int = 0, children: int = 0, size: int | None = None):
    f = {"id": fid, "displayName": name, "totalItemCount": total, "childFolderCount": children}
    if size is not None:
        f["singleValueExtendedProperties"] = [{"id": "Long 0x0e08", "value": str(size)}]
    return f


M1 = {
    "id": "m1", "internetMessageId": "<m1@x>", "isRead": True, "isDraft": False,
    "flag": {"flagStatus": "flagged"}, "categories": ["Red category", "Büro"],
    "receivedDateTime": "2024-01-02T03:04:05Z", "lastModifiedDateTime": "2024-01-03T00:00:00Z",
    "singleValueExtendedProperties": [{"id": "Integer 0xe08", "value": "120"}],
}
M2 = {
    "id": "m2", "isRead": False, "isDraft": True, "flag": {"flagStatus": "notFlagged"},
    "categories": [], "receivedDateTime": None, "lastModifiedDateTime": "2023-05-01T10:00:00Z",
}


def base_routes() -> dict:
    return {
        f"{U}/mailFolders/inbox": {"id": "f-inbox"},
        f"{U}/mailFolders/sentitems": {"id": "f-sent"},
        f"{U}/mailFolders/conversationhistory": {"id": "f-conv"},
        f"{U}/mailFolders": [
            folder("f-inbox", "Inbox", 2, children=1, size=5000),
            folder("f-sent", "Sent Items", 0, size=0),
            folder("f-conv", "Conversation History", 3, children=1, size=10),
            folder("f-proj", "Projects/2024", 0, size=0),
            folder("f-mysent", "  sent\x07 ", 0, size=0),
            folder("f-uni", "Bestätigungen", 0, size=0),
        ],
        f"{U}/mailFolders/f-inbox/childFolders": [folder("f-sub", "Sub", 0, size=0)],
        f"{U}/mailFolders/f-conv/childFolders": [folder("f-conv-child", "Team Chat", 1)],
        f"{U}/mailFolders/f-inbox/messages": [dict(M1), dict(M2)],
        f"{U}/mailFolders/f-inbox/messages/delta": ([{"id": "m1"}, {"id": "m2"}], DELTA_INBOX),
        f"{U}/messages/m1/$value": mime("<m1@x>", "one"),
        f"{U}/messages/m2/$value": mime(None, "two"),
        **{f"{U}/mailFolders/{fid}/messages": [] for fid in
           ("f-sent", "f-proj", "f-mysent", "f-uni", "f-sub")},
        **{f"{U}/mailFolders/{fid}/messages/delta": ([], f"https://graph.microsoft.com/d/{fid}")
           for fid in ("f-sent", "f-proj", "f-mysent", "f-uni", "f-sub")},
    }


@pytest.fixture
def env(tmp_path):
    cfg = make_config(tmp_path)
    state = State(tmp_path / "state.db")
    world = ImapWorld()
    graph = FakeGraph(base_routes())
    yield cfg, state, world, graph
    state.close()


def migrator(cfg, state, world, graph, dry_run=False, imap_cls=FakeImap):
    return MailMigrator(cfg, graph, state, lambda: imap_cls(world), MAPPING, dry_run)


# -- helpers ---------------------------------------------------------------------------

def test_flags_isc_42_to_45():
    assert imap_flags(M1) == ["\\Seen", "\\Flagged", "Red_category", "B_ro"]
    assert imap_flags(M2) == ["\\Draft"]
    assert "\\Seen" not in imap_flags({"isRead": False})


def test_category_keyword_replaces_disallowed_chars():
    assert category_keyword("Needs (review) *now*") == "Needs__review___now_"
    assert category_keyword('a"b\\c]d') == "a_b_c_d"


def test_internal_date_fallbacks_isc_46_103():
    assert internal_date(M1) == datetime(2024, 1, 2, 3, 4, 5, tzinfo=UTC)
    assert internal_date(M2) == datetime(2023, 5, 1, 10, 0, tzinfo=UTC)
    fixed = datetime(2020, 1, 1, tzinfo=UTC)
    assert internal_date({}, now=lambda: fixed) == fixed


def test_sanitize_folder_name_isc_104():
    assert sanitize_folder_name("  a\x00b\x1fc \t", "/") == "abc"
    assert sanitize_folder_name("a/b", "/") == "a_b"
    assert sanitize_folder_name("a.b", ".") == "a_b"
    assert sanitize_folder_name("\x01 ", "/") == "Unnamed"


# -- plan ------------------------------------------------------------------------------

def test_plan_maps_well_known_skips_and_hierarchy(env):
    cfg, state, world, graph = env
    plan = migrator(cfg, state, world, graph).plan()
    by_id = {f.folder_id: f for f in plan.folders}
    assert by_id["f-inbox"].dest_name == "INBOX"
    assert by_id["f-sent"].dest_name == "Sent"
    assert by_id["f-sub"].dest_name == "INBOX/Sub"  # child keeps hierarchy (ISC-39)
    assert by_id["f-proj"].dest_name == "Projects_2024"  # delimiter inside a name
    assert by_id["f-mysent"].dest_name == "sent (2)"  # collides with Sent (ISC-104)
    assert by_id["f-uni"].dest_name == "Bestätigungen"
    assert by_id["f-conv"].skip and by_id["f-conv"].skip_reason == "well-known folder " \
        "conversationhistory"
    assert by_id["f-conv-child"].skip  # children of skipped folders (ISC-38)
    assert not by_id["f-inbox"].skip
    # ISC-36/117: totals and bytes of migrated folders only
    assert plan.total_messages == 2
    assert plan.total_bytes == 5000 and plan.bytes_known is True
    assert plan.skipped_folders == 2
    # parents are listed before children
    ids = [f.folder_id for f in plan.folders]
    assert ids.index("f-inbox") < ids.index("f-sub")


def test_plan_without_folder_size_reports_bytes_unknown(env):
    cfg, state, world, graph = env
    graph.routes[f"{U}/mailFolders/f-inbox/childFolders"] = [folder("f-sub", "Sub", 1)]
    plan = migrator(cfg, state, world, graph).plan()
    assert plan.bytes_known is False and plan.total_messages == 3


def test_plan_uses_immutable_ids_and_makes_no_writes(env):
    cfg, state, world, graph = env
    migrator(cfg, state, world, graph).plan()
    assert graph.calls and all(h == {"Prefer": 'IdType="ImmutableId"'} for *_, h in graph.calls)
    assert world.connections == 0


def test_source_folder_skip_config(env, tmp_path):
    cfg, state, world, graph = env
    cfg = make_config(tmp_path, source_folder_skip=("Inbox/Sub",))
    plan = migrator(cfg, state, world, graph).plan()
    assert {f.folder_id: f.skip for f in plan.folders}["f-sub"] is True


# -- migrate ---------------------------------------------------------------------------

def test_migrate_appends_with_flags_dates_and_state(env):
    cfg, state, world, graph = env
    res = migrator(cfg, state, world, graph).migrate()
    assert not res.errors and not res.stopped
    inbox = world.folders["INBOX"]
    assert [s.message_id for s in inbox] == [None, "<m1@x>"]  # oldest first
    by_mid = {s.message_id: s for s in inbox}
    assert by_mid["<m1@x>"].flags == ["\\Seen", "\\Flagged", "Red_category", "B_ro"]
    assert by_mid["<m1@x>"].date == datetime(2024, 1, 2, 3, 4, 5, tzinfo=UTC)
    assert by_mid[None].flags == ["\\Draft"]
    assert by_mid["<m1@x>"].mime == mime("<m1@x>", "one")  # byte-for-byte (ISC-41)
    # every non-skipped folder exists, skipped ones do not
    assert {"INBOX", "INBOX/Sub", "Sent", "Projects_2024", "sent (2)", "Bestätigungen"} \
        <= set(world.folders)
    assert "Conversation History" not in world.folders
    # ISC-101/121: state has dest uid and uidvalidity
    assert state.done_message_ids(MAPPING.source, "f-inbox") == {"m1", "m2"}
    assert state.folder_uidvalidity(MAPPING.source, "f-inbox") == world.uidvalidity["INBOX"]
    # ISC-51: delta link stored after the full pass
    assert state.get_delta(MAPPING.source, "f-inbox") == DELTA_INBOX
    assert res.total("appended") == 2


def test_every_mail_request_has_prefer_and_select(env):
    cfg, state, world, graph = env
    migrator(cfg, state, world, graph).migrate()
    for _, path, params, headers in graph.calls:
        assert 'IdType="ImmutableId"' in headers["Prefer"], path  # ISC-120
        if path.endswith("/messages"):
            assert params["$select"] == MESSAGE_SELECT  # ISC-97
            assert params["$top"] >= 100


def test_second_run_appends_zero_isc_50(env):
    cfg, state, world, graph = env
    migrator(cfg, state, world, graph).migrate()
    graph.routes[DELTA_INBOX] = ([{"id": "m1"}], "https://graph.microsoft.com/d/inbox-2")
    for fid in ("f-sent", "f-proj", "f-mysent", "f-uni", "f-sub"):
        link = f"https://graph.microsoft.com/d/{fid}"
        graph.routes[link] = ([], link)
    graph.calls.clear()
    appends_before = world.appends
    res = migrator(cfg, state, world, graph).migrate()
    assert world.appends == appends_before
    assert res.total("appended") == 0
    assert not [p for p in graph.paths("get_bytes")]  # nothing re-downloaded (ISC-47)
    assert DELTA_INBOX in graph.paths("get_delta")  # used the stored link (ISC-51)
    assert state.get_delta(MAPPING.source, "f-inbox") == "https://graph.microsoft.com/d/inbox-2"


def test_message_id_hit_marks_done_without_append_isc_48(env):
    cfg, state, world, graph = env
    world.folders["INBOX"] = []
    world.uidvalidity["INBOX"] = 7
    # a true earlier copy of m1 (same MIME) already sits in the destination
    FakeImap(world).append("INBOX", mime("<m1@x>", "one"), [], datetime.now(UTC))
    searched: list[str] = []

    class Recording(FakeImap):
        def search_message_id(self, folder, message_id):
            searched.append(message_id)
            return super().search_message_id(folder, message_id)

    res = migrator(cfg, state, world, graph, imap_cls=Recording).migrate()
    inbox_res = next(f for f in res.folders if f.dest_name == "INBOX")
    assert inbox_res.dedup_hits == 1 and inbox_res.appended == 1
    assert len(world.folders["INBOX"]) == 2  # pre-existing + m2 only
    assert state.message_status(MAPPING.source, "f-inbox", "m1") == STATUS_DONE
    assert searched == ["<m1@x>"]  # m2 has no Message-ID: never searched (ISC-98)


def test_message_id_hit_with_different_content_is_appended_not_skipped(env):
    """A stranger can send mail carrying a known Message-ID (Silas M4): only a copy whose
    content matches counts as already migrated."""
    cfg, state, world, graph = env
    world.folders["INBOX"] = []
    world.uidvalidity["INBOX"] = 7
    FakeImap(world).append("INBOX", mime("<m1@x>", "forged"), [], datetime.now(UTC))
    res = migrator(cfg, state, world, graph).migrate()
    inbox_res = next(f for f in res.folders if f.dest_name == "INBOX")
    assert inbox_res.dedup_hits == 0 and inbox_res.appended == 2
    assert len(world.folders["INBOX"]) == 3  # forged + m1 + m2


def test_message_id_hit_only_dedupes_more_copies_than_recorded(env):
    """Two source items sharing a Message-ID (Cato F3): the second is not swallowed by
    the first one's destination copy."""
    cfg, state, world, graph = env
    world.folders["INBOX"] = []
    world.uidvalidity["INBOX"] = 7
    FakeImap(world).append("INBOX", mime("<m1@x>", "one"), [], datetime.now(UTC))
    # state already says one source item with this Message-ID is done here
    state.mark_message(MAPPING.source, "f-inbox", "m0", "INBOX", "<m1@x>", STATUS_DONE,
                       dest_uid=1, uidvalidity=7)
    res = migrator(cfg, state, world, graph).migrate()
    inbox_res = next(f for f in res.folders if f.dest_name == "INBOX")
    assert inbox_res.dedup_hits == 0 and inbox_res.appended == 2


def test_graph_download_failure_recorded_and_run_continues_isc_49(env):
    cfg, state, world, graph = env
    graph.routes[f"{U}/messages/m2/$value"] = GraphError(503, "retries exhausted", "x")
    res = migrator(cfg, state, world, graph).migrate()
    assert not res.stopped
    assert state.message_status(MAPPING.source, "f-inbox", "m2") == STATUS_FAILED
    assert state.message_status(MAPPING.source, "f-inbox", "m1") == STATUS_DONE
    assert res.total("failed") == 1
    # a folder with failures does not store its delta link, so a re-run retries (full list)
    assert state.get_delta(MAPPING.source, "f-inbox") is None


def test_quota_failure_stops_mailbox_with_one_error_isc_99(env):
    cfg, state, world, graph = env
    world.quota_after = 1
    res = migrator(cfg, state, world, graph).migrate()
    assert res.stopped and len(res.errors) == 1 and "OVERQUOTA" in res.errors[0]
    assert len(world.folders["INBOX"]) == 1
    assert state.message_status(MAPPING.source, "f-inbox", "m1") == STATUS_FAILED
    # folders after INBOX were not processed at all
    assert [f.dest_name for f in res.folders] == ["INBOX"]


def test_oversized_message_skipped_isc_52(env, tmp_path):
    cfg, state, world, graph = env
    cfg = make_config(tmp_path, max_message_bytes=100)  # M1 reports 120 bytes
    res = migrator(cfg, state, world, graph).migrate()
    assert state.message_status(MAPPING.source, "f-inbox", "m1") == STATUS_SKIPPED
    assert f"{U}/messages/m1/$value" not in graph.paths("get_bytes")
    assert res.total("skipped_too_large") == 1


def test_interrupted_run_resumes_without_reappending_isc_101(env):
    cfg, state, world, graph = env

    class Killed(Exception):
        pass

    class Dies(FakeImap):
        def append(self, *a, **kw):
            if self.world.appends >= 1:
                raise Killed()
            return super().append(*a, **kw)

    with pytest.raises(Killed):
        migrator(cfg, state, world, graph, imap_cls=Dies).migrate()
    assert len(world.folders["INBOX"]) == 1
    assert state.done_message_ids(MAPPING.source, "f-inbox") == {"m2"}
    migrator(cfg, state, world, graph).migrate()
    assert len(world.folders["INBOX"]) == 2  # only the missing one was appended
    assert state.done_message_ids(MAPPING.source, "f-inbox") == {"m1", "m2"}


def test_single_imap_connection_downloads_on_threads_isc_102(env):
    cfg, state, world, graph = env
    threads: set[str] = set()
    for gid in ("m1", "m2"):
        payload = graph.routes[f"{U}/messages/{gid}/$value"]

        def route(_params, payload=payload):
            threads.add(threading.current_thread().name)
            return payload

        graph.routes[f"{U}/messages/{gid}/$value"] = route
    migrator(cfg, state, world, graph).migrate()
    assert world.connections == 1
    assert world.append_threads == {threading.current_thread().name}
    assert threads and all(t.startswith("dl") for t in threads)


def test_delta_removed_counted_not_applied_isc_124(env):
    cfg, state, world, graph = env
    migrator(cfg, state, world, graph).migrate()
    before = [s.uid for s in world.folders["INBOX"]]
    graph.routes[DELTA_INBOX] = (
        [{"id": "m1", "@removed": {"reason": "deleted"}}, {"id": "m3"}],
        "https://graph.microsoft.com/d/inbox-2")
    graph.routes[f"{U}/messages/m3"] = {**M1, "id": "m3", "internetMessageId": "<m3@x>"}
    graph.routes[f"{U}/messages/m3/$value"] = mime("<m3@x>", "three")
    for fid in ("f-sent", "f-proj", "f-mysent", "f-uni", "f-sub"):
        graph.routes[f"https://graph.microsoft.com/d/{fid}"] = ([], None)
    res = migrator(cfg, state, world, graph).migrate()
    inbox = next(f for f in res.folders if f.dest_name == "INBOX")
    assert inbox.removed_in_source == 1 and inbox.appended == 1
    assert [s.uid for s in world.folders["INBOX"]][:2] == before  # nothing removed


def test_uidvalidity_change_forces_message_id_check_isc_121(env):
    cfg, state, world, graph = env
    migrator(cfg, state, world, graph).migrate()
    world.folders["INBOX"] = []  # folder recreated on the server
    world.uidvalidity["INBOX"] = 99999
    res = migrator(cfg, state, world, graph).migrate()
    inbox = next(f for f in res.folders if f.dest_name == "INBOX")
    assert inbox.uidvalidity_changed
    assert f"{U}/mailFolders/f-inbox/messages" in graph.paths("iter_pages")  # full listing
    # m1 (has Message-ID, missing) re-appended; m2 cannot be checked and stays done
    assert [s.message_id for s in world.folders["INBOX"]] == ["<m1@x>"]


def test_dry_run_makes_zero_writes_isc_96(env):
    cfg, state, world, graph = env

    def no_imap():
        raise AssertionError("dry run must not connect to IMAP")

    res = MailMigrator(cfg, graph, state, no_imap, MAPPING, True).migrate()
    assert res.dry_run and res.total("would_append") == 2
    assert world.connections == 0 and world.appends == 0
    assert state.message_counts(MAPPING.source) == {}
    assert state.get_delta(MAPPING.source, "f-inbox") is None
    assert not graph.paths("get_bytes")


# -- verify ----------------------------------------------------------------------------

def test_verify_counts_and_sample_isc_53_119_129(env):
    cfg, state, world, graph = env
    graph.routes[f"{U}/messages/m1/$value"] = mime("<m1@x>", "one").replace(b"\r\n", b"\n")
    migrator(cfg, state, world, graph).migrate()
    graph.routes[f"{U}/messages/m1"] = {"internetMessageId": "<m1@x>"}
    graph.routes[f"{U}/messages/m2"] = {"internetMessageId": None}
    v = migrator(cfg, state, world, graph).verify(sample=5)
    inbox = next(f for f in v.folders if f.dest_name == "INBOX")
    assert (inbox.graph_total, inbox.done, inbox.imap_count, inbox.expected) == (2, 2, 2, 2)
    assert not inbox.mismatch
    assert v.sample_checked == 1 and v.sample_unverifiable == 1 and not v.sample_mismatches
    assert {"path": "Conversation History", "reason": "well-known folder conversationhistory",
            "total": 3} in v.skipped_folders


def test_verify_flags_mismatch_and_bad_sample(env):
    cfg, state, world, graph = env
    migrator(cfg, state, world, graph).migrate()
    stored = next(s for s in world.folders["INBOX"] if s.message_id == "<m1@x>")
    stored.mime = b"Message-ID: <m1@x>\r\n\r\ntampered\r\n"
    world.folders["INBOX"].append(stored)  # extra copy -> count mismatch
    graph.routes[f"{U}/messages/m1"] = {"internetMessageId": "<m1@x>"}
    graph.routes[f"{U}/messages/m2"] = {}
    v = migrator(cfg, state, world, graph).verify(sample=2)
    inbox = next(f for f in v.folders if f.dest_name == "INBOX")
    assert inbox.mismatch and inbox.imap_count == 3
    assert v.sample_mismatches == ["INBOX: <m1@x>"]


def test_verify_expected_subtracts_failed_and_skipped(env):
    cfg, state, world, graph = env
    graph.routes[f"{U}/messages/m2/$value"] = GraphError(500, "boom", "x")
    migrator(cfg, state, world, graph).migrate()
    v = migrator(cfg, state, world, graph).verify()
    inbox = next(f for f in v.folders if f.dest_name == "INBOX")
    assert (inbox.graph_total, inbox.failed, inbox.expected, inbox.imap_count) == (2, 1, 1, 1)
    assert not inbox.mismatch


def test_imap_error_on_folder_status_is_folder_error_not_crash(env):
    cfg, state, world, graph = env

    class NoStatus(FakeImap):
        def folder_status(self, name):
            if name == "Projects_2024":
                raise ImapError("NO")
            return super().folder_status(name)

    res = migrator(cfg, state, world, graph, imap_cls=NoStatus).migrate()
    assert not res.stopped
    assert any("Projects/2024" in e for e in res.errors)
    assert len(world.folders["INBOX"]) == 2
