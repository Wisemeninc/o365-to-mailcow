"""State ledger: idempotency keys, persistence, file mode (ISC-101, 120, 121, 125, 130)."""

from __future__ import annotations

import sqlite3
import stat
import threading

import pytest

from o365_to_mailcow.state import STATUS_DONE, STATUS_FAILED, STATUS_SKIPPED, State


def test_db_file_mode_0600_isc_130(tmp_path):
    path = tmp_path / "sub" / "state.db"
    State(path).close()
    assert stat.S_IMODE(path.stat().st_mode) == 0o600


def test_messages_keyed_by_mailbox_folder_id_isc_120(tmp_path):
    s = State(tmp_path / "s.db")
    s.mark_message("a@x", "f1", "g1", "INBOX", "<m@x>", STATUS_DONE, dest_uid=5, uidvalidity=9)
    s.mark_message("a@x", "f2", "g1", "Sent", None, STATUS_FAILED, error="HTTP 500")
    s.mark_message("b@x", "f1", "g1", "INBOX", None, STATUS_SKIPPED)
    assert s.message_status("a@x", "f1", "g1") == STATUS_DONE
    assert s.message_status("a@x", "f2", "g1") == STATUS_FAILED
    assert s.message_status("a@x", "f3", "g1") is None
    assert s.done_message_ids("a@x", "f1") == {"g1"}
    assert s.done_message_ids("a@x", "f2") == set()
    assert s.folder_uidvalidity("a@x", "f1") == 9  # ISC-121
    assert s.folder_uidvalidity("a@x", "f2") is None
    assert s.message_counts("a@x") == {"INBOX": {"done": 1}, "Sent": {"failed": 1}}
    s.close()


def test_every_write_is_committed_isc_101(tmp_path):
    s = State(tmp_path / "s.db")
    s.mark_message("a@x", "f1", "g1", "INBOX", None, STATUS_DONE)
    # a second connection sees it without close/commit (autocommit per statement)
    other = State(tmp_path / "s.db")
    assert other.done_message_ids("a@x", "f1") == {"g1"}
    other.close()
    s.close()


def test_overwrite_status(tmp_path):
    s = State(tmp_path / "s.db")
    s.mark_message("a@x", "f1", "g1", "INBOX", None, STATUS_FAILED, error="x" * 1000)
    s.mark_message("a@x", "f1", "g1", "INBOX", None, STATUS_DONE)
    assert s.message_status("a@x", "f1", "g1") == STATUS_DONE
    s.close()


def test_delta_links(tmp_path):
    s = State(tmp_path / "s.db")
    assert s.get_delta("a@x", "f1") is None
    s.set_delta("a@x", "f1", "https://graph.microsoft.com/d1")
    s.set_delta("a@x", "f1", "https://graph.microsoft.com/d2")
    assert s.get_delta("a@x", "f1") == "https://graph.microsoft.com/d2"
    s.close()


def test_app_passwords_isc_125(tmp_path):
    s = State(tmp_path / "s.db")
    s.record_app_password("a@x", "12")
    s.record_app_password("b@x", "13")
    assert sorted(s.app_passwords()) == [("a@x", "12"), ("b@x", "13")]
    s.forget_app_password("a@x", "12")
    assert s.app_passwords() == [("b@x", "13")]
    s.close()


def test_events_and_contacts(tmp_path):
    s = State(tmp_path / "s.db")
    s.mark_event("a@x", "personal", "UID1", "2024-01-01T00:00:00Z", STATUS_DONE)
    s.mark_event("a@x", "personal", "UID2", "2024-01-01T00:00:00Z", STATUS_FAILED, "HTTP 403")
    assert s.event_last_modified("a@x", "personal", "UID1") == "2024-01-01T00:00:00Z"
    assert s.event_last_modified("a@x", "personal", "UID2") is None  # failed -> retried
    assert s.event_counts("a@x") == {"personal": {"done": 1, "failed": 1}}
    s.mark_contact("a@x", "c1", "personal", "lm1", STATUS_DONE)
    s.mark_contact("a@x", "c2", "suppliers", "lm2", STATUS_FAILED, "bad")
    assert s.contact_last_modified("a@x", "c1") == "lm1"
    assert s.contact_last_modified("a@x", "c2") is None
    assert s.contact_counts("a@x") == {"personal": {"done": 1}, "suppliers": {"failed": 1}}
    s.close()


def test_thread_safe_writes(tmp_path):
    s = State(tmp_path / "s.db")

    def work(n: int) -> None:
        for i in range(50):
            s.mark_message("a@x", "f", f"{n}-{i}", "INBOX", None, STATUS_DONE)

    threads = [threading.Thread(target=work, args=(n,)) for n in range(4)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert len(s.done_message_ids("a@x", "f")) == 200
    s.close()


def test_concurrent_reads_and_writes_from_many_threads_are_consistent(tmp_path):
    """The sqlite3 module is not thread-safe: rows must be fetched under the same lock
    as the statement. A live run saw COUNT(*) come back as None from a racing cursor."""
    import threading

    from o365_to_mailcow.state import STATUS_DONE, State

    st = State(tmp_path / "s.db")
    errors: list[BaseException] = []
    stop = threading.Event()

    def writer(mb: str) -> None:
        i = 0
        while not stop.is_set():
            st.mark_message(mb, "f", f"g{i}", "INBOX", f"<{i}@{mb}>", STATUS_DONE,
                            dest_uid=i, uidvalidity=1)
            i += 1

    def reader(mb: str) -> None:
        try:
            while not stop.is_set():
                n = st.done_count_for_message_id(mb, "f", "<0@" + mb + ">")
                assert isinstance(n, int)
                ids = st.done_message_ids(mb, "f")
                assert all(isinstance(g, str) for g in ids)
                st.message_counts_by_folder_id(mb)
                st.done_message_id_counts(mb, "f")
        except BaseException as exc:  # noqa: BLE001 - collected for the assertion
            errors.append(exc)

    threads = [threading.Thread(target=writer, args=(f"m{k}",)) for k in range(3)]
    threads += [threading.Thread(target=reader, args=(f"m{k}",)) for k in range(3)]
    for t in threads:
        t.start()
    import time
    time.sleep(1.0)
    stop.set()
    for t in threads:
        t.join(10)
    st.close()
    assert errors == []


# -- labels of failed items --------------------------------------------------------------

def test_label_round_trip_for_a_failed_message(tmp_path):
    s = State(tmp_path / "s.db")
    s.mark_message("a@x", "f1", "g1", "INBOX", None, STATUS_FAILED, error="graph HTTP 500: x")
    s.set_label("a@x", "mail", "f1", "g1", "INBOX", "Invoice", "from b@y · received 2024")
    assert s.failed_items("a@x", "mail") == ([{
        "place": "INBOX", "title": "Invoice", "hint": "from b@y · received 2024",
        "status": "failed", "error": "graph HTTP 500: x"}], 1)
    s.close()


def test_failed_items_without_label_fall_back_per_kind(tmp_path):
    s = State(tmp_path / "s.db")
    s.mark_message("a@x", "f1", "g1", "INBOX", None, STATUS_SKIPPED)
    s.mark_event("a@x", "personal", "UID1", None, STATUS_FAILED, "DAV HTTP 500: ")
    s.mark_contact("a@x", "k1", "suppliers", None, STATUS_FAILED, None)
    s.mark_contact("a@x", "k2", "personal", None, STATUS_DONE)
    blank = {"title": "", "hint": ""}
    assert s.failed_items("a@x", "mail") == (
        [{"place": "INBOX", **blank, "status": "skipped", "error": ""}], 1)
    assert s.failed_items("a@x", "calendar") == (
        [{"place": "personal", **blank, "status": "failed", "error": "DAV HTTP 500: "}], 1)
    assert s.failed_items("a@x", "contacts") == (
        [{"place": "suppliers", **blank, "status": "failed", "error": ""}], 1)
    assert s.failed_items("b@x", "mail") == ([], 0)
    s.close()


def test_failed_items_labels_match_each_kinds_key(tmp_path):
    s = State(tmp_path / "s.db")
    s.mark_event("a@x", "team", "UID1", None, STATUS_FAILED, "x")
    s.set_label("a@x", "calendar", "team", "UID1", "Team Events", "Standup", "starts …")
    s.set_label("a@x", "calendar", "personal", "UID1", "Calendar", "wrong slug", "")
    s.mark_contact("a@x", "k1", "personal", None, STATUS_FAILED, "x")
    s.set_label("a@x", "contacts", "", "k1", "Contacts", "Jane", "jane@y")
    s.set_label("a@x", "mail", "", "k1", "INBOX", "wrong kind", "")
    assert [i["title"] for i in s.failed_items("a@x", "calendar")[0]] == ["Standup"]
    assert [i["title"] for i in s.failed_items("a@x", "contacts")[0]] == ["Jane"]
    s.close()


def test_failed_items_limit_total_and_order(tmp_path, monkeypatch):
    clock = iter(range(100, 200))
    monkeypatch.setattr("o365_to_mailcow.state.time.time", lambda: next(clock))
    s = State(tmp_path / "s.db")
    for gid in ("g3", "g1", "g2"):  # written in this order: oldest first, not by id
        s.mark_message("a@x", "f1", gid, "INBOX", None, STATUS_FAILED, error=gid)
    items, total = s.failed_items("a@x", "mail", limit=2)
    assert total == 3 and [i["error"] for i in items] == ["g3", "g1"]
    assert s.failed_items("a@x", "mail", limit=0) == ([], 3)
    s.close()


def test_item_copied_later_is_not_listed_and_its_label_is_deleted(tmp_path):
    s = State(tmp_path / "s.db")
    s.mark_message("a@x", "f1", "g1", "INBOX", None, STATUS_FAILED, error="x")
    s.set_label("a@x", "mail", "f1", "g1", "INBOX", "Secret subject", "")
    s.mark_message("a@x", "f1", "g1", "INBOX", None, STATUS_DONE, dest_uid=1, uidvalidity=1)
    assert s.failed_items("a@x", "mail") == ([], 0)
    assert s._rows("SELECT * FROM item_labels") == []
    s.close()


def test_label_texts_trimmed_to_300(tmp_path):
    s = State(tmp_path / "s.db")
    s.mark_contact("a@x", "k1", "personal", None, STATUS_FAILED, "x")
    s.set_label("a@x", "contacts", "", "k1", "p" * 400, "t" * 400, "h" * 400)
    item = s.failed_items("a@x", "contacts")[0][0]
    assert (len(item["place"]), len(item["title"]), len(item["hint"])) == (300, 300, 300)
    s.set_label("a@x", "contacts", "", "k1", None, None, None)
    assert s.failed_items("a@x", "contacts")[0][0]["title"] == ""
    s.close()


def test_failed_items_unknown_kind_raises(tmp_path):
    s = State(tmp_path / "s.db")
    with pytest.raises(ValueError, match="unknown item kind"):
        s.failed_items("a@x", "tasks")
    s.close()


def test_database_of_the_previous_version_opens_and_stays_compatible(tmp_path):
    path = tmp_path / "old.db"
    old = sqlite3.connect(path)
    old.executescript("""
        CREATE TABLE messages (
            mailbox TEXT NOT NULL, folder_id TEXT NOT NULL, graph_id TEXT NOT NULL,
            folder TEXT NOT NULL, message_id TEXT, status TEXT NOT NULL, error TEXT,
            dest_uid INTEGER, uidvalidity INTEGER, updated_at REAL NOT NULL,
            PRIMARY KEY (mailbox, folder_id, graph_id));
        CREATE TABLE events (
            mailbox TEXT NOT NULL, calendar TEXT NOT NULL, uid TEXT NOT NULL,
            last_modified TEXT, status TEXT NOT NULL, error TEXT, updated_at REAL NOT NULL,
            PRIMARY KEY (mailbox, calendar, uid));
        CREATE TABLE contacts (
            mailbox TEXT NOT NULL, graph_id TEXT NOT NULL, book TEXT NOT NULL,
            last_modified TEXT, status TEXT NOT NULL, error TEXT, updated_at REAL NOT NULL,
            PRIMARY KEY (mailbox, graph_id));
        INSERT INTO events VALUES ('a@x', 'personal', 'UID1', NULL, 'failed', 'DAV HTTP 500: ', 1);
    """)
    old.commit()
    old.close()
    s = State(path)
    assert s.failed_items("a@x", "calendar") == ([{
        "place": "personal", "title": "", "hint": "", "status": "failed",
        "error": "DAV HTTP 500: "}], 1)
    s.close()
    # the previous version's positional insert still fits the (unchanged) messages table
    old = sqlite3.connect(path)
    old.execute("INSERT OR REPLACE INTO messages VALUES (?,?,?,?,?,?,?,?,?,?)",
                ("a@x", "f1", "g1", "INBOX", None, "done", None, 1, 1, 1.0))
    old.commit()
    old.close()
