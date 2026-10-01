"""State ledger: idempotency keys, persistence, file mode (ISC-101, 120, 121, 125, 130)."""

from __future__ import annotations

import stat
import threading

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
