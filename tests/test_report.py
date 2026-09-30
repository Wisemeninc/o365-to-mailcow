"""Progress output, JSON run report and verify summary (ISC-114, 118, 119, 128)."""

from __future__ import annotations

import io
import json
import stat
import threading
from datetime import UTC, datetime

from o365_to_mailcow.mail import FolderVerify, MailVerify
from o365_to_mailcow.report import (
    CollectionsVerify,
    CollectionVerify,
    Progress,
    RunReport,
    _plain,
    verify_summary,
)


def test_progress_prints_done_total_rate_every_interval_isc_118():
    clock = {"t": 0.0}
    out = io.StringIO()
    p = Progress(out=out, interval=30.0, clock=lambda: clock["t"])
    p.start("a@x mail", 100)
    for _ in range(10):
        clock["t"] += 2
        p.advance("a@x mail")
    assert out.getvalue() == ""  # 20 s: not yet
    clock["t"] = 30
    p.advance("a@x mail")
    assert out.getvalue().strip() == "[a@x mail] 11/100 items, 22/min"
    p.finish("a@x mail")
    assert out.getvalue().splitlines()[-1].endswith("(finished)")


def test_progress_ticker_prints_without_advances():
    out = io.StringIO()
    p = Progress(out=out, interval=0.01)
    p.start("a@x calendar", 5)
    stop = threading.Event()
    t = threading.Thread(target=p.run_ticker, args=(stop,))
    t.start()
    threading.Event().wait(0.08)
    stop.set()
    t.join()
    assert "[a@x calendar] 0/5 items" in out.getvalue()


def test_run_report_written_with_timestamp_and_mode_isc_114(tmp_path):
    now = datetime(2026, 9, 30, 20, 15, 0, tzinfo=UTC)
    r = RunReport("migrate", tmp_path, now=lambda: now)
    r.mailbox("a@x", "a@y")
    r.set("a@x", "mail", MailVerify(mailbox="a@x", folders=[
        FolderVerify("INBOX", 3, 3, 0, 0, 3, 3, False)]))
    r.error("a@x", "boom")
    path = r.write(1)
    assert path == tmp_path / "reports" / "20260930T201500Z.json"
    assert stat.S_IMODE(path.stat().st_mode) == 0o600
    data = json.loads(path.read_text())
    assert data["command"] == "migrate" and data["exit_code"] == 1
    assert data["mailboxes"]["a@x"]["mail"]["folders"][0]["dest_name"] == "INBOX"
    assert data["mailboxes"]["a@x"]["errors"] == ["boom"]
    assert "duration_s" in data and data["started"].startswith("2026-09-30T20:15")
    second = RunReport("verify", tmp_path, now=lambda: now).write(0)
    assert second != path and second.exists()


def test_plain_handles_nested_types(tmp_path):
    assert _plain({"a": (1, {2}), "p": tmp_path}) == {"a": [1, [2]], "p": str(tmp_path)}


def _entry(**sections):
    return {"a@x": {"destination": "a@y", "status": "ok", "errors": [], **sections}}


def _mail(folders, skipped=(), **kw):
    v = MailVerify(mailbox="a@x", folders=list(folders), skipped_folders=list(skipped), **kw)
    return _plain(v)


def test_verify_all_clear_only_without_problems_isc_128():
    ok = _mail([FolderVerify("INBOX", 3, 3, 0, 0, 3, 3, False)],
               skipped=[{"path": "Outbox", "reason": "well-known folder outbox", "total": 0}])
    cal = _plain(CollectionsVerify("a@x", "calendar", [
        CollectionVerify("Calendar", "personal", 2, 2, 0, 2, 2, False)],
        skipped=["Bob (shared by bob@x)"]))
    lines, problems = verify_summary(_entry(mail=ok, calendar=cal))
    assert problems == 0
    assert lines[-1].startswith("All counts match")
    assert any("Outbox" in line for line in lines)  # empty skipped folder still listed
    assert any("shared by bob" in line for line in lines)


def test_verify_lists_every_skipped_and_failed_category_isc_119_128():
    mail = _mail(
        [FolderVerify("INBOX", 10, 7, 2, 1, 7, 7, False),
         FolderVerify("Sent", 4, 4, 0, 0, 3, 4, True)],
        skipped=[{"path": "Conversation History", "reason": "well-known folder "
                  "conversationhistory", "total": 12}],
        sample_requested=3, sample_checked=3, sample_mismatches=["INBOX: <m@x>"])
    contacts = _plain(CollectionsVerify("a@x", "contacts", [
        CollectionVerify("Contacts", "personal", 5, 4, 1, 4, 4, False)]))
    lines, problems = verify_summary(_entry(mail=mail, contacts=contacts))
    text = "\n".join(lines)
    assert "INBOX: 2 failed" in text and "INBOX: 1 skipped (too large)" in text
    assert "Sent: expected 4, IMAP has 3" in text
    assert "Conversation History" in text and "12 messages" in text
    assert "sample SHA-256 mismatch: INBOX: <m@x>" in text
    assert "contacts Contacts: 1 failed" in text
    assert problems == 2 + 1 + 1 + 12 + 1 + 1
    assert "All counts match" not in text and lines[-1].startswith("VERIFY FAILED")
    # the table shows graph, skipped, failed, expected and IMAP numbers (ISC-119)
    row = next(line for line in lines if line.strip().startswith("INBOX"))
    assert row.split()[1:6] == ["10", "1", "2", "7", "7"]


def test_verify_missing_mailbox_and_errors_are_problems():
    entry = {"a@x": {"destination": "a@y", "status": "missing", "errors": ["gone"]}}
    lines, problems = verify_summary(entry)
    assert problems == 2 and "does not exist in mailcow" in "\n".join(lines)
