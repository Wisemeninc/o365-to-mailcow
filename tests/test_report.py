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
    failed_item_lines,
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
    assert "sample content mismatch: INBOX: <m@x>" in text
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


def test_progress_phase_is_shown_while_the_counter_cannot_move():
    out = io.StringIO()
    clock = {"t": 0.0}
    p = Progress(out=out, interval=30.0, clock=lambda: clock["t"])
    p.start("a@x mail", 100)
    p.phase("a@x mail", "listing Inbox: 5000")
    clock["t"] = 31.0
    assert p.maybe_print(force=True)
    assert "[a@x mail] 0/100 items, 0/min (listing Inbox: 5000)" in out.getvalue()
    p.phase("a@x mail", "")
    p.advance("a@x mail", 10)
    clock["t"] = 62.0
    p.maybe_print(force=True)
    assert out.getvalue().splitlines()[-1] == "[a@x mail] 10/100 items, 10/min"
    p.finish("a@x mail")
    assert out.getvalue().splitlines()[-1].endswith("(finished)")


def test_run_report_records_the_scope_only_when_given(tmp_path):
    now = datetime(2026, 9, 30, 20, 15, 0, tzinfo=UTC)
    scope = {"only": "mail", "mailbox": "a@x", "mail_since": "2025-01-01"}
    data = json.loads(RunReport("verify", tmp_path, now=lambda: now, scope=scope).write(0)
                      .read_text())
    assert data["scope"] == scope
    plain = json.loads(RunReport("verify", tmp_path, now=lambda: now).write(0).read_text())
    assert "scope" not in plain


# -- failed item lines ----------------------------------------------------------------------

def _item(**kw):
    return {"place": "INBOX", "title": "Invoice", "hint": "from b@y", "status": "failed",
            "error": "graph HTTP 500: x", **kw}


def test_failed_item_lines_format():
    lines = failed_item_lines({"failed_items": [
        _item(), _item(title="", hint="", error="", status="skipped")], "failed_items_total": 2})
    assert lines == ["    not copied: INBOX · Invoice (from b@y): graph HTTP 500: x",
                     "    not copied: INBOX · (title not recorded): skipped"]


def test_failed_item_lines_cap_and_more_line():
    lines = failed_item_lines({"failed_items": [_item()] * 100, "failed_items_total": 250})
    assert len(lines) == 21 and lines[:20] == [lines[0]] * 20
    assert lines[-1] == "    … and 230 more (the report file lists up to 100)"


def test_failed_item_lines_remove_control_characters():
    line = failed_item_lines({"failed_items": [_item(title="a\x1b[31mb\u202ec\nd")]})[0]
    assert "a[31mbcd" in line and "\x1b" not in line and "\n" not in line


def test_failed_item_lines_tolerate_malformed_input():
    for section in (None, "x", {}, {"failed_items": None}, {"failed_items": "abc"},
                    {"failed_items": [1, 2], "failed_items_total": "many"},
                    {"failed_items": [{"place": 1, "title": 2, "hint": 3, "error": 4}]}):
        failed_item_lines(section)
    assert failed_item_lines({"failed_items": [{"place": 1, "title": None}]}) == [
        "    not copied:  · (title not recorded): "]


def test_failed_items_in_verify_summary_add_lines_not_problems():
    folders = [FolderVerify("INBOX", 3, 2, 1, 0, 2, 2, False)]
    contacts = CollectionsVerify("a@x", "contacts", [
        CollectionVerify("Contacts", "personal", 2, 1, 1, 1, 1, False)])
    plain_lines, plain_problems = verify_summary(_entry(
        mail=_mail(folders), contacts=_plain(contacts)))
    contacts.failed_items, contacts.failed_items_total = [_item(place="Contacts")], 1
    lines, problems = verify_summary(_entry(
        mail=_mail(folders, failed_items=[_item()], failed_items_total=1),
        contacts=_plain(contacts)))
    assert problems == plain_problems == 2
    assert lines[-1] == plain_lines[-1]
    added = [line for line in lines if line not in plain_lines]
    assert added == ["    not copied: INBOX · Invoice (from b@y): graph HTTP 500: x",
                     "    not copied: Contacts · Invoice (from b@y): graph HTTP 500: x"]
    assert not any(line.startswith("  !") for line in added)
