"""The page's view model of a run report (``summary.summarize_report``).

Fixtures are built from the real result dataclasses through ``report._plain`` (exactly what
``RunReport.set`` stores), so a change of a report shape breaks these tests instead of the page.
"""

from __future__ import annotations

import json

import pytest

from o365_to_mailcow.mail import FolderResult, FolderVerify, MailPlan, MailResult, MailVerify
from o365_to_mailcow.report import (
    CollectionPlan,
    CollectionResult,
    CollectionsPlan,
    CollectionsResult,
    CollectionsVerify,
    CollectionVerify,
    _plain,
)
from o365_to_mailcow.summary import (
    MAX_DETAIL,
    MAX_MAILBOXES,
    summarize_report,
    summarize_with_step,
    unreadable_summary,
)

ANNA, BEN = "anna@example.com", "ben@example.com"


FULL = {"only": None, "mailbox": None, "mail_since": None}  # what cli.Runner records


def report(command: str, mailboxes: object, exit_code: object = 0, dry_run: bool = False,
           scope: object = FULL) -> dict:
    data = {"command": command, "dry_run": dry_run, "started": "2026-10-01T01:00:00+00:00",
            "finished": "2026-10-01T02:00:00+00:00", "duration_s": 3600.5,
            "exit_code": exit_code, "mailboxes": mailboxes, "scope": scope}
    if scope is None:  # a report written before scopes were recorded
        del data["scope"]
    return data


def entry(status: str = "ok", errors: list | None = None, **sections) -> dict:
    return {"destination": ANNA.replace("example.com", "example.net"), "status": status,
            "errors": errors or [], **{k: _plain(v) for k, v in sections.items()}}


def folder(name: str = "INBOX", graph: int = 10, imap: int = 10, expected: int | None = None,
           failed: int = 0, skipped: int = 0, **kw) -> FolderVerify:
    expected = imap if expected is None else expected
    return FolderVerify(name, graph, graph - failed - skipped, failed, skipped, imap, expected,
                        imap != expected, **kw)


def mail_verify(*folders: FolderVerify, **kw) -> MailVerify:
    return MailVerify(mailbox=ANNA, folders=list(folders), **kw)


def collection(name: str = "Calendar", graph: int = 2, dav: int = 2, expected: int | None = None,
               failed: int = 0) -> CollectionVerify:
    expected = dav if expected is None else expected
    return CollectionVerify(name, "personal", graph, graph - failed, failed, dav, expected,
                            dav != expected)


def only(summary: dict) -> dict:
    assert len(summary["mailboxes"]) == 1
    return summary["mailboxes"][0]


# -- verify ----------------------------------------------------------------------------------

def test_verify_complete_is_everything_arrived():
    cal = CollectionsVerify(ANNA, "calendar", [collection()], skipped=["Bob (shared by bob)"])
    s = summarize_report(report("verify", {ANNA: entry(
        mail=mail_verify(folder("INBOX", 1200, 1200), folder("Sent", 34, 34)), calendar=cal)}))
    assert s["headline"] == {"level": "ok", "text": "Everything arrived",
                             "detail": "1,236 of 1,236 items are in mailcow. Items skipped by "
                                       "design are listed per mailbox."}
    assert s["columns"] == ["mail", "calendar"]
    row = only(s)
    assert row["key"] == ANNA and row["destination"] == "anna@example.net"
    assert row["result"] == {"level": "ok", "text": "Complete"}
    assert row["cells"]["mail"] == {"primary": "1,234 of 1,234", "secondary": "complete",
                                    "level": "ok"}
    assert row["cells"]["contacts"] is None
    mail_card, cal_card = row["detail"]["cards"]
    assert mail_card == {"title": "Mail", "rows": [["In Microsoft 365", "1,234"],
                                                   ["In mailcow", "1,234"]],
                         "verdict": {"level": "ok", "text": "Complete"},
                         "note": "2 folders checked"}
    # shared calendars are skipped by design: listed, never a problem
    assert cal_card["note"] == "1 calendar checked · 1 shared calendar not copied"
    assert row["detail"]["skipped"][0]["level"] == "ok"
    assert row["detail"]["problems"] == [] and row["detail"]["sample"] is None


def test_verify_counts_that_contradict_the_mismatch_flag_are_never_ok():
    """The tool derives ``mismatch`` from the two counts; a row where they disagree was not
    written by it, and ``verify_summary`` (which trusts the flag) would call it clean."""
    e = entry(mail=mail_verify(folder("INBOX", 10, 7, expected=10)))
    e["mail"]["folders"][0]["mismatch"] = False
    s, step = summarize_with_step(report("verify", {ANNA: e}, exit_code=0))
    assert s["headline"]["level"] != "ok" and step["level"] != "ok"
    assert only(s)["result"]["level"] != "ok"
    assert only(s)["cells"]["mail"]["level"] != "ok"


def test_verify_status_ok_with_a_mismatch_is_not_ok():
    """The bug the redesign exists to fix: status "ok" only means the sections exist."""
    s = summarize_report(report("verify", {ANNA: entry(
        mail=mail_verify(folder("INBOX", 10, 7, expected=10)))}, exit_code=1))
    assert s["headline"]["level"] == "warn"
    assert s["headline"]["text"] == "3 items have not arrived in mailcow"
    assert s["headline"]["detail"] == "7 of 10 items are at the destination."
    row = only(s)
    assert row["result"] == {"level": "warn", "text": "3 missing"}
    assert row["cells"]["mail"] == {"primary": "7 of 10", "secondary": "3 missing",
                                    "level": "warn"}
    assert row["detail"]["problems"] == ["mail INBOX: expected 10, IMAP has 7"]
    assert row["detail"]["differences"] == [
        {"kind": "Mail", "name": "INBOX", "source": "10", "destination": "7",
         "delta": "3 missing", "why": ""}]
    assert summarize_with_step(report("verify", {ANNA: entry(
        mail=mail_verify(folder("INBOX", 10, 7, expected=10)))}, exit_code=1))[1]["text"] \
        == "3 items missing"


def test_verify_missing_and_extra_are_counted_per_row():
    s = summarize_report(report("verify", {ANNA: entry(mail=mail_verify(
        folder("A", 10, 7, expected=10), folder("B", 5, 8, expected=5, note="copied twice")))},
        exit_code=1))
    row = only(s)
    assert row["cells"]["mail"]["secondary"] == "3 missing · 3 more than expected"
    assert [d["delta"] for d in row["detail"]["differences"]] == ["3 missing",
                                                                 "3 more than expected"]
    assert row["detail"]["differences"][1]["why"] == "copied twice"
    assert s["headline"]["detail"].endswith("3 more than expected.")


def test_verify_failed_and_too_large_in_contract_order():
    s = summarize_report(report("verify", {ANNA: entry(mail=mail_verify(
        folder("INBOX", 20, 14, expected=17, failed=2, skipped=1)))}, exit_code=1))
    row = only(s)
    assert row["cells"]["mail"]["secondary"] == "3 missing · 2 failed · 1 too large"
    assert row["detail"]["differences"][0]["delta"] == "3 missing · 2 failed · 1 too large"
    assert s["headline"]["detail"] == ("14 of 20 items are at the destination. "
                                       "2 failed, 1 too large.")


def test_verify_failures_without_missing_items_count_as_problems():
    s = summarize_report(report("verify", {ANNA: entry(mail=mail_verify(
        folder("INBOX", 10, 9, failed=1)))}, exit_code=1))
    assert only(s)["result"] == {"level": "warn", "text": "1 problem"}
    assert s["headline"]["text"] == "Verify found 1 problem"
    assert summarize_with_step(report("verify", {ANNA: entry(mail=mail_verify(
        folder("INBOX", 10, 9, failed=1)))}, exit_code=1))[1]["text"] == "1 problem"


def test_verify_skipped_folders_with_items_are_warnings():
    mail = mail_verify(folder(), skipped_folders=[
        {"path": "Sync Issues/Conflicts", "reason": "parent skipped", "total": 4},
        {"path": "Outbox", "reason": "well-known folder outbox", "total": 0}])
    s = summarize_report(report("verify", {ANNA: entry(mail=mail)}, exit_code=1))
    row = only(s)
    assert row["detail"]["skipped"] == [
        {"kind": "Mail", "name": "Sync Issues/Conflicts", "reason": "parent skipped",
         "items": "4", "level": "warn"},
        {"kind": "Mail", "name": "Outbox", "reason": "well-known folder outbox", "items": "0",
         "level": "ok"}]
    assert row["result"] == {"level": "warn", "text": "4 problems"}
    assert row["cells"]["mail"]["secondary"] == "4 in skipped folders"
    assert s["headline"]["detail"].endswith("4 in skipped folders.")


def test_verify_contacts_fallbacks_are_problems():
    contacts = CollectionsVerify(ANNA, "contacts", [collection("Contacts")],
                                 fallbacks=["Friends: could not create the address book"])
    s = summarize_report(report("verify", {ANNA: entry(contacts=contacts)}, exit_code=1))
    row = only(s)
    assert row["cells"]["contacts"] == {"primary": "2 of 2", "secondary": "1 error",
                                        "level": "warn"}
    assert row["result"]["level"] == "warn"
    assert s["headline"]["level"] == "warn"
    assert row["detail"]["problems"] == [
        "contacts fallback (address book not created): Friends: could not create the "
        "address book"]


def test_verify_section_errors_reach_the_cell_and_the_error_list():
    cal = CollectionsVerify(ANNA, "calendar", [collection()], errors=["DAV 500"])
    row = only(summarize_report(report("verify", {ANNA: entry(calendar=cal)}, exit_code=1)))
    assert row["cells"]["calendar"]["secondary"] == "1 error"
    assert row["detail"]["errors"] == ["calendar: DAV 500"]


@pytest.mark.parametrize("mismatches, level, text", [
    ([], "ok", "20 of 20 sampled messages match"),
    (["INBOX: <m@x>"], "warn", "19 of 20 sampled messages match"),
])
def test_verify_sample(mismatches, level, text):
    mail = mail_verify(folder(), sample_requested=20, sample_checked=20,
                       sample_mismatches=mismatches, sample_regenerated=2)
    sample = only(summarize_report(report("verify", {ANNA: entry(mail=mail)})))["detail"][
        "sample"]
    assert sample["level"] == level and sample["text"] == text
    assert "2 re-rendered by Exchange" in sample["detail"]


def test_verify_missing_mailbox_and_exit_code_2_are_bad():
    missing = {"destination": "anna@example.net", "status": "missing",
               "errors": ["destination mailbox anna@example.net does not exist in mailcow"]}
    s, step = summarize_with_step(report("verify", {ANNA: missing}, exit_code=1))
    assert only(s)["result"] == {"level": "bad", "text": "No mailbox in mailcow"}
    assert s["headline"] == {"level": "bad", "text": "Verify could not finish",
                             "detail": missing["errors"][0]}
    assert step["text"] == "Could not finish" and step["level"] == "bad"
    for status in ("failed", "pending"):
        row = only(summarize_report(report("verify", {ANNA: entry(status=status)}, 1)))
        assert row["result"] == {"level": "bad", "text": "Did not finish"}
    s = summarize_report(report("verify", {ANNA: entry(mail=mail_verify(folder()))}, 2))
    assert s["headline"]["level"] == "bad"


def test_verify_nonzero_exit_without_problems_is_unknown():
    s = summarize_report(report("verify", {ANNA: entry(mail=mail_verify(folder()))}, 1))
    assert s["headline"] == {"level": "unknown", "text": "Verify ended with exit code 1",
                             "detail": ""}
    assert summarize_report(report("verify", {ANNA: entry(mail=mail_verify(folder()))},
                                    None))["headline"]["text"] == "Verify ended"
    assert summarize_report(report("verify", {}))["headline"]["level"] == "unknown"


def test_verify_headline_counts_every_mailbox_not_only_the_shown_ones():
    boxes = {f"user{i:04d}@example.com": entry(mail=mail_verify(folder("INBOX", 2, 1,
                                                                       expected=2)))
             for i in range(MAX_MAILBOXES + 1)}
    s = summarize_report(report("verify", boxes, exit_code=1))
    assert len(s["mailboxes"]) == MAX_MAILBOXES and s["more_mailboxes"] == 1
    assert s["headline"]["text"] == "501 items have not arrived in mailcow"
    assert s["mailboxes"][0]["key"] == "user0000@example.com"  # sorted by key


# -- migrate ---------------------------------------------------------------------------------

def migrate_entry(status: str = "ok", stopped: bool = False, failed: int = 0,
                  error: str | None = None, dry_run: bool = False, errors=None) -> dict:
    mail = MailResult(ANNA, dry_run, [
        FolderResult("f1", "INBOX", listed=1300, appended=0 if dry_run else 1200,
                     already_done=30, dedup_hits=4, failed=failed, skipped_too_large=2,
                     would_append=1200 if dry_run else 0, error=error)],
        skipped_folders=["Sync Issues"], stopped=stopped)
    cal = CollectionsResult(ANNA, "calendar", dry_run, [
        CollectionResult("c1", "Calendar", "personal", listed=40, put=0 if dry_run else 34,
                         unchanged=6, would_put=34 if dry_run else 0)])
    return entry(status=status, errors=errors, mail=mail, calendar=cal)


def test_migrate_real_run_ok():
    s, step = summarize_with_step(report("migrate", {ANNA: migrate_entry()}))
    assert s["headline"] == {"level": "ok", "text": "Migrate finished: 1,234 items copied",
                             "detail": "40 were already there. 2 too large were skipped."}
    assert step["text"] == "1,234 items copied" and step["level"] == "ok"
    row = only(s)
    assert row["result"] == {"level": "ok", "text": "Copied"}
    assert row["cells"]["mail"] == {"primary": "1,200 copied",
                                    "secondary": "34 already there · 2 too large",
                                    "level": "ok"}
    assert row["cells"]["calendar"]["secondary"] == "6 already there"
    assert row["detail"]["cards"][0]["rows"] == [["Copied", "1,200"], ["Already there", "34"],
                                                 ["Failed", "0"], ["Too large", "2"]]
    assert row["detail"]["skipped"] == [{"kind": "Mail", "name": "Sync Issues", "reason": "",
                                         "items": "", "level": "ok"}]
    assert row["detail"]["problems"] == []


def test_migrate_failures_are_bad():
    s, step = summarize_with_step(report("migrate", {ANNA: migrate_entry(
        status="failed", failed=3, error="IMAP APPEND refused")}, exit_code=1))
    row = only(s)
    assert row["result"] == {"level": "bad", "text": "4 failed"}
    assert row["cells"]["mail"]["level"] == "bad"
    assert row["detail"]["differences"] == [
        {"kind": "Mail", "name": "INBOX", "source": "", "destination": "",
         "delta": "4 failed", "why": "IMAP APPEND refused"}]
    assert s["headline"]["text"] == "Migrate finished with 4 failures"
    assert s["headline"]["detail"] == "1,234 items copied, 40 already there."
    assert step["text"] == "4 failures"


def test_migrate_errors_without_counts_and_stopped_and_pending():
    row = only(summarize_report(report("migrate", {ANNA: migrate_entry(
        status="failed", errors=["RuntimeError: boom"])}, exit_code=1)))
    assert row["result"] == {"level": "bad", "text": "Failed"}
    s, step = summarize_with_step(report("migrate", {ANNA: migrate_entry(stopped=True)}, 1))
    assert only(s)["result"] == {"level": "bad", "text": "Stopped early"}
    assert s["headline"]["text"] == "Migrate did not finish" and step["text"] == "Did not finish"
    row = only(summarize_report(report("migrate", {ANNA: entry(status="pending")}, 1)))
    assert row["result"] == {"level": "bad", "text": "Did not finish"}


def test_migrate_dry_run():
    s = summarize_report(report("migrate", {ANNA: migrate_entry(dry_run=True)}, dry_run=True))
    assert s["dry_run"] is True
    assert s["headline"]["text"] == "Dry run: 1,234 items would be copied"
    row = only(s)
    assert row["result"] == {"level": "ok", "text": "Dry run"}
    assert row["cells"]["mail"] == {"primary": "1,200 to copy",
                                    "secondary": "34 already there · 2 too large",
                                    "level": "ok"}
    assert row["detail"]["cards"][0]["rows"][0] == ["To copy", "1,200"]


# -- plan ------------------------------------------------------------------------------------

def plan_entry(status: str = "ok", errors=None) -> dict:
    mail = MailPlan(ANNA, "/", [], total_messages=1500, total_bytes=10, skipped_folders=2)
    cal = CollectionsPlan(ANNA, "calendar", [
        CollectionPlan("c1", "Calendar", "personal", 40),
        CollectionPlan("c2", "Birthdays", "birthdays", 9, skip=True, skip_reason="generated")])
    return entry(status=status, errors=errors, mail_plan=mail, calendar_plan=cal)


def test_plan_ok_and_missing_mailbox():
    s = summarize_report(report("plan", {ANNA: plan_entry()}, dry_run=True))
    assert s["headline"] == {"level": "ok", "text": "Plan: 1,540 items in 1 mailbox",
                             "detail": ""}
    row = only(s)
    assert row["result"] == {"level": "ok", "text": "Ready"}
    assert row["cells"]["calendar"] == {"primary": "40 to copy", "secondary": "", "level": "ok"}
    assert row["detail"]["cards"][0]["rows"] == [["Items", "1,500"]]
    missing = plan_entry(status="missing", errors=["destination mailbox does not exist"])
    s = summarize_report(report("plan", {ANNA: missing, BEN: plan_entry()}, 1, True))
    assert s["headline"] == {"level": "warn",
                             "text": "1 mailbox does not exist in mailcow yet",
                             "detail": "Run Provision to create them."}
    assert s["mailboxes"][0]["result"] == {"level": "bad", "text": "No mailbox in mailcow"}
    s = summarize_report(report("plan", {ANNA: plan_entry(status="failed", errors=["x"])}, 1))
    assert s["headline"] == {"level": "bad", "text": "Plan could not finish", "detail": "x"}


# -- provision and cleanup ------------------------------------------------------------------

def test_provision_rows_and_aliases_note():
    boxes = {ANNA: {"errors": [], "provision": "created"},
             BEN: {"errors": [], "provision": "exists"},
             "aliases": {"errors": [], "created": 3}}
    s, step = summarize_with_step(report("provision", boxes))
    assert [m["key"] for m in s["mailboxes"]] == [ANNA, BEN]  # aliases is not a mailbox
    assert [m["result"] for m in s["mailboxes"]] == [{"level": "ok", "text": "Created"},
                                                    {"level": "ok", "text": "Already exists"}]
    assert s["mailboxes"][0]["destination"] is None and s["columns"] == []
    assert s["notes"] == ["Aliases created: 3"]
    assert s["headline"]["text"] == "Provision: 1 created, 1 already existed"
    assert step["text"] == "1 created, 1 existed"


def test_provision_dry_run_and_failures():
    boxes = {ANNA: {"errors": [], "provision": "would create"},
             "aliases": {"errors": [], "would_create": 2}}
    s = summarize_report(report("provision", boxes, dry_run=True))
    assert s["headline"]["text"] == "Dry run: 1 would be created, 0 already exist"
    assert only(s)["result"]["text"] == "Would be created"
    assert s["notes"] == ["Aliases that would be created: 2"]
    boxes = {ANNA: {"errors": ["provisioning failed: HTTP 500"]},
             BEN: {"errors": [], "provision": "created"}}
    s, step = summarize_with_step(report("provision", boxes, exit_code=1))
    assert s["mailboxes"][0]["result"] == {"level": "bad", "text": "Failed"}
    assert s["headline"] == {"level": "bad", "text": "Provision failed for 1 mailbox",
                             "detail": "provisioning failed: HTTP 500"}
    assert step["text"] == "1 failed"
    boxes = {ANNA: {"errors": [], "provision": "created"},
             "aliases": {"errors": ["cannot list aliases: HTTP 500"]}}
    s = summarize_report(report("provision", boxes, exit_code=1))
    assert s["notes"] == ["Aliases: cannot list aliases: HTTP 500"]
    assert s["headline"]["level"] == "bad"


def test_cleanup_is_keyed_by_destination():
    boxes = {"anna@example.net": {"errors": [], "deleted_app_passwords": 2},
             "ben@example.net": {"errors": [], "deleted_app_passwords": 1}}
    s, step = summarize_with_step(report("cleanup", boxes))
    assert [(m["key"], m["destination"]) for m in s["mailboxes"]] == [
        ("anna@example.net", "anna@example.net"), ("ben@example.net", "ben@example.net")]
    assert s["mailboxes"][1]["result"] == {"level": "ok", "text": "1 app password deleted"}
    assert s["headline"]["text"] == "Clean up: 3 app passwords deleted"
    assert step["text"] == "3 deleted"
    boxes["ben@example.net"] = {"errors": ["mailcow API HTTP 500"]}
    s, step = summarize_with_step(report("cleanup", boxes, exit_code=1))
    assert s["mailboxes"][1]["result"] == {"level": "bad", "text": "Failed"}
    assert s["headline"] == {"level": "bad", "text": "Clean up failed for 1 mailbox",
                             "detail": "mailcow API HTTP 500"}
    assert step["text"] == "1 failed"


# -- malformed input -------------------------------------------------------------------------

GOOD_VERIFY = entry(mail=mail_verify(folder()))
MALFORMED = [
    None, [], "x", 7, {}, {"command": "verify"},
    report("verify", ["not", "a", "dict"]),
    report("verify", "x"),
    report("rm -rf", {ANNA: GOOD_VERIFY}),
    report(None, {ANNA: GOOD_VERIFY}),
    report("verify", {ANNA: "a string"}),
    report("verify", {ANNA: None}),
    report("verify", {ANNA: [1, 2]}),
    report("verify", {ANNA: {**GOOD_VERIFY, "mail": "x"}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "mail": {"folders": "x"}}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "mail": {"folders": [1, None]}}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "mail": {"folders": [
        {**_plain(folder()), "imap_count": "10"}]}}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "mail": {"folders": [
        {**_plain(folder()), "graph_total": None}]}}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "mail": {"folders": [
        {**_plain(folder()), "imap_count": -5}]}}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "mail": {"folders": [
        {**_plain(folder()), "graph_total": 10**15 + 1}]}}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "mail": {"folders": [
        {**_plain(folder()), "expected": -(10**15) - 1}]}}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "mail": {"folders": [
        {**_plain(folder()), "failed": True}]}}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "mail": {"folders": [
        {**_plain(folder()), "expected": float("nan")}]}}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "errors": "boom"}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "calendar": {"collections": {"a": 1}}}}),
    report("verify", {ANNA: {**GOOD_VERIFY, "status": "weird"}}),
    report("verify", {ANNA: {"status": "ok", "errors": []}}),
    report("migrate", {ANNA: {"status": "ok", "errors": [], "mail": {"folders": [{}]}}}),
    report("migrate", {ANNA: {"status": "ok", "errors": [], "mail": [1]}}),
    report("plan", {ANNA: {"status": "ok", "errors": [], "mail_plan": {"total_messages": "9"}}}),
    report("provision", {ANNA: 5, "aliases": "x"}),
    report("provision", {"aliases": {"errors": [], "created": "3"}}),
    report("cleanup", {"anna@example.net": {"errors": [], "deleted_app_passwords": "2"}}),
    report("cleanup", {"anna@example.net": None}),
]


@pytest.mark.parametrize("data", MALFORMED, ids=range(len(MALFORMED)))
def test_malformed_reports_never_raise_and_are_never_ok(data):
    summary, step = summarize_with_step(data)
    assert summary["headline"]["level"] != "ok"
    assert step["level"] == summary["headline"]["level"]
    for box in summary["mailboxes"]:
        assert box["result"]["level"] != "ok"
    json.dumps(summary)  # always serialisable


@pytest.mark.parametrize("code", ["0", True, 0.0, None])
def test_an_unreadable_exit_code_is_never_a_green_headline(code):
    s = summarize_report(report("verify", {ANNA: GOOD_VERIFY}, exit_code=code))
    assert s["exit_code"] is None
    assert s["headline"] == {"level": "unknown", "text": "Verify ended", "detail": ""}
    assert only(s)["result"]["level"] == "ok"  # the mailbox data itself is fine


def test_unknown_report_shape():
    assert summarize_report(None) == {
        "command": "", "dry_run": False, "started": None, "finished": None,
        "duration_s": None, "exit_code": None, "partial": False, "scope": "",
        "headline": {"level": "unknown", "text": "Run ended", "detail": ""},
        "columns": [], "mailboxes": [], "more_mailboxes": 0, "notes": []}
    s = summarize_report(report("rm", {ANNA: GOOD_VERIFY}, exit_code=3))
    assert s["headline"]["text"] == "rm ended with exit code 3" and s["mailboxes"] == []


def test_strings_are_cleaned_and_truncated():
    evil = "\x1b[31mINBOX‮" + "x" * 1000
    s = summarize_report(report("verify", {ANNA + "\x07": entry(
        errors=["e" * 1000], mail=mail_verify(folder(evil, 10, 7, expected=10)))}, 1))
    row = only(s)
    assert row["key"] == ANNA
    diff = row["detail"]["differences"][0]
    assert diff["name"].startswith("[31mINBOX") and "\x1b" not in diff["name"]
    assert "‮" not in diff["name"] and len(diff["name"]) == 300
    assert all(len(p) <= 300 for p in row["detail"]["problems"])
    assert row["detail"]["errors"] == ["e" * 300]


def test_detail_lists_are_capped():
    folders = [folder(f"F{i:03d}", 2, 1, expected=2) for i in range(MAX_DETAIL + 1)]
    row = only(summarize_report(report("verify", {ANNA: entry(mail=mail_verify(*folders))}, 1)))
    assert len(row["detail"]["differences"]) == MAX_DETAIL
    assert row["detail"]["more_differences"] == 1
    assert len(row["detail"]["problems"]) == MAX_DETAIL
    assert row["result"] == {"level": "warn", "text": "201 missing"}


def test_step_carries_the_report_times_and_exit_code():
    _, step = summarize_with_step(report("verify", {ANNA: GOOD_VERIFY}))
    assert step == {"started": "2026-10-01T01:00:00+00:00",
                    "finished": "2026-10-01T02:00:00+00:00", "exit_code": 0, "level": "ok",
                    "text": "Everything arrived"}


# -- review-gate follow-up -----------------------------------------------------------------

def test_sample_mismatch_makes_the_mail_cell_warn():
    mail = mail_verify(folder("INBOX", 5, 5), sample_requested=5, sample_checked=5,
                       sample_mismatches=["m1"])
    row = only(summarize_report(report("verify", {ANNA: entry(mail=mail)}, exit_code=1)))
    assert row["cells"]["mail"] == {"primary": "5 of 5", "secondary": "1 sample mismatch",
                                    "level": "warn"}
    assert row["detail"]["cards"][0]["verdict"] == {"level": "warn",
                                                    "text": "1 sample mismatch"}
    assert row["result"] == {"level": "warn", "text": "1 problem"}
    mail.sample_mismatches = ["m1", "m2"]
    row = only(summarize_report(report("verify", {ANNA: entry(mail=mail)}, exit_code=1)))
    assert row["cells"]["mail"]["secondary"] == "2 sample mismatches"


@pytest.mark.parametrize("scope, text", [
    ({"only": "mail", "mailbox": None, "mail_since": None}, "mail only"),
    ({"only": None, "mailbox": ANNA, "mail_since": None}, "1 mailbox"),
    ({"only": "mail", "mailbox": ANNA, "mail_since": "2025-01-01"},
     "mail only · 1 mailbox · mail since 2025-01-01"),
    ({"only": None, "mailbox": None, "mail_since": None, "future": 1}, "a narrowed run"),
    ([1], "the scope could not be read"),
])
def test_a_clean_partial_verify_says_what_it_checked(scope, text):
    summary, step = summarize_with_step(report("verify", {ANNA: entry(
        mail=mail_verify(folder()))}, scope=scope))
    assert summary["partial"] is True and summary["scope"] == text
    assert summary["headline"] == {
        "level": "ok", "text": "Everything that was checked has arrived",
        "detail": f"10 of 10 items are in mailcow. Only part was checked: {text}."}
    assert step["level"] == "unknown" and step["text"] == "Partial check passed"


def test_a_partial_run_with_problems_keeps_its_level_and_step():
    summary, step = summarize_with_step(report("verify", {ANNA: entry(
        mail=mail_verify(folder("INBOX", 10, 7, expected=10)))}, exit_code=1,
        scope={"only": "mail", "mailbox": None, "mail_since": None}))
    assert summary["partial"] is True and summary["headline"]["level"] == "warn"
    assert summary["headline"]["text"] == "3 items have not arrived in mailcow"
    assert step == {**step, "level": "warn", "text": "3 items missing"}


def test_a_full_scope_is_not_partial():
    summary, step = summarize_with_step(report("verify", {ANNA: entry(
        mail=mail_verify(folder()))}))
    assert summary["partial"] is False and summary["scope"] == ""
    assert summary["headline"]["text"] == "Everything arrived" and step["level"] == "ok"


def test_other_commands_partial_ok_gets_a_detail_suffix():
    summary, step = summarize_with_step(report(
        "migrate", {ANNA: migrate_entry()},
        scope={"only": None, "mailbox": ANNA, "mail_since": None}))
    assert summary["headline"]["text"] == "Migrate finished: 1,234 items copied"
    assert summary["headline"]["detail"].endswith(" Partial run: 1 mailbox.")
    assert step["level"] == "unknown" and step["text"] == "Partial run"


@pytest.mark.parametrize("sections, partial, text", [
    ({"mail": mail_verify(folder())}, True, "mail only"),
    ({"mail": mail_verify(folder()), "calendar": CollectionsVerify(ANNA, "calendar",
                                                                    [collection()])},
     True, "mail and calendar only"),
    ({"mail": mail_verify(folder()),
      "calendar": CollectionsVerify(ANNA, "calendar", [collection()]),
      "contacts": CollectionsVerify(ANNA, "contacts", [collection("Contacts")])}, False, ""),
])
def test_legacy_reports_infer_the_scope_from_the_kinds_present(sections, partial, text):
    summary = summarize_report(report("verify", {ANNA: entry(**sections)}, scope=None))
    assert summary["partial"] is partial and summary["scope"] == text
    legacy_cleanup = report("cleanup", {"anna@example.net": {"errors": [],
                                                             "deleted_app_passwords": 1}},
                            scope=None)
    assert summarize_report(legacy_cleanup)["partial"] is False


@pytest.mark.parametrize("value", [None, 5, [], {"x": 1}])
def test_command_is_always_a_string(value):
    assert summarize_report({"command": value, "mailboxes": {}})["command"] == ""


def test_negative_expected_is_real_data_after_source_deletions():
    """mail.py: expected = graph_total - skipped - failed; failed items deleted at the source
    since make it negative. Not malformed: the derived counts are clamped."""
    row = {**_plain(folder("INBOX", 0, 0)), "failed": 2, "expected": -2, "mismatch": True}
    s = summarize_report(report("verify", {ANNA: entry(mail={"folders": [row]})}, 1))
    box = only(s)
    assert box["cells"]["mail"]["level"] == "warn"  # readable, not "unknown"
    assert box["cells"]["mail"]["secondary"] == "2 failed · 2 more than expected"
    assert s["headline"]["level"] == "warn"


def test_absurd_counts_are_not_believed():
    row = {**_plain(folder()), "graph_total": 10**16, "imap_count": 10**16}
    box = only(summarize_report(report("verify", {ANNA: entry(mail={"folders": [row]})})))
    assert box["cells"]["mail"]["level"] == "unknown"
    assert box["result"]["level"] == "unknown"


def test_provision_alias_failures_are_warnings_not_failures():
    boxes = {ANNA: {"errors": ["alias a1@example.net: HTTP 500"], "provision": "created"},
             BEN: {"errors": ["alias b1@example.net: x", "alias b2@example.net: y"],
                   "provision": "exists"},
             "aliases": {"errors": [], "created": 0}}
    s, step = summarize_with_step(report("provision", boxes, exit_code=1))
    assert [m["result"] for m in s["mailboxes"]] == [
        {"level": "warn", "text": "Created · alias failed"},
        {"level": "warn", "text": "Already exists · alias failed"}]
    assert s["headline"]["level"] == "warn"
    assert s["headline"]["text"] == "Provision finished, 3 alias problems"
    assert step["text"] == "3 alias problems" and step["level"] == "warn"
    boxes[ANNA]["errors"].append("provisioning failed: HTTP 500")  # not only aliases
    assert summarize_report(report("provision", boxes, 1))["mailboxes"][0]["result"] == {
        "level": "bad", "text": "Failed"}


def test_unreadable_alias_count_is_not_an_alias_error():
    boxes = {ANNA: {"errors": [], "provision": "created"},
             "aliases": {"errors": [], "created": "3"}}
    s = summarize_report(report("provision", boxes))
    assert s["notes"] == ["Aliases: the result could not be read"]
    assert s["headline"]["level"] == "unknown"
    assert "Aliases created" not in json.dumps(s)


def test_dry_run_cleanup():
    boxes = {"anna@example.net": {"errors": [], "deleted_app_passwords": 0}}
    s, step = summarize_with_step(report("cleanup", boxes, dry_run=True))
    assert s["headline"] == {"level": "ok", "text": "Dry run: no app passwords were deleted",
                             "detail": ""}
    assert only(s)["result"] == {"level": "ok", "text": "Dry run"}


def test_migrate_mail_section_errors_without_failed_rows():
    mail = MailResult(ANNA, False, [FolderResult("f1", "INBOX", listed=1, appended=1)],
                      errors=["Inbox: listing failed", "Sent: listing failed"])
    row = only(summarize_report(report("migrate", {ANNA: entry(status="failed", mail=mail)},
                                       exit_code=1)))
    assert row["cells"]["mail"] == {"primary": "1 copied", "secondary": "2 errors",
                                    "level": "bad"}
    card = row["detail"]["cards"][0]
    assert card["verdict"] == {"level": "bad", "text": "2 errors"}
    assert ["Errors", "2"] in card["rows"]


def test_unreadable_summary_shape():
    s = unreadable_summary()
    assert s["headline"] == {"level": "unknown", "text": "The newest report could not be read",
                             "detail": ""}
    assert s["command"] == "" and s["mailboxes"] == [] and s["partial"] is False
