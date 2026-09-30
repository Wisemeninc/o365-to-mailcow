"""Regression tests for the findings of the independent reviews (Cato, Silas, Reviewer A)."""

from __future__ import annotations

import io
import json
import logging
import os
import threading
from datetime import UTC, datetime
from pathlib import Path

import pytest
import requests
import responses

from fakes_o365 import FakeGraph, make_config
from o365_to_mailcow import cli, config, graph, mail
from o365_to_mailcow.calendar_sync import CalendarMigrator
from o365_to_mailcow.config import ConfigError, MailboxMapping
from o365_to_mailcow.dav import DavError, SogoDav
from o365_to_mailcow.mailcow import MailcowApi, MailcowError
from o365_to_mailcow.report import RunReport, clean
from o365_to_mailcow.state import STATUS_DONE, State

# -- Graph client ----------------------------------------------------------------------

class _Tokens:
    def get_token(self) -> str:
        return "tok"

    def invalidate(self) -> None:
        pass


class _Session(requests.Session):
    def __init__(self, status: int, headers: dict | None = None, body: bytes = b"") -> None:
        super().__init__()
        self.status, self.hdrs, self.body, self.calls = status, headers or {}, body, 0

    def get(self, url, **kw):
        self.calls += 1
        resp = requests.Response()
        resp.status_code = self.status
        resp.headers.update(self.hdrs)
        resp._content = self.body
        resp.raw = io.BytesIO(self.body)
        return resp


def test_graph_refuses_redirects_silas_m2():
    g = graph.GraphClient(_Tokens(), session=_Session(302, {"Location": "https://evil/"}))
    with pytest.raises(graph.GraphError) as exc:
        g.get("/me")
    assert exc.value.status == 302


def test_graph_download_abandoned_past_limit_cato_f5():
    g = graph.GraphClient(_Tokens(), session=_Session(200, body=b"x" * 1000))
    with pytest.raises(graph.GraphTooLarge):
        g.get_bytes("/users/a/messages/1/$value", max_bytes=100)
    assert g.get_bytes("/users/a/messages/1/$value", max_bytes=1000) == b"x" * 1000


def test_malformed_retry_after_falls_back_to_backoff_cato_f14():
    sleeps: list[float] = []
    sess = _Session(429, {"Retry-After": "soon-ish"}, b"{}")
    g = graph.GraphClient(_Tokens(), session=sess, max_retries=1, sleep=sleeps.append)
    with pytest.raises(graph.GraphError):
        g.get("/me")
    assert len(sleeps) == 1 and 0 < sleeps[0] <= 60


def test_retry_after_is_clamped():
    sleeps: list[float] = []
    sess = _Session(429, {"Retry-After": "999999"}, b"{}")
    g = graph.GraphClient(_Tokens(), session=sess, max_retries=1, sleep=sleeps.append)
    with pytest.raises(graph.GraphError):
        g.get("/me")
    assert sleeps == [graph.MAX_RETRY_AFTER]


def test_delta_expired_detection():
    assert graph.delta_expired(graph.GraphError(410, "Gone", "/x"))
    assert graph.delta_expired(graph.GraphError(400, "SyncStateNotFound: token", "/x"))
    assert graph.delta_expired(graph.GraphError(404, "link unknown", "/x"))  # any other 4xx
    assert not graph.delta_expired(graph.GraphError(429, "throttled", "/x"))
    assert not graph.delta_expired(graph.GraphError(503, "busy", "/x"))


# -- mailcow and DAV clients -----------------------------------------------------------

@responses.activate
def test_mailcow_refuses_redirects_silas_m2():
    responses.get("https://mail.example.net/api/v1/get/mailbox/a@example.net", status=301,
                  headers={"Location": "http://mail.example.net/api/v1/get/mailbox/a"})
    api = MailcowApi("mail.example.net", "key")
    with pytest.raises(MailcowError) as exc:
        api.mailbox_exists("a@example.net")
    assert exc.value.status == 301 and len(responses.calls) == 1


@responses.activate
def test_dav_refuses_redirects_silas_m2():
    responses.add("PROPFIND", "https://mail.example.net/SOGo/dav/a@example.net/Calendar/",
                  status=302, headers={"Location": "https://other/"})
    dav = SogoDav("mail.example.net", "a@example.net", "pw")
    with pytest.raises(DavError) as exc:
        dav.calendar_home_exists()
    assert exc.value.status == 302


# -- mail helpers ----------------------------------------------------------------------

def test_category_keywords_are_capped_and_never_system_flags_silas_m5():
    assert mail.category_keyword("$Junk") == "_Junk"
    assert mail.category_keyword("\\Deleted") == "_Deleted"
    assert len(mail.category_keyword("x" * 200)) == mail.MAX_KEYWORD_LEN
    flags = mail.imap_flags({"categories": [f"c{i}" for i in range(50)]})
    assert len(flags) == mail.MAX_KEYWORDS


@pytest.mark.parametrize("value,ok", [
    ("<abc@example.net>", True),
    ("<a b@x>", False),
    ("abc@example.net", False),
    ("<abc@x>\r\nINJECT", False),
    ("<" + "a" * 2000 + "@x>", False),
    (None, False),
])
def test_message_id_validation_silas_l3(value, ok):
    assert (mail.valid_message_id(value) is not None) is ok


def test_same_message_identical_regenerated_different_cato_f9():
    head = b"Message-ID: <m@x>\r\nDate: Tue, 1 Sep 2026 10:00:00 +0000\r\n"
    a = head + b"From: a@x\r\nSubject: hi\r\n\r\nbody\r\n"
    b = a.replace(b"\r\n", b"\n")
    # Exchange re-rendering: header order and extra headers change, payload does not
    extra = b"Subject: hi\r\nX-MS-Exchange-Organization-Foo: 1\r\nFrom: a@x\r\n"
    c = extra + head + b"\r\nbody\r\n"
    d = head + b"From: b@x\r\nSubject: hi\r\n\r\nbody\r\n"
    e = head + b"From: a@x\r\nSubject: hi\r\n\r\nbody but not the same\r\n"
    assert mail.same_message(a, b) == "identical"
    assert mail.same_message(a, c) == "regenerated"
    assert mail.same_message(a, d) == "different"
    assert mail.same_message(a, e) == "different"  # a changed leaf part is never "regenerated"


# -- delta expiry (Cato F2) ----------------------------------------------------------------

def _mail_env(tmp_path: Path):
    from test_mail import M1, M2, U, folder, mime  # reuse the module's fixtures

    cfg = make_config(tmp_path)
    state = State(tmp_path / "state.db")
    routes = {
        f"{U}/mailFolders/inbox": {"id": "f-inbox"},
        f"{U}/mailFolders": [folder("f-inbox", "Inbox", total=2)],
        f"{U}/mailFolders/f-inbox/messages": [M1, M2],
        f"{U}/mailFolders/f-inbox/messages/delta": ([{"id": "m1"}, {"id": "m2"}], "https://graph.microsoft.com/v1.0/d2"),
        f"{U}/messages/m1/$value": mime("<m1@x>", "one"),
        f"{U}/messages/m2/$value": mime(None, "two"),
    }
    for name in ("sentitems", "drafts", "deleteditems", "junkemail", "archive",
                 "conversationhistory", "outbox", "syncissues",
                 "recoverableitemsdeletions", "serverfailures", "localfailures"):
        routes[f"{U}/mailFolders/{name}"] = graph.GraphError(404, "no", name)
    return cfg, state, FakeGraph(routes)


def test_expired_delta_link_is_cleared_and_folder_listed_fully_cato_f2(tmp_path):
    from fakes_o365 import FakeImap, ImapWorld

    cfg, state, fake = _mail_env(tmp_path)
    stale = "https://graph.microsoft.com/v1.0/stale"
    state.set_delta("alice@contoso.com", "f-inbox", stale)
    fake.routes[stale] = graph.GraphError(410, "Gone: syncStateNotFound", stale)
    world = ImapWorld()
    mapping = MailboxMapping("alice@contoso.com", "alice@example.net")
    res = mail.MailMigrator(cfg, fake, state, lambda: FakeImap(world), mapping, False).migrate()
    inbox = next(f for f in res.folders if f.dest_name == "INBOX")
    assert inbox.delta_reset and inbox.appended == 2 and inbox.error is None
    assert state.get_delta("alice@contoso.com", "f-inbox") == "https://graph.microsoft.com/v1.0/d2"


def test_non_quota_append_refusal_fails_one_message_only_silas_m5(tmp_path):
    from fakes_o365 import FakeImap, ImapWorld

    cfg, state, fake = _mail_env(tmp_path)
    world = ImapWorld()

    class Picky(FakeImap):
        def append(self, folder, mime, flags, internal_date, message_id=None):
            if b"one" in mime:
                raise mail.ImapError("IMAP APPEND refused: NO [CANNOT] Invalid keyword")
            return super().append(folder, mime, flags, internal_date, message_id)

    mapping = MailboxMapping("alice@contoso.com", "alice@example.net")
    res = mail.MailMigrator(cfg, fake, state, lambda: Picky(world), mapping, False).migrate()
    inbox = next(f for f in res.folders if f.dest_name == "INBOX")
    assert inbox.failed == 1 and inbox.appended == 1 and not res.stopped


# -- persistent collection slugs (Cato F8) ---------------------------------------------

def test_calendar_slugs_are_stable_across_listing_order(tmp_path):
    cfg = make_config(tmp_path)
    state = State(tmp_path / "state.db")
    cal_a = {"id": "A", "name": "Team", "owner": {"address": "alice@contoso.com"}}
    cal_b = {"id": "B", "name": "Team", "owner": {"address": "alice@contoso.com"}}
    mapping = MailboxMapping("alice@contoso.com", "alice@example.net")
    route = "/users/alice@contoso.com/calendars"
    first = CalendarMigrator(cfg, FakeGraph({route: [cal_a, cal_b]}),
                             state, None, mapping, False)._calendars()[0]
    second = CalendarMigrator(cfg, FakeGraph({route: [cal_b, cal_a]}),
                              state, None, mapping, False)._calendars()[0]
    expected = {"A": "team", "B": "team-2"}
    assert {c["id"]: s for c, s in first} == expected
    assert {c["id"]: s for c, s in second} == expected


# -- config (Silas M1, M3, L8, NITs) ---------------------------------------------------

def _write(tmp_path: Path, extra: str) -> str:
    p = tmp_path / "c.toml"
    p.write_text(
        '[microsoft]\ntenant_id = "t"\nclient_id = "c"\n[mailcow]\nhost = "mail.example.net"\n'
        '[run]\nstate_dir = "' + str(tmp_path / "s") + '"\n' + extra, encoding="utf-8")
    os.chmod(p, 0o600)
    return str(p)


ENV = {"O365MIG_CLIENT_SECRET": "secret-value-1", "O365MIG_MAILCOW_API_KEY": "key-value-1"}


def test_duplicate_destination_rejected_silas_m3(tmp_path):
    path = _write(tmp_path, 'mailboxes = [{ source = "a@t", destination = "u@x" }, '
                            '{ source = "b@t", destination = "u@x" }]\n')
    with pytest.raises(ConfigError, match="more than once"):
        config.load_config(path, env=ENV)


def test_tls_toggle_is_gone_silas_m1(tmp_path):
    path = _write(tmp_path, 'mailboxes = ["a@example.net"]\nverify_tls = false\n')
    cfg = config.load_config(path, env=ENV)
    assert not hasattr(cfg, "verify_tls")


@pytest.mark.parametrize("host", ["a@b", "mail.example.net:993", "mail example.net", "-x.y"])
def test_bad_hostnames_rejected_silas_l8(tmp_path, host):
    p = tmp_path / "c.toml"
    p.write_text(f'[microsoft]\ntenant_id = "t"\nclient_id = "c"\n[mailcow]\nhost = "{host}"\n'
                 f'[run]\nmailboxes = ["a@example.net"]\n', encoding="utf-8")
    os.chmod(p, 0o600)
    with pytest.raises(ConfigError, match="bare hostname"):
        config.load_config(str(p), env=ENV)


def test_non_boolean_and_out_of_range_values_rejected(tmp_path):
    bad_bool = _write(tmp_path, 'mailboxes = ["a@example.net"]\ncontacts_photos = "false"\n')
    with pytest.raises(ConfigError, match="true or false"):
        config.load_config(bad_bool, env=ENV)
    too_many = _write(tmp_path, 'mailboxes = ["a@example.net"]\nparallel_mailboxes = 50\n')
    with pytest.raises(ConfigError, match="between 1 and 4"):
        config.load_config(too_many, env=ENV)


# -- secret redaction (Cato F13 / Silas L5) --------------------------------------------

def test_secret_filter_redacts_tracebacks_and_report_errors(tmp_path):
    flt = cli.SecretFilter()
    flt.add("hunter22-secret")
    stream = io.StringIO()
    handler = logging.StreamHandler(stream)
    handler.addFilter(flt)
    logger = logging.getLogger("test.redact")
    logger.addHandler(handler)
    logger.setLevel(logging.DEBUG)
    try:
        raise RuntimeError("token hunter22-secret leaked")
    except RuntimeError:
        logger.debug("boom", exc_info=True)
    logger.removeHandler(handler)
    assert "hunter22-secret" not in stream.getvalue() and "***" in stream.getvalue()

    report = RunReport("migrate", tmp_path, redactor=flt.redact)
    report.error("a@x", "failed with hunter22-secret\x1b[31m")
    written = json.loads(report.write(1).read_text())
    assert written["mailboxes"]["a@x"]["errors"] == ["failed with ***[31m"]


def test_clean_strips_controls_and_bidi():
    raw = "Inbox" + chr(0x1B) + "[2J" + chr(0x202E) + "evil" + chr(13) + chr(10) + "forged"
    assert clean(raw) == "Inbox[2Jevilforged"  # ESC, bidi override and CRLF removed


# -- state files (ISC-130) -------------------------------------------------------------

def test_state_side_files_are_private(tmp_path):
    s = State(tmp_path / "state.db")
    s.mark_message("a@x", "f", "g", "INBOX", None, STATUS_DONE)
    for suffix in ("", "-wal", "-shm"):
        assert oct((tmp_path / f"state.db{suffix}").stat().st_mode & 0o777) == "0o600"
    s.close()


def test_state_is_safe_under_threads(tmp_path):
    s = State(tmp_path / "state.db")

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


def test_internal_date_falls_back(tmp_path):
    fixed = datetime(2026, 1, 1, tzinfo=UTC)
    assert mail.internal_date({}, now=lambda: fixed) == fixed


# -- reconnect during APPEND (Cato F10) ------------------------------------------------

def test_append_after_reconnect_returns_existing_uid_instead_of_duplicating():
    from unittest import mock

    from o365_to_mailcow import imap_dest

    first = mock.MagicMock(name="client1")
    first.append.side_effect = OSError("connection reset")
    second = mock.MagicMock(name="client2")
    second.search.return_value = [42]
    # the copy committed before the drop is byte-identical to what we are appending
    second.fetch.return_value = {42: {b"BODY[]": b"Message-ID: <m@x>\r\n\r\nbody\r\n"}}
    clients = iter([first, second])
    dest = imap_dest.ImapDestination("mail.example.net", 993, "a@x", "pw",
                                     client_factory=lambda *a, **k: next(clients))
    uid = dest.append("INBOX", b"Message-ID: <m@x>\r\n\r\nbody\r\n", [], datetime.now(UTC),
                      message_id="<m@x>")
    assert uid == 42
    second.append.assert_not_called()  # the copy was already committed before the drop
    second.select_folder.assert_called_once_with("INBOX", readonly=True)


def test_append_after_reconnect_appends_when_existing_copy_differs():
    from unittest import mock

    from o365_to_mailcow import imap_dest

    first = mock.MagicMock(name="client1")
    first.append.side_effect = OSError("connection reset")
    second = mock.MagicMock(name="client2")
    second.search.return_value = [7]  # a same-Message-ID message that is NOT ours
    second.fetch.return_value = {7: {b"BODY[]": b"Message-ID: <m@x>\r\n\r\nforged\r\n"}}
    second.append.return_value = b"[APPENDUID 1 8] done"
    clients = iter([first, second])
    dest = imap_dest.ImapDestination("mail.example.net", 993, "a@x", "pw",
                                     client_factory=lambda *a, **k: next(clients))
    uid = dest.append("INBOX", b"Message-ID: <m@x>\r\n\r\nbody\r\n", [], datetime.now(UTC),
                      message_id="<m@x>")
    assert uid == 8 and second.append.call_count == 1


def test_same_message_survives_malformed_headers():
    bad = b'From: "\r\nMessage-ID: <x@[>\r\nDate: Tue, 1 Sep 99999999999 10:00:00 +0000\r\n\r\nx'
    good = b"From: a@x\r\nMessage-ID: <m@x>\r\n\r\ny"
    assert mail.same_message(bad, good) == "different"
    assert mail.same_message(good, bad) == "different"


@pytest.mark.parametrize("value", ["<a(b@x>", "<a@x)>", "<a{b}@x>", "<a%b@x>", "<a*@x>",
                                   "<x@[10.0.0.1]>", '<a"b@x>', "<a\\b@x>"])
def test_message_ids_with_imap_specials_are_rejected(value):
    assert mail.valid_message_id(value) is None


def test_duplicate_source_rejected(tmp_path):
    path = _write(tmp_path, 'mailboxes = [{ source = "a@t", destination = "u@x" }, '
                            '{ source = "a@t", destination = "v@x" }]\n')
    with pytest.raises(ConfigError, match="source mailbox a@t is listed more than once"):
        config.load_config(path, env=ENV)


def test_concurrent_runs_on_one_state_dir_are_refused(tmp_path):
    held = cli._acquire_lock(tmp_path)
    assert held is not None
    assert cli._acquire_lock(tmp_path) is None
    held.close()
    assert cli._acquire_lock(tmp_path) is not None
