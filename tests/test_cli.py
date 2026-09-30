"""End-to-end CLI flows with fake Graph/IMAP/DAV and a recorded mailcow transport."""

from __future__ import annotations

import json
import logging
import os
import re
import threading

import pytest
import responses

from fakes_o365 import FakeDav, FakeGraph, FakeImap, ImapWorld
from o365_to_mailcow import calendar_sync, cli, contacts_sync
from o365_to_mailcow.auth import AuthError
from o365_to_mailcow.mailcow import MailcowApi, MailcowError
from o365_to_mailcow.state import State

API = "https://mail.example.net/api/v1/"
CLIENT_SECRET = "client-secret-XYZ-987"
API_KEY = "mailcow-api-key-ABC-123"
CONFIG = """
[microsoft]
tenant_id = "tenant"
client_id = "client"
auth_mode = "{mode}"

[mailcow]
host = "mail.example.net"

[run]
state_dir = "{state}"
mailboxes = [
    {{ source = "alice@contoso.com", destination = "alice@example.net" }},
    {{ source = "bob@contoso.com", destination = "bob@example.net" }},
]
"""


def graph_routes(user: str) -> dict:
    u = f"/users/{user}"
    return {
        f"{u}/mailFolders/inbox": {"id": "f-inbox"},
        f"{u}/mailFolders": [{"id": "f-inbox", "displayName": "Inbox", "totalItemCount": 1,
                              "childFolderCount": 0,
                              "singleValueExtendedProperties": [
                                  {"id": "Long 0xe08", "value": "2048"}]}],
        f"{u}/mailFolders/f-inbox/messages": [
            {"id": "m1", "internetMessageId": "<m1@x>", "isRead": True,
             "receivedDateTime": "2024-01-01T00:00:00Z"}],
        f"{u}/mailFolders/f-inbox/messages/delta": ([], f"https://graph.microsoft.com/d/{user}"),
        f"https://graph.microsoft.com/d/{user}": ([], f"https://graph.microsoft.com/d/{user}"),
        f"{u}/messages/m1/$value": b"Message-ID: <m1@x>\r\n\r\nhi\r\n",
        f"{u}/messages/m1": {"internetMessageId": "<m1@x>"},
        f"{u}/calendars": [{"id": "c1", "name": "Calendar", "isDefaultCalendar": True,
                            "owner": {"address": user}}],
        f"{u}/calendars/c1/events": [{"id": "e1", "iCalUId": "UID1", "type": "singleInstance",
                                      "lastModifiedDateTime": "2024-01-01T00:00:00Z"}],
        f"{u}/contactFolders": [],
        f"{u}/contacts": [{"id": "k1", "lastModifiedDateTime": "2024-01-01T00:00:00Z"}],
    }


class FakeTokens:
    error: Exception | None = None

    def __init__(self, cfg, out=None) -> None:
        self.cfg = cfg

    def get_token(self) -> str:
        if FakeTokens.error:
            raise FakeTokens.error
        return "tok"


class World:
    """Everything the fakes share across one test, plus construction counters."""

    def __init__(self) -> None:
        self.imaps: dict[str, ImapWorld] = {}  # one fake IMAP server per mailbox
        self.quota_after: int | None = None
        self.davs: dict[str, FakeDav] = {}
        self.graphs: list[FakeGraph] = []
        self.imap_users: list[str] = []
        self.passwords: list[str] = []
        self.next_id = 100
        self.app_passwords: dict[str, list[dict]] = {}
        self.imap_connect_error: Exception | None = None


@pytest.fixture
def world(tmp_path, monkeypatch):
    w = World()
    FakeTokens.error = None
    routes = {**graph_routes("alice@contoso.com"), **graph_routes("bob@contoso.com")}

    def make_graph(tokens):
        g = FakeGraph(routes)
        w.graphs.append(g)
        return g

    class Imap(FakeImap):
        def __init__(self, host, port, user, password, verify=True):
            if user not in w.imaps:
                w.imaps[user] = ImapWorld(quota_after=w.quota_after)
            super().__init__(w.imaps[user])
            assert host == "mail.example.net" and verify is True
            w.imap_users.append(user)

        def connect(self):
            if w.imap_connect_error:
                raise w.imap_connect_error
            super().connect()

    def make_dav(host, user, password, verify=True):
        return w.davs.setdefault(user, FakeDav())

    def convert_event(master, instances, *, window, attendees="keep"):
        return calendar_sync.calendar_conv.ConvertedEvent(master["iCalUId"], b"ICS", None, ())

    def convert_contact(contact, photo=None):
        return contacts_sync.contacts_conv.ConvertedContact(f"uid-{contact['id']}", b"VCF",
                                                            None)

    monkeypatch.setattr(cli, "TokenProvider", FakeTokens)
    monkeypatch.setattr(cli, "GraphClient", make_graph)
    monkeypatch.setattr(cli, "ImapDestination", Imap)
    monkeypatch.setattr(cli, "SogoDav", make_dav)
    monkeypatch.setattr(calendar_sync.calendar_conv, "convert_event", convert_event)
    monkeypatch.setattr(contacts_sync.contacts_conv, "convert_contact", convert_contact)
    monkeypatch.setenv("O365MIG_CLIENT_SECRET", CLIENT_SECRET)
    monkeypatch.setenv("O365MIG_MAILCOW_API_KEY", API_KEY)
    return w


@pytest.fixture
def config(tmp_path):
    def write(mode: str = "app") -> str:
        p = tmp_path / "config.toml"
        p.write_text(CONFIG.format(mode=mode, state=tmp_path / "state"), encoding="utf-8")
        os.chmod(p, 0o600)
        return str(p)
    return write


@pytest.fixture
def mailcow(world):
    """Recorded mailcow API: every call lands in responses.calls."""
    lock = threading.Lock()  # mailbox threads hit the fake concurrently
    with responses.RequestsMock(assert_all_requests_are_fired=False) as rsps:
        def get_mailbox(req):
            addr = req.url.rsplit("/", 1)[1]
            known = addr in ("alice@example.net", "bob@example.net")
            return 200, {}, json.dumps({"username": addr} if known else {})

        def list_pw(req):
            addr = req.url.rsplit("/", 1)[1]
            with lock:
                return 200, {}, json.dumps(world.app_passwords.get(addr, []))

        def add_pw(req):
            body = json.loads(req.body)
            with lock:
                world.passwords.append(body["app_passwd"])
                world.next_id += 1
                world.app_passwords.setdefault(body["username"], []).append(
                    {"id": world.next_id, "name": body["app_name"]})
            return 200, {}, json.dumps([{"type": "success", "msg": "app_passwd_added",
                                         "log": ["app_passwd", "add", body]}])

        def delete_pw(req):
            body = json.loads(req.body)
            assert isinstance(body, list), "mailcow delete takes a bare JSON array of ids"
            ids = {str(i) for i in body}
            with lock:
                for addr, items in list(world.app_passwords.items()):
                    world.app_passwords[addr] = [p for p in items if str(p["id"]) not in ids]
            return 200, {}, json.dumps([{"type": "success", "msg": "deleted"}])

        rsps.add_callback("GET", re.compile(API + "get/mailbox/.*"), callback=get_mailbox)
        rsps.add_callback("GET", re.compile(API + "get/app-passwd/all/.*"), callback=list_pw)
        rsps.add_callback("POST", API + "add/app-passwd", callback=add_pw)
        rsps.add_callback("POST", API + "delete/app-passwd", callback=delete_pw)
        yield rsps


def mailcow_calls(rsps, method=None):
    return [(c.request.method, c.request.url.removeprefix(API)) for c in rsps.calls
            if method is None or c.request.method == method]


def dav_writes(world):
    return [w for d in world.davs.values() for w in d.writes]


# -- help, config errors ---------------------------------------------------------------

@pytest.mark.parametrize("argv", [["--help"], ["o365mig", "--help"]])
def test_help_lists_commands_isc_6(argv, capsys):
    with pytest.raises(SystemExit) as exc:
        cli.main(argv)
    assert exc.value.code == 0
    out = capsys.readouterr().out
    for cmd in ("plan", "migrate", "verify", "cleanup"):
        assert cmd in out


def test_missing_config_keys_exit_2_isc_18(tmp_path, capsys, monkeypatch):
    monkeypatch.delenv("O365MIG_CLIENT_SECRET", raising=False)
    monkeypatch.delenv("O365MIG_MAILCOW_API_KEY", raising=False)
    p = tmp_path / "c.toml"
    p.write_text("[microsoft]\n", encoding="utf-8")
    os.chmod(p, 0o600)
    assert cli.main(["--config", str(p), "plan"]) == 2
    err = capsys.readouterr().err
    assert "microsoft.tenant_id" in err and "mailcow.host" in err and "run.mailboxes" in err


def test_unknown_mailbox_filter_exit_2(world, config):
    assert cli.main(["--config", config(), "--mailbox", "nobody@x.com", "plan"]) == 2


# -- migrate ---------------------------------------------------------------------------

def test_migrate_success_full_lifecycle(world, config, mailcow, tmp_path, capsys):
    code = cli.main(["--config", config(), "migrate"])
    assert code == 0, capsys.readouterr()
    # both mailboxes: mail appended, event and contact PUT
    assert [len(i.folders["INBOX"]) for i in world.imaps.values()] == [1, 1]
    assert ("PUT", "Calendar/personal/UID1.ics") in world.davs["alice@example.net"].writes
    assert ("PUT", "Contacts/personal/uid-k1.vcf") in world.davs["bob@example.net"].writes
    # ISC-92/113: one app password per mailbox, created and deleted again
    assert len(mailcow_calls(mailcow, "POST")) == 4
    assert [u for m, u in mailcow_calls(mailcow, "POST")].count("add/app-passwd") == 2
    assert all(not v for v in world.app_passwords.values())
    state = State(tmp_path / "state" / "state.db")
    assert state.app_passwords() == []
    state.close()
    # one GraphClient per mailbox (4-in-flight cap is per mailbox)
    assert len(world.graphs) == 2
    # ISC-95: only the four allowed endpoints
    for _, url in mailcow_calls(mailcow):
        assert url.startswith(("get/mailbox/", "get/app-passwd/", "add/app-passwd",
                               "delete/app-passwd"))
    # ISC-114: report and log files
    reports = list((tmp_path / "state" / "reports").glob("*.json"))
    data = json.loads(reports[0].read_text())
    assert data["exit_code"] == 0
    assert data["mailboxes"]["alice@contoso.com"]["mail"]["folders"][0]["appended"] == 1
    assert data["mailboxes"]["bob@contoso.com"]["calendar"]["collections"][0]["put"] == 1
    assert list((tmp_path / "state" / "logs").glob("*.log"))


def test_second_migrate_is_idempotent(world, config, mailcow):
    assert cli.main(["--config", config(), "migrate"]) == 0
    appends = sum(i.appends for i in world.imaps.values())
    writes = len(dav_writes(world))
    assert appends == 2
    assert cli.main(["--config", config(), "migrate"]) == 0
    assert sum(i.appends for i in world.imaps.values()) == appends  # ISC-50
    assert len(dav_writes(world)) == writes  # ISC-81


def test_dry_run_zero_writes_isc_96(world, config, mailcow, tmp_path, capsys):
    assert cli.main(["--config", config(), "--dry-run", "migrate"]) == 0
    assert mailcow_calls(mailcow, "POST") == []
    assert world.imap_users == [] and world.imaps == {}
    assert world.davs == {}
    out = capsys.readouterr().out
    assert "would append 1" in out and "would put 1" in out
    state = State(tmp_path / "state" / "state.db")
    assert state.message_counts("alice@contoso.com") == {}
    state.close()


def test_plan_creates_nothing_isc_112_117(world, config, mailcow, capsys):
    assert cli.main(["--config", config(), "plan"]) == 0
    assert mailcow_calls(mailcow, "POST") == []
    assert {u.split("/")[0] + "/" + u.split("/")[1] for _, u in mailcow_calls(mailcow)} == {
        "get/mailbox"}
    assert world.imap_users == [] and world.davs == {}
    out = capsys.readouterr().out
    assert "mail: 1 messages, 2.0 KiB" in out
    assert "INBOX" in out and "calendar: 1 items" in out


def test_plan_reports_missing_destination_isc_94(world, config, mailcow, tmp_path, capsys):
    p = tmp_path / "config.toml"
    text = CONFIG.format(mode="app", state=tmp_path / "state").replace(
        "bob@example.net", "ghost@example.net")
    p.write_text(text, encoding="utf-8")
    os.chmod(p, 0o600)
    assert cli.main(["--config", str(p), "plan"]) == 1
    assert "MISSING IN MAILCOW" in capsys.readouterr().out
    assert cli.main(["--config", str(p), "migrate"]) == 1
    assert "ghost@example.net" not in world.app_passwords  # never touched
    assert [u for _, u in mailcow_calls(mailcow, "POST")].count("add/app-passwd") == 1


def test_app_password_deleted_on_exception_isc_113(world, config, mailcow, tmp_path):
    world.imap_connect_error = RuntimeError("IMAP exploded")
    assert cli.main(["--config", config(), "--only", "mail", "migrate"]) == 1
    posts = [u for _, u in mailcow_calls(mailcow, "POST")]
    assert posts.count("add/app-passwd") == 2 and posts.count("delete/app-passwd") == 2
    assert all(not v for v in world.app_passwords.values())


def test_keep_app_passwords(world, config, mailcow, tmp_path):
    assert cli.main(["--config", config(), "--mailbox", "alice@contoso.com",
                     "migrate", "--keep-app-passwords"]) == 0
    assert "delete/app-passwd" not in [u for _, u in mailcow_calls(mailcow, "POST")]
    state = State(tmp_path / "state" / "state.db")
    assert state.app_passwords() == [("alice@example.net", "101")]
    state.close()


def test_leftover_app_passwords_deleted_first_isc_125(world, config, mailcow, tmp_path):
    (tmp_path / "state").mkdir()
    state = State(tmp_path / "state" / "state.db")
    state.record_app_password("alice@example.net", "7")
    state.record_app_password("alice@example.net", "8")  # no longer in mailcow
    state.close()
    world.app_passwords["alice@example.net"] = [{"id": 7, "name": "o365-migration"}]
    assert cli.main(["--config", config(), "--mailbox", "alice@contoso.com", "--only",
                     "mail", "migrate"]) == 0
    deletes = [json.loads(c.request.body) for c in mailcow.calls
               if c.request.url.endswith("delete/app-passwd")]
    assert deletes[0] == ["7"]
    state = State(tmp_path / "state" / "state.db")
    assert state.app_passwords() == []
    state.close()


def test_only_filter_isc_116(world, config, mailcow):
    assert cli.main(["--config", config(), "--only", "contacts", "migrate"]) == 0
    assert world.imap_users == []
    paths = [p for g in world.graphs for p in g.paths()]
    assert paths and all("contact" in p for p in paths)


def test_item_failure_exit_1_isc_115(world, config, mailcow):
    world.quota_after = 0  # every APPEND answers NO [OVERQUOTA]
    assert cli.main(["--config", config(), "--only", "mail", "migrate"]) == 1


def test_secrets_never_logged_isc_21(world, config, mailcow, tmp_path, caplog, capsys):
    with caplog.at_level(logging.DEBUG):
        cli.main(["--config", config(), "-v", "migrate"])
        cli.main(["--config", config(), "-v", "verify"])
    logs = "".join(p.read_text() for p in (tmp_path / "state" / "logs").glob("*.log"))
    out = capsys.readouterr()
    reports = "".join(p.read_text() for p in (tmp_path / "state" / "reports").glob("*"))
    assert world.passwords
    for secret in (CLIENT_SECRET, API_KEY, *world.passwords):
        for text in (logs, caplog.text, out.out, out.err, reports):
            assert secret not in text


def test_secret_filter_redacts():
    f = cli.SecretFilter()
    f.add("hunter2-long")
    rec = logging.LogRecord("x", logging.INFO, "f", 1, "pw=%s", ("hunter2-long",), None)
    f.filter(rec)
    assert rec.getMessage() == "pw=***"


# -- verify ----------------------------------------------------------------------------

def test_verify_all_clear_exit_0(world, config, mailcow, capsys):
    assert cli.main(["--config", config(), "migrate"]) == 0
    capsys.readouterr()
    assert cli.main(["--config", config(), "verify", "--sample", "1"]) == 0
    out = capsys.readouterr().out
    assert "All counts match" in out
    assert "sample: 1 compared, 0 mismatched" in out
    posts = [u for _, u in mailcow_calls(mailcow, "POST")]
    assert posts.count("add/app-passwd") == posts.count("delete/app-passwd")


def test_verify_mismatch_exit_1_isc_128(world, config, mailcow, capsys):
    assert cli.main(["--config", config(), "migrate"]) == 0
    world.davs["alice@example.net"].calendars["personal"].clear()
    capsys.readouterr()
    assert cli.main(["--config", config(), "verify"]) == 1
    out = capsys.readouterr().out
    assert "VERIFY FAILED" in out and "All counts match" not in out
    assert "expected 1, DAV has 0" in out


def test_verify_rejects_dry_run(world, config, mailcow):
    assert cli.main(["--config", config(), "--dry-run", "verify"]) == 2
    assert mailcow_calls(mailcow, "POST") == []


# -- cleanup ---------------------------------------------------------------------------

def test_cleanup_deletes_named_and_recorded_isc_93(world, config, mailcow, tmp_path, capsys):
    (tmp_path / "state").mkdir()
    state = State(tmp_path / "state" / "state.db")
    state.record_app_password("old@example.net", "55")  # not in config any more
    state.close()
    world.app_passwords = {
        "alice@example.net": [{"id": 1, "name": "o365-migration"},
                              {"id": 2, "name": "thunderbird"}],
        "bob@example.net": [{"id": 3, "name": "o365-migration"}],
        "old@example.net": [{"id": 55, "name": "o365-migration"}],
    }
    assert cli.main(["--config", config(), "cleanup"]) == 0
    assert world.app_passwords["alice@example.net"] == [{"id": 2, "name": "thunderbird"}]
    assert world.app_passwords["bob@example.net"] == []
    assert world.app_passwords["old@example.net"] == []
    assert "deleted 3 app password(s)" in capsys.readouterr().out


def test_cleanup_dry_run_deletes_nothing(world, config, mailcow, capsys):
    world.app_passwords = {"alice@example.net": [{"id": 1, "name": "o365-migration"}]}
    assert cli.main(["--config", config(), "--dry-run", "cleanup"]) == 0
    assert mailcow_calls(mailcow, "POST") == []
    assert "would delete app password id=1" in capsys.readouterr().out


# -- auth ------------------------------------------------------------------------------

def test_device_code_refused_explains_policy_isc_126(world, config, mailcow, capsys):
    FakeTokens.error = AuthError("authorization_declined: AADSTS50199 device code flow "
                                 "blocked by policy")
    assert cli.main(["--config", config("delegated"), "plan"]) == 2
    err = capsys.readouterr().err
    assert "authentication flows" in err and 'auth_mode = "app"' in err


def test_app_mode_auth_error_exit_2_plain_message(world, config, mailcow, capsys):
    FakeTokens.error = AuthError("invalid_client: AADSTS7000215 Invalid client secret")
    assert cli.main(["--config", config(), "migrate"]) == 2
    err = capsys.readouterr().err
    assert "Microsoft sign-in failed" in err and "authentication flows" not in err
    assert mailcow_calls(mailcow, "POST") == []


# -- provision ---------------------------------------------------------------------------

def _provision_routes(mailcow, world, known_domains=("example.net",)):
    def get_domain(req):
        domain = req.url.rsplit("/", 1)[1]
        return 200, {}, json.dumps({"domain_name": domain} if domain in known_domains else {})

    def add_mailbox(req):
        body = json.loads(req.body)
        world.created_mailboxes.append(body)
        return 200, {}, json.dumps([{"type": "success", "msg": ["mailbox_added"], "log": []}])

    mailcow.add_callback("GET", re.compile(API + "get/domain/.*"), callback=get_domain)
    mailcow.add_callback("POST", API + "add/mailbox", callback=add_mailbox)


def test_provision_creates_missing_mailboxes_and_writes_password_file(world, config, mailcow,
                                                                      tmp_path, capsys):
    world.created_mailboxes = []
    _provision_routes(mailcow, world)
    # bob@example.net is "known" to the fake get/mailbox; ghost is not
    csv = tmp_path / "boxes.csv"
    csv.write_text("ghost@contoso.com,ghost@example.net,Ghost Rider,2048\n", encoding="utf-8")
    assert cli.main(["--config", config(), "--mailboxes", str(csv), "provision"]) == 0
    out = capsys.readouterr().out
    assert "alice@example.net: exists" in out and "ghost@example.net: created" in out
    assert len(world.created_mailboxes) == 1
    body = world.created_mailboxes[0]
    assert body["local_part"] == "ghost" and body["domain"] == "example.net"
    assert body["name"] == "Ghost Rider" and body["quota"] == "2048"
    assert body["force_pw_update"] == "1" and body["password"] == body["password2"]
    files = list((tmp_path / "state").glob("provisioned-*.csv"))
    assert len(files) == 1 and oct(files[0].stat().st_mode & 0o777) == "0o600"
    assert f"ghost@example.net,{body['password']}" in files[0].read_text()
    # the generated password never reaches stdout or the log
    assert body["password"] not in out
    logs = "".join(p.read_text() for p in (tmp_path / "state" / "logs").glob("*.log"))
    assert body["password"] not in logs


def test_provision_dry_run_creates_nothing_and_refuses_unknown_domain(world, config, mailcow,
                                                                     tmp_path, capsys):
    world.created_mailboxes = []
    _provision_routes(mailcow, world)
    csv = tmp_path / "boxes.csv"
    csv.write_text("ghost@contoso.com,ghost@example.net\nnew@contoso.com,new@other.tld\n",
                   encoding="utf-8")
    code = cli.main(["--config", config(), "--mailboxes", str(csv), "--dry-run", "provision"])
    assert code == 1  # other.tld does not exist in mailcow
    out = capsys.readouterr().out
    assert "ghost@example.net: would create" in out and "other.tld does not exist" in out
    assert world.created_mailboxes == [] and mailcow_calls(mailcow, "POST") == []


def test_provisioning_endpoints_are_locked_for_every_other_command():
    api = MailcowApi("mail.example.net", "key")
    with pytest.raises(MailcowError, match="non-allowlisted"):
        api.create_mailbox("a@example.net", "A", 1024, "pw")
    with pytest.raises(MailcowError, match="non-allowlisted"):
        api.domain_exists("example.net")
