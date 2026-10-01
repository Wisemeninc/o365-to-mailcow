"""Local web UI server (ISC-155..161, 163, 166, 167): contract tests over real HTTP.

The server runs on an ephemeral port in a thread; requests go through ``http.client``.
Graph and MSAL are faked, mailcow is recorded with ``responses``.
"""

from __future__ import annotations

import hashlib
import http.client
import io
import json
import os
import re
import stat
import threading
import time
from dataclasses import dataclass
from pathlib import Path

import pytest
import responses

from fakes_o365 import FakeGraph, make_config
from o365_to_mailcow import __version__, cli, web
from o365_to_mailcow.auth import AuthError
from o365_to_mailcow.config import MailboxMapping, load_config
from o365_to_mailcow.graph import GraphError
from o365_to_mailcow.mail import FolderVerify, MailVerify
from o365_to_mailcow.report import Progress, RunReport, _plain
from o365_to_mailcow.summary import summarize_report

API = "https://mail.example.net/api/v1/"
TOKEN = "web-token-" + "Q" * 33
CLIENT_SECRET = "client-secret-web-777"
API_KEY = "mailcow-api-key-web-888"
SECRETS = (TOKEN, CLIENT_SECRET, API_KEY)
CONFIG = """
[microsoft]
tenant_id = "tenant"
client_id = "client"

[mailcow]
host = "mail.example.net"

[run]
state_dir = "{state}"
"""
PAGE = ('<!doctype html><html><head><style nonce="__CSP_NONCE__">b{}</style></head>'
        '<body><script nonce="__CSP_NONCE__">1</script></body></html>')
SECURITY_HEADERS = {
    "x-content-type-options": "nosniff",
    "referrer-policy": "no-referrer",
    "cache-control": "no-store",
    "x-frame-options": "DENY",
}
USERS = [
    {"id": "u1", "displayName": "Zed Licensed", "mail": "Zed@Contoso.com",
     "userPrincipalName": "zed@contoso.onmicrosoft.com", "userType": "Member",
     "accountEnabled": True, "assignedLicenses": [{"skuId": "sku"}],
     "proxyAddresses": ["SMTP:zed@contoso.com", "smtp:Z.Alias@Contoso.com",
                        "SIP:zed@contoso.com", "X500:/o=ExchangeLabs/cn=zed",
                        "smtp:zed@contoso.onmicrosoft.com"]},
    {"id": "u2", "displayName": "Info Shared", "mail": "info@contoso.com",
     "userPrincipalName": "info@contoso.com", "userType": "Member", "accountEnabled": False,
     "assignedLicenses": [], "proxyAddresses": ["SMTP:info@contoso.com",
                                                "smtp:info@contoso.com"]},
    {"id": "u3", "displayName": "Guest Person", "mail": "guest@partner.example",
     "userPrincipalName": "guest_partner.example#EXT#@contoso.com", "userType": "Guest",
     "accountEnabled": True, "assignedLicenses": [], "proxyAddresses": []},
    {"id": "u4", "displayName": "No Mailbox", "mail": None, "userPrincipalName": "nomail@x",
     "userType": "Member", "accountEnabled": True, "assignedLicenses": [{"skuId": "sku"}]},
    {"id": "u5", "displayName": "Amy\u202e Evil", "mail": "amy@contoso.com",
     "userPrincipalName": "amy@contoso.com", "userType": None, "accountEnabled": True,
     "assignedLicenses": [{"skuId": "sku"}], "proxyAddresses": None},
]
GOOD = {"source": "a@contoso.com", "destination": "a@example.net", "name": "A",
        "quota_mib": None}


# -- harness ---------------------------------------------------------------------------

@dataclass
class Resp:
    status: int
    headers: dict[str, str]
    body: bytes

    def json(self):
        return json.loads(self.body)


class Client:
    def __init__(self, port: int) -> None:
        self.port = port

    def request(self, method: str, path: str, body: object = None, *, token: str | None = TOKEN,
                headers: dict | None = None, raw: bytes | None = None) -> Resp:
        conn = http.client.HTTPConnection("127.0.0.1", self.port, timeout=15)
        hdrs = dict(headers or {})
        if token is not None:
            hdrs["Authorization"] = f"Bearer {token}"
        data = raw if raw is not None else (None if body is None else json.dumps(body).encode())
        if data is not None:
            hdrs.setdefault("Content-Type", "application/json")
        try:
            conn.request(method, path, body=data, headers=hdrs)
            resp = conn.getresponse()
            return Resp(resp.status, {k.lower(): v for k, v in resp.getheaders()}, resp.read())
        finally:
            conn.close()

    def get(self, path: str, **kw) -> Resp:
        return self.request("GET", path, **kw)

    def post(self, path: str, body: object, **kw) -> Resp:
        return self.request("POST", path, body, **kw)

    def put(self, path: str, body: object, **kw) -> Resp:
        return self.request("PUT", path, body, **kw)


class FakeTokens:
    def __init__(self, cfg, cache_path=None, out=None) -> None:
        self.cfg = cfg

    def get_token(self) -> str:
        return "tok"


@pytest.fixture
def conf(tmp_path, monkeypatch):
    """Config file without any mailbox list, plus the secrets in the environment."""
    monkeypatch.setenv("O365MIG_CLIENT_SECRET", CLIENT_SECRET)
    monkeypatch.setenv("O365MIG_MAILCOW_API_KEY", API_KEY)
    monkeypatch.delenv("O365MIG_WEB_TOKEN", raising=False)
    monkeypatch.delenv("O365MIG_CONFIG", raising=False)
    p = tmp_path / "config.toml"
    p.write_text(CONFIG.format(state=tmp_path / "state"), encoding="utf-8")
    os.chmod(p, 0o600)
    return p


@pytest.fixture
def graph(monkeypatch):
    g = FakeGraph({"/users": list(USERS)})
    monkeypatch.setattr(web, "TokenProvider", FakeTokens)
    monkeypatch.setattr(web, "GraphClient", lambda tokens: g)
    return g


@pytest.fixture
def client(tmp_path, conf, graph):
    page = tmp_path / "index.html"
    page.write_text(PAGE, encoding="utf-8")
    cfg = make_config(tmp_path, client_secret=CLIENT_SECRET, mailcow_api_key=API_KEY,
                      mailboxes=())
    srv = web.make_server(cfg, "127.0.0.1", 0, TOKEN, page, config_path=str(conf),
                          out=io.StringIO())
    thread = threading.Thread(target=srv.serve_forever, kwargs={"poll_interval": 0.05},
                              daemon=True)
    thread.start()
    c = Client(srv.server_address[1])
    c.state_dir = tmp_path / "state"
    c.server = srv
    c.selection = tmp_path / "state" / "mailboxes.csv"
    c.conf = conf
    yield c
    srv.shutdown()
    srv.server_close()
    thread.join(5)


def wait_job(client: Client, job_id: str, timeout: float = 15.0) -> dict:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        job = client.get(f"/api/jobs/{job_id}").json()
        if not job["running"]:
            return job
        time.sleep(0.02)
    raise AssertionError(f"job {job_id} did not finish")


def assert_security_headers(resp: Resp) -> None:
    for name, value in SECURITY_HEADERS.items():
        assert resp.headers.get(name) == value, (name, resp.headers)
    assert "default-src 'none'" in resp.headers.get("content-security-policy", "")
    assert "access-control-allow-origin" not in resp.headers


# -- auth, page, headers, routing ------------------------------------------------------

@pytest.mark.parametrize("header", [None, "Bearer wrong", f"Basic {TOKEN}", f"Bearer {TOKEN}x",
                                    "Bearer", f"Bearer {TOKEN[:-1]}", "Bearer \u00e9\u00e9"])
def test_api_requires_the_bearer_token_isc_156(client, header, monkeypatch):
    started = []
    monkeypatch.setattr(cli, "main", lambda *a, **kw: started.append(a) or 0)
    headers = {"Authorization": header} if header else {}
    for method, path, body in (("GET", "/api/status", None), ("GET", "/api/selection", None),
                               ("POST", "/api/jobs", {"command": "plan"}),
                               ("GET", "/api/overview", None),
                               ("GET", "/api/nonexistent", None)):
        r = client.request(method, path, body, token=None, headers=headers)
        assert r.status == 401, (method, path)
        assert r.json() == {"error": "unauthorized"}
        assert_security_headers(r)
    assert started == []
    assert client.get("/api/status").status == 200


def test_page_is_served_with_a_fresh_csp_nonce_isc_163(client):
    first = client.get("/", token=None)
    second = client.get("/", token=None)
    assert first.status == 200
    assert first.headers["content-type"] == "text/html; charset=utf-8"
    body = first.body.decode()
    assert "__CSP_NONCE__" not in body
    nonce = re.search(r'nonce="([^"]+)"', body).group(1)
    assert body.count(f'nonce="{nonce}"') == 2
    csp = first.headers["content-security-policy"]
    for part in ("default-src 'none'", f"script-src 'nonce-{nonce}'",
                 f"style-src 'nonce-{nonce}'", "connect-src 'self'", "img-src 'self' data:",
                 "base-uri 'none'", "form-action 'none'"):
        assert part in csp
    assert nonce not in second.body.decode()  # per response, never reused
    assert_security_headers(first)


@pytest.mark.parametrize("method, path, token, raw", [
    ("GET", "/", None, None),
    ("GET", "/api/status", TOKEN, None),
    ("GET", "/api/status", None, None),
    ("GET", "/nope", None, None),
    ("DELETE", "/api/selection", TOKEN, None),
    ("PUT", "/api/selection", TOKEN, b"{not json"),
    ("GET", "/api/reports/latest", TOKEN, None),
    ("GET", "/api/overview", TOKEN, None),
    ("GET", "/api/overview", None, None),
    ("POST", "/api/overview", TOKEN, None),
])
def test_security_headers_on_every_response_isc_163(client, method, path, token, raw):
    assert_security_headers(client.request(method, path, token=token, raw=raw))


@pytest.mark.parametrize("path", ["/api/nope", "/api/status/", "/api/jobs/x/y", "/index.html",
                                  "/../etc/passwd", "/web_static/index.html", "/favicon.ico",
                                  "/api/jobs/" + "a" * 65, "/api/jobs/unknown",
                                  "/api/jobs/unknown/output", "/api//status", "/api/status%00"])
def test_unknown_paths_are_404_isc_166(client, path):
    r = client.get(path)
    assert r.status == 404
    assert r.json() == {"error": "not found"}


@pytest.mark.parametrize("method", ["DELETE", "PATCH", "OPTIONS", "TRACE", "FOO"])
def test_other_methods_are_405(client, method):
    r = client.request(method, "/api/selection")
    assert r.status == 405
    assert r.json() == {"error": "method not allowed"}
    assert r.headers["allow"] == "GET, POST, PUT"
    assert_security_headers(r)


def test_known_path_with_wrong_method_is_405(client):
    for method, path in (("PUT", "/api/status"), ("POST", "/"), ("GET", "/api/mailcow/check"),
                         ("PUT", "/api/jobs"), ("POST", "/api/overview"),
                         ("PUT", "/api/overview")):
        r = client.request(method, path, {})
        assert r.status == 405, (method, path)


def test_body_over_1_mib_is_413_isc_163(client):
    r = client.request("PUT", "/api/selection", raw=b"x" * (2 * 1024 * 1024))
    assert r.status == 413
    assert "too large" in r.json()["error"]
    assert_security_headers(r)
    assert not client.selection.exists()
    # exactly 1 MiB passes the size gate (and then fails JSON parsing)
    assert client.request("PUT", "/api/selection", raw=b" " * (1024 * 1024)).status == 400


@pytest.mark.parametrize("raw", [b"{not json", b"\xff\xfe", b"[]", b""])
def test_bad_json_bodies_are_400(client, raw):
    r = client.request("PUT", "/api/selection", raw=raw)
    assert r.status == 400
    assert "error" in r.json()


# -- status ----------------------------------------------------------------------------

def test_status_has_settings_and_no_secrets_isc_166(client, tmp_path):
    r = client.get("/api/status")
    assert r.status == 200
    assert r.json() == {
        "version": __version__, "auth_mode": "app", "mailcow_host": "mail.example.net",
        "state_dir": str(tmp_path / "state"),
        "selection_path": str(tmp_path / "state" / "mailboxes.csv"), "running_job": None,
        "configured": True,
    }
    for secret in SECRETS:
        assert secret not in r.body.decode()


# -- settings ----------------------------------------------------------------------------

def test_settings_roundtrip_never_echoes_secrets(client, tmp_path):
    r = client.get("/api/settings")
    assert r.status == 200
    data = r.json()
    assert data["microsoft"]["client_secret_set"] is True and data["mailcow"]["api_key_set"]
    for secret in SECRETS:
        assert secret not in r.body.decode()
    body = {"microsoft": {"tenant_id": "tenant-2", "client_id": "client-2", "auth_mode": "app",
                          "client_secret": "new-secret-value-xyz"},
            "mailcow": {"host": "Mail2.Example.net", "api_key": "new-api-key-value-xyz"}}
    r = client.put("/api/settings", body)
    assert r.status == 200, r.body
    saved = r.json()
    assert saved["microsoft"]["tenant_id"] == "tenant-2"
    assert saved["mailcow"]["host"] == "mail2.example.net" and saved["configured"] is True
    assert "new-secret-value-xyz" not in r.body.decode()
    path = tmp_path / "state" / "settings.toml"
    assert oct(path.stat().st_mode & 0o777) == "0o600"
    text = path.read_text()
    assert 'client_secret = "new-secret-value-xyz"' in text
    assert 'api_key = "new-api-key-value-xyz"' in text
    # the effective config now uses the saved values, and status reflects the new host
    assert client.get("/api/status").json()["mailcow_host"] == "mail2.example.net"
    # saving again without secrets keeps the stored ones
    body["microsoft"].pop("client_secret")
    body["mailcow"].pop("api_key")
    assert client.put("/api/settings", body).status == 200
    assert "new-secret-value-xyz" in path.read_text()


def test_settings_validation(client):
    bad_host = {"microsoft": {"tenant_id": "t", "client_id": "c", "auth_mode": "app"},
                "mailcow": {"host": "mail.example.net:993"}}
    assert client.put("/api/settings", bad_host).status == 400
    bad_mode = {"microsoft": {"tenant_id": "t", "client_id": "c", "auth_mode": "magic"},
                "mailcow": {"host": "mail.example.net"}}
    assert client.put("/api/settings", bad_mode).status == 400
    assert client.put("/api/settings", []).status == 400


# -- tenant listing --------------------------------------------------------------------

def test_tenant_mailboxes_mapping_isc_157(client, graph):
    r = client.get("/api/tenant/mailboxes")
    assert r.status == 200, r.body
    data = r.json()
    assert re.match(r"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d", data["fetched_at"])
    assert data["mailboxes"] == [
        {"id": "u5", "display_name": "Amy Evil", "mail": "amy@contoso.com",
         "upn": "amy@contoso.com", "kind": "user", "enabled": True, "aliases": []},
        {"id": "u2", "display_name": "Info Shared", "mail": "info@contoso.com",
         "upn": "info@contoso.com", "kind": "shared", "enabled": False, "aliases": []},
        {"id": "u1", "display_name": "Zed Licensed", "mail": "zed@contoso.com",
         "upn": "zed@contoso.onmicrosoft.com", "kind": "user", "enabled": True,
         "aliases": ["z.alias@contoso.com", "zed@contoso.onmicrosoft.com"]},
    ]
    method, path, params, _ = graph.calls[0]
    assert (method, path) == ("iter_pages", "/users")
    assert params["$top"] == "999"
    assert set(params["$select"].split(",")) == {
        "id", "displayName", "mail", "userPrincipalName", "userType", "accountEnabled",
        "assignedLicenses", "proxyAddresses"}


def test_tenant_listing_is_cached_until_refresh_or_expiry_isc_157(client, graph, monkeypatch):
    first = client.get("/api/tenant/mailboxes").json()
    assert client.get("/api/tenant/mailboxes").json() == first
    assert len(graph.calls) == 1
    graph.routes["/users"] = USERS[:1]
    refreshed = client.get("/api/tenant/mailboxes?refresh=1").json()
    assert len(graph.calls) == 2
    assert [m["id"] for m in refreshed["mailboxes"]] == ["u1"]
    monkeypatch.setattr(web, "TENANT_CACHE_SECONDS", 0.0)
    client.get("/api/tenant/mailboxes")
    assert len(graph.calls) == 3


def test_tenant_listing_403_names_the_missing_permission(client, graph):
    graph.routes["/users"] = GraphError(403, "Authorization_RequestDenied: Insufficient "
                                             "privileges to complete the operation.", "/users")
    r = client.get("/api/tenant/mailboxes")
    assert r.status == 502
    assert r.json()["error"].startswith("GraphError: HTTP 403")
    assert "User.Read.All" in r.json()["error"]


@pytest.mark.parametrize("exc, prefix", [
    (GraphError(503, "ServiceUnavailable: try later", "/users"), "GraphError: HTTP 503"),
    (AuthError(f"invalid_client: bad secret {CLIENT_SECRET}"), "AuthError: invalid_client"),
])
def test_tenant_listing_upstream_errors_are_502_and_redacted(client, graph, exc, prefix):
    graph.routes["/users"] = exc
    r = client.get("/api/tenant/mailboxes")
    assert r.status == 502
    assert r.json()["error"].startswith(prefix)
    assert CLIENT_SECRET not in r.body.decode()


# -- mailcow check ---------------------------------------------------------------------

def test_mailcow_check_reports_existence_isc_158(client):
    def get_mailbox(req):
        addr = req.url.rsplit("/", 1)[1]
        return 200, {}, json.dumps({"username": addr} if addr == "alice@example.net" else {})

    with responses.RequestsMock() as rsps:
        rsps.add_callback("GET", re.compile(API + "get/mailbox/.*"), callback=get_mailbox)
        r = client.post("/api/mailcow/check",
                        {"addresses": ["alice@example.net", "Ghost@Example.net"]})
        assert r.status == 200, r.body
        assert r.json() == {"exists": {"alice@example.net": True, "Ghost@Example.net": False}}
        assert [c.request.url for c in rsps.calls] == [
            API + "get/mailbox/alice@example.net", API + "get/mailbox/ghost@example.net"]
        assert all(c.request.method == "GET" for c in rsps.calls)


def test_mailcow_check_api_error_is_502_without_key(client):
    with responses.RequestsMock() as rsps:
        rsps.add("GET", API + "get/mailbox/alice@example.net", status=500,
                 body=f"boom key={API_KEY}")
        r = client.post("/api/mailcow/check", {"addresses": ["alice@example.net"]})
    assert r.status == 502
    assert r.json()["error"].startswith("MailcowError: mailcow API HTTP 500")
    assert API_KEY not in r.body.decode()


@pytest.mark.parametrize("body", [
    {"addresses": ["x@example.net"] * 501}, {"addresses": "x@example.net"}, {},
    {"addresses": [7]}, {"addresses": ["not-an-address"]}, ["x@example.net"],
])
def test_mailcow_check_rejects_bad_input(client, body):
    with responses.RequestsMock(assert_all_requests_are_fired=False) as rsps:
        r = client.post("/api/mailcow/check", body)
        assert len(rsps.calls) == 0  # nothing reaches mailcow
    assert r.status == 400


# -- selection -------------------------------------------------------------------------

def test_selection_roundtrip_writes_a_0600_csv_the_cli_reads_isc_159(client, tmp_path):
    r = client.get("/api/selection")
    assert r.status == 200
    assert r.json() == {"path": str(client.selection), "rows": [], "digest": None}
    rows = [
        {"source": " Alice@Contoso.com ", "destination": "Alice@Example.NET",
         "name": "Alice\x07 Liddell", "quota_mib": 2048,
         "aliases": ["Alice.Liddell@Example.net", "alice@example.net", "al@example.net"]},
        {"source": "bob@contoso.com", "destination": "robert@example.net",
         "name": 'Smith, Bob "B"', "quota_mib": None},
        {"source": "info@contoso.com", "destination": "info@example.net", "name": None,
         "quota_mib": None},
    ]
    expected = [  # aliases: normalised, sorted, the destination itself dropped
        {"source": "alice@contoso.com", "destination": "alice@example.net",
         "name": "Alice Liddell", "quota_mib": 2048,
         "aliases": ["al@example.net", "alice.liddell@example.net"]},
        {"source": "bob@contoso.com", "destination": "robert@example.net",
         "name": 'Smith, Bob "B"', "quota_mib": None, "aliases": []},
        {"source": "info@contoso.com", "destination": "info@example.net", "name": "",
         "quota_mib": None, "aliases": []},
    ]
    r = client.put("/api/selection", {"path": "/ignored", "rows": rows})
    assert r.status == 200, r.body
    body = r.json()
    assert isinstance(body.pop("digest"), str)
    assert body == {"path": str(client.selection), "rows": expected}
    assert client.get("/api/selection").json()["rows"] == expected
    assert stat.S_IMODE(client.selection.stat().st_mode) == 0o600
    assert client.selection.read_text(encoding="utf-8").startswith("#")
    assert not [p for p in client.selection.parent.iterdir() if p.name.endswith(".tmp")]
    # exactly the file `--mailboxes` accepts
    cfg = load_config(str(client.conf), mailboxes_csv=str(client.selection))
    assert cfg.mailboxes == (
        MailboxMapping("alice@contoso.com", "alice@example.net", "Alice Liddell", 2048,
                       ("al@example.net", "alice.liddell@example.net")),
        MailboxMapping("bob@contoso.com", "robert@example.net", 'Smith, Bob "B"', None),
        MailboxMapping("info@contoso.com", "info@example.net", None, None),
    )
    # an empty selection is a valid selection
    assert client.put("/api/selection", {"rows": []}).json()["rows"] == []
    assert client.get("/api/selection").json()["rows"] == []


@pytest.mark.parametrize("rows, fragment", [
    ([GOOD, {**GOOD, "source": "b@contoso.com", "destination": "A@Example.net"}],
     "already used by row 1"),
    ([GOOD, {**GOOD, "destination": "b@example.net"}], "duplicate source"),
    ([{**GOOD, "source": "not-an-address"}], "source"),
    ([{**GOOD, "destination": "a,b@example.net"}], "destination"),
    ([{**GOOD, "destination": "#a@example.net"}], "destination"),
    ([{**GOOD, "source": "-x@contoso.com"}], "source"),
    ([{**GOOD, "destination": "a@example.net\nb@example.net"}], "destination"),
    ([{**GOOD, "name": "x" * 201}], "name"),
    ([{**GOOD, "name": 5}], "name"),
    ([{**GOOD, "quota_mib": 0}], "quota_mib"),
    ([{**GOOD, "quota_mib": 1_000_001}], "quota_mib"),
    ([{**GOOD, "quota_mib": "5"}], "quota_mib"),
    ([{**GOOD, "quota_mib": True}], "quota_mib"),
    ([{"source": "a@contoso.com"}], "destination"),
    ([["a@contoso.com", "a@example.net"]], "row"),
    ("a@contoso.com", "rows"),
])
def test_selection_put_validation_isc_159(client, rows, fragment):
    assert client.put("/api/selection", {"rows": [GOOD]}).status == 200
    before = client.selection.read_bytes()
    r = client.put("/api/selection", {"rows": rows})
    assert r.status == 400
    assert fragment in r.json()["error"]
    assert client.selection.read_bytes() == before  # a rejected PUT changes nothing


def test_selection_unreadable_file_is_reported(client):
    client.state_dir.mkdir(parents=True, exist_ok=True)
    client.selection.write_text("alice@contoso.com,alice@example.net,A,lots\n", encoding="utf-8")
    r = client.get("/api/selection")
    assert r.status == 409  # a hand-edited, invalid file: the page says to re-save it
    assert "quota" in r.json()["error"] and "invalid" in r.json()["error"]


# -- jobs ------------------------------------------------------------------------------

def test_job_runs_cli_main_and_captures_output_isc_160_161(client, monkeypatch):
    seen: dict = {}

    def fake_main(argv=None, *, stdout=None, stderr=None):
        seen["argv"] = list(argv)
        print("plan output line", file=stdout)
        stderr.write(f"leak {API_KEY} {CLIENT_SECRET} {TOKEN}\n")
        print("no trailing newline", end="", file=stdout)
        snapshot = argv[argv.index("--mailboxes") + 1]
        seen["snapshot_bytes"] = Path(snapshot).read_bytes()
        seen["snapshot_mode"] = oct(Path(snapshot).stat().st_mode & 0o777)
        return 0

    monkeypatch.setattr(cli, "main", fake_main)
    assert client.put("/api/selection", {"rows": [GOOD]}).status == 200
    digest = client.get("/api/selection").json()["digest"]
    # a job must name the selection it was confirmed against (M2: no TOCTOU on the list)
    stale = client.post("/api/jobs", {"command": "plan", "dry_run": True, "only": None,
                                      "mailbox": None, "sample": 0, "selection_digest": "x"})
    assert stale.status == 409
    r = client.post("/api/jobs", {"command": "plan", "dry_run": True, "only": None,
                                  "mailbox": None, "sample": 0, "selection_digest": digest})
    assert r.status == 202, r.body
    job_id = r.json()["id"]
    job = wait_job(client, job_id)
    argv = seen["argv"]
    snapshot = argv[argv.index("--mailboxes") + 1]  # a private copy of the confirmed bytes
    assert seen["snapshot_bytes"] == client.selection.read_bytes()
    assert seen["snapshot_mode"] == "0o600" and not Path(snapshot).exists()
    assert argv == ["--config", str(client.conf), "--mailboxes", snapshot,
                    "--mailboxes-only", "plan", "--dry-run"]
    assert job["exit_code"] == 0 and job["running"] is False
    assert job["command"] == "plan" and job["args"] == seen["argv"]
    assert job["started"] <= job["finished"]
    assert job["output_tail"] == ["plan output line", "leak *** *** ***", "no trailing newline"]
    out = client.get(f"/api/jobs/{job_id}/output")
    assert out.status == 200
    assert out.headers["content-type"] == "text/plain; charset=utf-8"
    assert out.body.decode() == "plan output line\nleak *** *** ***\nno trailing newline\n"
    assert_security_headers(out)
    for secret in SECRETS:
        assert secret not in out.body.decode() and secret not in json.dumps(job)
    listing = client.get("/api/jobs").json()
    assert [j["id"] for j in listing] == [job_id]
    assert set(listing[0]) == {"id", "command", "args", "started", "finished", "exit_code",
                               "dry_run", "outcome"}
    assert listing[0]["dry_run"] is True and listing[0]["outcome"] is None  # wrote no report
    assert job["progress"] is None  # plan prints no progress lines


def _saved(client) -> str:
    """Save one row and return the digest a job must present."""
    return client.put("/api/selection", {"rows": [GOOD]}).json()["digest"]


def test_job_options_become_cli_arguments_isc_160(client, monkeypatch):
    argvs: list[list[str]] = []
    monkeypatch.setattr(cli, "main", lambda argv, **kw: argvs.append(list(argv)) or 1)
    digest = _saved(client)
    r = client.post("/api/jobs", {"command": "verify", "dry_run": False, "only": "mail",
                                  "mailbox": "Alice@Example.net", "sample": 5,
                                  "selection_digest": digest})
    job = wait_job(client, r.json()["id"])
    assert job["exit_code"] == 1
    snapshot = argvs[0][3]
    assert snapshot.startswith(str(client.state_dir / "jobs" / "selection-"))
    assert argvs[0] == ["--config", str(client.conf), "--mailboxes", snapshot,
                        "--mailboxes-only", "verify", "--only", "mail",
                        "--mailbox=alice@example.net", "--sample", "5"]
    assert not Path(snapshot).exists()  # removed when the job ended
    r = client.post("/api/jobs", {"command": "migrate", "sample": 5,  # sample: verify only
                                  "selection_digest": digest})
    wait_job(client, r.json()["id"])
    assert argvs[1][4:] == ["--mailboxes-only", "migrate"]
    listing = client.get("/api/jobs").json()
    assert [j["command"] for j in listing] == ["migrate", "verify"]  # newest first


@pytest.mark.parametrize("body", [
    {"command": "rm"}, {"command": "web"}, {}, {"command": "plan", "extra": 1},
    {"command": "plan", "only": "email"}, {"command": "plan", "dry_run": "yes"},
    {"command": "verify", "sample": -1}, {"command": "verify", "sample": True},
    {"command": "verify", "sample": "3"}, {"command": "plan", "mailbox": "--config=/etc/x"},
    {"command": "plan", "mailbox": 7}, ["plan"],
])
def test_job_rejects_unknown_commands_and_options(client, monkeypatch, body):
    started = []
    monkeypatch.setattr(cli, "main", lambda *a, **kw: started.append(a) or 0)
    r = client.post("/api/jobs", body)
    assert r.status == 400
    assert "error" in r.json()
    assert started == [] and client.get("/api/jobs").json() == []


def test_second_job_while_one_runs_is_409_isc_160(client, monkeypatch):
    release, entered = threading.Event(), threading.Event()

    def blocking_main(argv=None, *, stdout=None, stderr=None):
        entered.set()
        release.wait(15)
        print("finished", file=stdout)
        return 1

    monkeypatch.setattr(cli, "main", blocking_main)
    try:
        digest = _saved(client)
        first = client.post("/api/jobs", {"command": "migrate", "selection_digest": digest})
        assert first.status == 202
        job_id = first.json()["id"]
        assert entered.wait(5)
        second = client.post("/api/jobs", {"command": "plan", "selection_digest": digest})
        assert second.status == 409
        assert job_id in second.json()["error"]
        assert client.get("/api/status").json()["running_job"] == {"id": job_id,
                                                                  "command": "migrate"}
        running = client.get(f"/api/jobs/{job_id}").json()
        assert running["running"] is True and running["exit_code"] is None
        assert running["finished"] is None
    finally:
        release.set()
    assert wait_job(client, job_id)["exit_code"] == 1
    assert client.get("/api/status").json()["running_job"] is None
    third = client.post("/api/jobs", {"command": "plan", "selection_digest": digest})
    assert third.status == 202
    wait_job(client, third.json()["id"])


def test_job_crash_and_argparse_exit_are_recorded(client, monkeypatch):
    digest = _saved(client)
    def crashing_main(argv=None, *, stdout=None, stderr=None):
        raise RuntimeError(f"exploded with {API_KEY}")

    monkeypatch.setattr(cli, "main", crashing_main)
    job = wait_job(client, client.post("/api/jobs", {"command": "plan",
                                                     "selection_digest": digest}).json()["id"])
    assert job["exit_code"] == 1
    assert job["output_tail"] == ["job failed: RuntimeError: exploded with ***"]

    def exiting_main(argv=None, *, stdout=None, stderr=None):
        raise SystemExit(2)

    monkeypatch.setattr(cli, "main", exiting_main)
    job = wait_job(client, client.post("/api/jobs", {"command": "plan",
                                                     "selection_digest": digest}).json()["id"])
    assert job["exit_code"] == 2


def test_real_cli_job_output_never_contains_secrets_isc_160(client):
    rows = [{**GOOD, "source": "alice@contoso.com", "destination": "alice@example.net"},
            {**GOOD, "source": "bob@contoso.com", "destination": "bob@example.net"}]
    assert client.put("/api/selection", {"rows": rows}).status == 200

    def list_pw(req):
        if req.url.endswith("bob@example.net"):  # an error body echoing the key
            return 500, {}, f"internal error, X-API-Key: {API_KEY}"
        return 200, {}, json.dumps([{"id": 1, "name": "o365-migration-abc"}])

    with responses.RequestsMock(assert_all_requests_are_fired=False) as rsps:
        rsps.add_callback("GET", re.compile(API + "get/app-passwd/all/.*"), callback=list_pw)
        r = client.post("/api/jobs", {"command": "cleanup", "dry_run": True})
        assert r.status == 202
        job = wait_job(client, r.json()["id"])
        assert all(c.request.method == "GET" for c in rsps.calls)
    text = client.get(f"/api/jobs/{job['id']}/output").body.decode()
    assert job["exit_code"] == 1, text  # bob's listing failed
    assert "would delete app password id=1 of alice@example.net" in text
    assert "bob@example.net: cleanup failed" in text
    assert "report: " in text
    for secret in SECRETS:
        assert secret not in text


def test_job_output_buffer_is_capped_and_redacts_across_writes():
    out = web.JobOutput(lambda s: s.replace("hunter2", "***"), max_lines=3)
    out.write("a\nb\n")
    out.write("c\nhunt")
    out.write("er2\n")  # a secret split over two writes is still redacted
    out.write("e")
    out.close()
    assert out.tail(10) == ["c", "***", "e"]
    assert out.tail(1) == ["e"]
    assert out.text() == "[2 earlier line(s) dropped]\nc\n***\ne\n"


# -- reports ---------------------------------------------------------------------------

def test_latest_report_isc_161(client):
    assert client.get("/api/reports/latest").status == 404
    reports = client.state_dir / "reports"
    reports.mkdir(parents=True)
    old = reports / "20260101T000000Z.json"
    old.write_text(json.dumps({"command": "verify", "exit_code": 0}), encoding="utf-8")
    new = reports / "20260102T000000Z.json"
    new.write_text(json.dumps({"command": "plan", "exit_code": 1}), encoding="utf-8")
    (reports / "notes.txt").write_text("ignored", encoding="utf-8")
    os.utime(old, (1_000, 1_000))
    os.utime(new, (2_000, 2_000))
    assert client.get("/api/reports/latest").json() == {"command": "plan", "exit_code": 1}
    assert client.get("/api/reports/latest?command=verify").json()["command"] == "verify"
    assert client.get("/api/reports/latest?command=cleanup").status == 404
    assert client.get("/api/reports/latest?command=bogus").status == 400


# -- cli integration -------------------------------------------------------------------

def test_cli_main_writes_to_the_given_streams_isc_167(conf, tmp_path, capsys):
    csv_path = tmp_path / "boxes.csv"
    csv_path.write_text("alice@contoso.com,alice@example.net\n", encoding="utf-8")
    out, err = io.StringIO(), io.StringIO()
    with responses.RequestsMock() as rsps:
        rsps.add("GET", API + "get/app-passwd/all/alice@example.net",
                 json=[{"id": 1, "name": "o365-migration"}])
        code = cli.main(["--config", str(conf), "--mailboxes", str(csv_path), "--dry-run",
                         "cleanup"], stdout=out, stderr=err)
    assert code == 0
    assert "would delete app password id=1 of alice@example.net" in out.getvalue()
    assert "cleanup: dry run, deleted 0 app password(s)" in out.getvalue()
    assert "report: " in err.getvalue()
    assert f"o365mig {__version__} cleanup: 1 mailbox(es)" in err.getvalue()  # log handler
    assert cli.main(["--config", str(tmp_path / "missing.toml"), "plan"], stderr=err) == 2
    assert "configuration error" in err.getvalue()
    captured = capsys.readouterr()
    assert captured.out == "" and captured.err == ""


def test_runner_sends_prompts_and_progress_to_the_given_streams(tmp_path, monkeypatch):
    seen = {}

    class Tokens:
        def __init__(self, cfg, cache_path=None, out=None) -> None:
            seen["out"] = out

    monkeypatch.setattr(cli, "TokenProvider", Tokens)
    out, err = io.StringIO(), io.StringIO()
    opts = cli.Options("plan", None, None, None, None, False, False, 0, 0, False)
    runner = cli.Runner(make_config(tmp_path), opts, [], cli.SecretFilter(), out=out, err=err)
    assert runner.tokens and seen["out"] is out  # device-code prompt goes to the job's stdout
    runner.progress.start("alice", 3)
    runner.progress.maybe_print(force=True)
    assert "[alice] 0/3 items" in err.getvalue()


def test_cli_web_prints_the_url_and_serves_isc_155(conf, monkeypatch):
    calls = []
    monkeypatch.setattr(web, "serve", lambda *a, **kw: calls.append((a, kw)))
    monkeypatch.setenv("O365MIG_WEB_TOKEN", TOKEN)
    err = io.StringIO()
    assert cli.main(["--config", str(conf), "web", "--port", "8099"], stderr=err) == 0
    # a token that came from the environment is not repeated in the log
    assert "web UI: http://127.0.0.1:8099/#token=<O365MIG_WEB_TOKEN from .env>\n" in err.getvalue()
    assert TOKEN not in err.getvalue()
    assert "WARNING" not in err.getvalue()
    (cfg, bind, port, token, page), kw = calls[0]
    assert (bind, port, token) == ("127.0.0.1", 8099, TOKEN)
    assert page.name == "index.html" and page.parent.name == "web_static"
    assert kw["config_path"] == str(conf.resolve())
    assert cfg.mailboxes == ()  # `web` needs no mailbox list
    err = io.StringIO()
    assert cli.main(["--config", str(conf), "web"], stderr=err) == 0
    assert "web UI: http://127.0.0.1:8080/#token=" in err.getvalue()
    # every other command still requires a mailbox list
    assert cli.main(["--config", str(conf), "plan"], stderr=err) == 2
    assert "run.mailboxes" in err.getvalue()


def test_cli_web_generates_a_token_and_warns_off_loopback_isc_155(conf, monkeypatch):
    calls = []
    monkeypatch.setattr(web, "serve", lambda *a, **kw: calls.append(a))
    err = io.StringIO()
    wildcard = "0.0.0.0"  # noqa: S104 - the off-loopback warning is what this tests
    assert cli.main(["--config", str(conf), "web", "--bind", wildcard], stderr=err) == 0
    text = err.getvalue()
    token = re.search(r"web UI: http://0\.0\.0\.0:8080/#token=(\S+)", text).group(1)
    assert len(token) >= 43 and calls[0][3] == token
    assert "WARNING" in text and "reachable from the network" in text
    err = io.StringIO()
    assert cli.main(["--config", str(conf), "web", "--bind", "::1"], stderr=err) == 0
    assert "http://[::1]:8080/" in err.getvalue() and "WARNING" not in err.getvalue()


def test_cli_web_refuses_a_short_token_and_reports_bind_errors(conf, monkeypatch):
    def busy(*a, **kw):
        raise OSError(98, "Address already in use")

    monkeypatch.setattr(web, "serve", busy)
    err = io.StringIO()
    assert cli.main(["--config", str(conf), "web"], stderr=err) == 2
    assert "Address already in use" in err.getvalue()
    monkeypatch.setenv("O365MIG_WEB_TOKEN", "short")
    err = io.StringIO()
    assert cli.main(["--config", str(conf), "web"], stderr=err) == 2
    assert "O365MIG_WEB_TOKEN" in err.getvalue()


def test_serve_binds_serves_and_stops_on_keyboard_interrupt(tmp_path, monkeypatch):
    page = tmp_path / "index.html"
    page.write_text(PAGE, encoding="utf-8")
    cfg = make_config(tmp_path, mailboxes=())

    def interrupted(self, poll_interval=0.5):
        raise KeyboardInterrupt

    monkeypatch.setattr(web.WebServer, "serve_forever", interrupted)
    web.serve(cfg, "127.0.0.1", 0, TOKEN, page)  # returns instead of raising
    with pytest.raises(OSError):
        web.serve(cfg, "127.0.0.1", 0, TOKEN, tmp_path / "missing.html")


# -- hardening after the security review of the web UI ------------------------------------

def test_host_change_requires_the_api_key_in_the_same_request(client, tmp_path):
    """A token holder must not be able to point the stored API key at another host."""
    base = {"microsoft": {"tenant_id": "t1", "client_id": "c1", "auth_mode": "app",
                          "client_secret": "secret-one-value"},
            "mailcow": {"host": "mail.example.net", "api_key": "key-one-value"}}
    assert client.put("/api/settings", base).status == 200
    moved = {"microsoft": base["microsoft"], "mailcow": {"host": "attacker.example"}}
    r = client.put("/api/settings", moved)
    assert r.status == 400 and b"API key together with the host" in r.body
    assert "attacker.example" not in (tmp_path / "state" / "settings.toml").read_text()
    # same for the tenant/client ids and the client secret
    ids = {"microsoft": {"tenant_id": "t2", "client_id": "c2", "auth_mode": "app"},
           "mailcow": {"host": "mail.example.net"}}
    r = client.put("/api/settings", ids)
    assert r.status == 400 and b"client secret again" in r.body
    # with the paired secrets the change is accepted
    moved["mailcow"]["api_key"] = "key-two-value"
    assert client.put("/api/settings", moved).status == 200
    text = (tmp_path / "state" / "settings.toml").read_text()
    assert 'host = "attacker.example"' in text and 'api_key = "key-two-value"' in text
    assert "key-one-value" not in text


def test_settings_reject_ip_literals_and_del_characters(client):
    body = {"microsoft": {"tenant_id": "t", "client_id": "c", "auth_mode": "app",
                          "client_secret": "s3cret-value"},
            "mailcow": {"host": "169.254.169.254", "api_key": "key-value"}}
    assert client.put("/api/settings", body).status == 400
    body["mailcow"]["host"] = "mail.example.net"
    body["mailcow"]["api_key"] = "key\x7fvalue"
    assert client.put("/api/settings", body).status == 400


def test_saved_host_never_borrows_a_secret_from_env_or_file(tmp_path, conf, monkeypatch):
    """config.load_config pairs a saved host only with a saved key (H1)."""
    from o365_to_mailcow import config as config_mod

    state = tmp_path / "state"
    state.mkdir(exist_ok=True)
    (state / "settings.toml").write_text('[mailcow]\nhost = "other.example.net"\n')
    with pytest.raises(config_mod.ConfigError, match="mailcow.api_key"):
        config_mod.load_config(str(conf), require_mailboxes=False)
    cfg = config_mod.load_config(str(conf), require_mailboxes=False, require_credentials=False)
    assert cfg.mailcow_host == "other.example.net" and cfg.mailcow_api_key == ""


def test_host_header_and_origin_are_checked(client):
    r = client.get("/api/status", headers={"Host": "rebind.attacker.example"})
    assert r.status == 421
    r = client.post("/api/mailcow/check", {"addresses": []},
                    headers={"Origin": "https://evil.example"})
    assert r.status == 403
    r = client.post("/api/mailcow/check", {"addresses": []},
                    headers={"Origin": f"http://127.0.0.1:{client.port}"})
    assert r.status == 200


def test_unauthenticated_requests_do_not_read_the_body(client):
    r = client.request("POST", "/api/jobs", token=None,
                       headers={"Content-Length": "5000000"}, raw=b"x")
    assert r.status == 401


def test_lock_settings_refuses_changes(tmp_path, conf, graph, monkeypatch):
    from o365_to_mailcow import config as config_mod

    cfg = config_mod.load_config(str(conf), require_mailboxes=False)
    page = Path(web.__file__).parent / "web_static" / "index.html"
    srv = web.make_server(cfg, "127.0.0.1", 0, TOKEN, page, config_path=str(conf),
                          lock_settings=True)
    thread = threading.Thread(target=srv.serve_forever, daemon=True)
    thread.start()
    try:
        c = Client(srv.server_address[1])
        body = {"microsoft": {"tenant_id": "t", "client_id": "c", "auth_mode": "app",
                              "client_secret": "s"}, "mailcow": {"host": "mail.example.net",
                                                                  "api_key": "k"}}
        assert c.put("/api/settings", body).status == 403
        assert c.get("/api/settings").status == 200
    finally:
        srv.shutdown()
        srv.server_close()


def test_job_output_is_byte_capped():
    out = web.JobOutput(lambda t: t, max_lines=10_000, max_bytes=1000)
    for i in range(200):
        out.write(f"line {i:04d} " + "x" * 40 + "\n")
    text = out.text()
    assert len(text) <= 1400 and "line 0199" in text and "line 0000" not in text


# -- second review round -----------------------------------------------------------------

def test_host_change_requires_the_client_secret_again_and_drops_sign_ins(client, tmp_path):
    """Silas #1: a host change with a fresh key must not keep the Microsoft credential."""
    state = tmp_path / "state"
    base = {"microsoft": {"tenant_id": "t1", "client_id": "c1", "auth_mode": "app",
                          "client_secret": "secret-one-value"},
            "mailcow": {"host": "mail.example.net", "api_key": "key-one-value"}}
    assert client.put("/api/settings", base).status == 200
    (state / "msal_cache_0123456789abcdef.bin").write_text("cached refresh token")
    (state / "msal_cache_web_0123456789abcdef.bin").write_text("cached refresh token")
    moved = {"microsoft": {"tenant_id": "t1", "client_id": "c1", "auth_mode": "app"},
             "mailcow": {"host": "sink.example.org", "api_key": "attacker-key-value"}}
    r = client.put("/api/settings", moved)
    assert r.status == 400 and b"client secret again" in r.body
    assert "sink.example.org" not in (state / "settings.toml").read_text()
    assert len(list(state.glob("msal_cache*.bin"))) == 2  # a refused save keeps sign-ins
    moved["microsoft"]["client_secret"] = "secret-two-value"
    assert client.put("/api/settings", moved).status == 200
    assert not list(state.glob("msal_cache*.bin"))


def test_put_selection_returns_digest_and_job_accepts_it(client, monkeypatch):
    calls: list[list[str]] = []

    def fake_main(argv=None, *, stdout=None, stderr=None):
        calls.append(list(argv))
        return 0

    monkeypatch.setattr(cli, "main", fake_main)
    r = client.put("/api/selection", {"rows": [GOOD]})
    assert r.status == 200
    digest = r.json()["digest"]
    assert digest == hashlib.sha256(client.selection.read_bytes()).hexdigest()
    r = client.post("/api/jobs", {"command": "plan", "dry_run": True, "only": None,
                                  "mailbox": None, "sample": 0, "selection_digest": digest})
    assert r.status == 202, r.body
    wait_job(client, r.json()["id"])
    argv = calls[0]
    # the job read a private snapshot of the confirmed bytes, not the live file
    snapshot = argv[argv.index("--mailboxes") + 1]
    assert snapshot != str(client.selection) and "--mailboxes-only" in argv
    assert not Path(snapshot).exists()  # deleted once the job ended (it read it first)


def test_job_without_a_saved_selection_is_refused(client):
    r = client.post("/api/jobs", {"command": "migrate", "dry_run": True, "only": None,
                                  "mailbox": None, "sample": 0, "selection_digest": "x"})
    assert r.status == 409 and b"no saved selection" in r.body


def test_origin_on_another_localhost_port_is_refused(client):
    port = client.port
    ok = client.post("/api/mailcow/check", {"addresses": []},
                     headers={"Origin": f"http://127.0.0.1:{port}"})
    assert ok.status == 200
    other = client.post("/api/mailcow/check", {"addresses": []},
                        headers={"Origin": "http://127.0.0.1:1234"})
    assert other.status == 403


def test_job_output_bytes_are_counted_after_truncation():
    out = web.JobOutput(lambda t: t, max_lines=10_000, max_bytes=200_000)
    for _ in range(300):
        out.write("x" * 60_000 + "\n")
    for i in range(5):
        out.write(f"summary line {i}\n")
    text = out.text()
    assert all(f"summary line {i}" in text for i in range(5))


# -- third review round (delta re-review) ------------------------------------------------

def _settings(tenant="t1", client_id="c1", mode="app", secret=None, host="mail.example.net",
              key=None) -> dict:
    ms = {"tenant_id": tenant, "client_id": client_id, "auth_mode": mode}
    if secret is not None:
        ms["client_secret"] = secret
    mc = {"host": host}
    if key is not None:
        mc["api_key"] = key
    return {"microsoft": ms, "mailcow": mc}


def test_switching_the_sign_in_mode_never_carries_the_client_secret(client, tmp_path):
    """BLOCKER: app -> delegated (new host, own key) -> app must not revive the saved secret."""
    settings = tmp_path / "state" / "settings.toml"
    assert client.put("/api/settings", _settings(secret="secret-one-value",
                                                 key="key-one-value")).status == 200
    assert "secret-one-value" in settings.read_text()
    # delegated mode needs no secret, so a token holder can move the host with their own key
    r = client.put("/api/settings", _settings(mode="delegated", host="sink.example.org",
                                              key="attacker-key-value"))
    assert r.status == 200, r.body
    assert "secret-one-value" not in settings.read_text()  # gone, not merely hidden
    # ... and switching back to app-only must demand the secret again
    r = client.put("/api/settings", _settings(host="sink.example.org"))
    assert r.status == 400 and b"client secret again" in r.body
    assert "secret-one-value" not in settings.read_text()
    assert "auth_mode = \"delegated\"" in settings.read_text()  # the refused save changed nothing


def test_delegated_mode_stores_no_secret_even_when_one_is_sent(client, tmp_path):
    settings = tmp_path / "state" / "settings.toml"
    body = _settings(mode="delegated", secret="secret-one-value", key="key-one-value")
    assert client.put("/api/settings", body).status == 200
    assert "secret-one-value" not in settings.read_text()
    # a plain mode switch (same ids, same host) still re-pairs: secret required
    r = client.put("/api/settings", _settings())
    assert r.status == 400 and b"sign-in mode" in r.body


def test_settings_cannot_change_while_a_job_runs(client, monkeypatch, tmp_path):
    release = threading.Event()
    monkeypatch.setattr(cli, "main", lambda *a, **kw: (release.wait(10), 0)[1])
    digest = _saved(client)
    r = client.post("/api/jobs", {"command": "plan", "selection_digest": digest})
    assert r.status == 202
    try:
        r = client.put("/api/settings", _settings(secret="secret-one-value", key="key-one"))
        assert r.status == 409 and b"job is running" in r.body
        assert not (tmp_path / "state" / "settings.toml").exists()
    finally:
        release.set()


def test_msal_caches_are_keyed_by_connection(tmp_path, monkeypatch):
    from o365_to_mailcow.auth import TokenProvider
    from o365_to_mailcow.config import connection_id

    a = make_config(tmp_path, mailcow_host="mail.example.net")
    b = make_config(tmp_path, mailcow_host="sink.example.org")
    assert connection_id(a) != connection_id(b)
    monkeypatch.setattr("msal.ConfidentialClientApplication", lambda *a, **kw: object())
    assert TokenProvider(a)._cache_path.name == f"msal_cache_{connection_id(a)}.bin"
    assert TokenProvider(b)._cache_path.name == f"msal_cache_{connection_id(b)}.bin"
    page = tmp_path / "index.html"
    page.write_text(PAGE, encoding="utf-8")
    app = web.WebApp(a, TOKEN, PAGE)
    assert app._web_cache_path().name == f"msal_cache_web_{connection_id(a)}.bin"


def test_stale_job_snapshots_are_removed_at_startup(tmp_path):
    jobs = tmp_path / "state" / "jobs"
    jobs.mkdir(parents=True)
    stale = jobs / "selection-dead.csv"
    stale.write_text("source,destination\n")
    other = jobs / "keep.txt"
    other.write_text("x")
    web.WebApp(make_config(tmp_path), TOKEN, PAGE)
    assert not stale.exists() and other.exists()


def test_duplicate_host_header_is_refused(client):
    conn = http.client.HTTPConnection("127.0.0.1", client.port, timeout=5)
    conn.putrequest("GET", "/api/status", skip_host=True)
    conn.putheader("Host", f"127.0.0.1:{client.port}")
    conn.putheader("Host", "evil.example")
    conn.putheader("Authorization", f"Bearer {TOKEN}")
    conn.endheaders()
    resp = conn.getresponse()
    assert resp.status == 400
    conn.close()


@pytest.mark.parametrize("origin, status", [
    ("http://127.0.0.1:{port}", 200), ("http://localhost:{port}", 403),
    ("null", 403), ("http://evil.example", 403),
])
def test_origin_must_equal_the_host_authority(client, origin, status):
    r = client.post("/api/mailcow/check", {"addresses": []},
                    headers={"Origin": origin.format(port=client.port)})
    assert r.status == status


def test_selection_refuses_alias_collisions_across_rows(client):
    a = dict(GOOD, aliases=["shared@example.net"])
    b = dict(GOOD, source="bob@example.com", destination="bob@example.net",
             aliases=["shared@example.net"])
    r = client.put("/api/selection", {"rows": [a, b]})
    assert r.status == 400 and b"row 2" in r.body and b"shared@example.net" in r.body
    c = dict(GOOD, source="carol@example.com", destination="shared@example.net")
    r = client.put("/api/selection", {"rows": [a, c]})
    assert r.status == 400 and b"already used by row 1" in r.body
    assert not client.selection.exists()


def test_selection_with_an_invalid_saved_alias_is_a_409(client):
    _saved(client)
    text = client.selection.read_text()
    client.selection.write_text(text.rstrip("\n") + "not-an-address\n")
    r = client.get("/api/selection")
    assert r.status == 409 and b"saved selection is invalid" in r.body


def test_allow_host_accepts_bracketed_ipv6_with_port():
    assert web.host_only("[::1]:8080") == "::1"
    assert web.host_only("[::1]") == "::1"
    assert web.host_only("ui.example.net:8080") == "ui.example.net"
    assert web.host_only("ui.example.net") == "ui.example.net"


# -- overview, job outcome and progress (web UI redesign) -----------------------------------

ANNA = "anna@example.com"
FULL = {"only": None, "mailbox": None, "mail_since": None}  # what cli.Runner records


def _write_report(state_dir: Path, name: str, data: object, mtime: float) -> Path:
    reports = state_dir / "reports"
    reports.mkdir(parents=True, exist_ok=True)
    path = reports / name
    path.write_text(data if isinstance(data, str) else json.dumps(data), encoding="utf-8")
    os.utime(path, (mtime, mtime))
    return path


def _verify_report(imap: int, exit_code: int, started: str = "2026-10-01T01:00:00+00:00",
                   dry_run: bool = False) -> dict:
    mail = MailVerify(ANNA, [FolderVerify("INBOX", 10, 10, 0, 0, imap, 10, imap != 10)])
    return {"command": "verify", "dry_run": dry_run, "started": started, "scope": FULL,
            "finished": "2026-10-01T02:00:00+00:00", "duration_s": 3.0, "exit_code": exit_code,
            "mailboxes": {ANNA: {"destination": "anna@example.net", "status": "ok",
                                 "errors": [], "mail": _plain(mail)}}}


def _migrate_report(copied: int, dry_run: bool) -> dict:
    folder = {"folder_id": "f", "dest_name": "INBOX", "listed": copied, "appended": copied,
              "already_done": 0, "dedup_hits": 0, "failed": 0, "skipped_too_large": 0,
              "would_append": copied if dry_run else 0, "error": None}
    return {"command": "migrate", "dry_run": dry_run, "started": "2026-10-01T00:00:00+00:00",
            "scope": FULL,
            "finished": "2026-10-01T00:30:00+00:00", "duration_s": 1.0, "exit_code": 0,
            "mailboxes": {ANNA: {"destination": "anna@example.net", "status": "ok",
                                 "errors": [], "mail": {"folders": [folder],
                                                        "skipped_folders": [], "errors": [],
                                                        "stopped": False}}}}


def test_overview_without_reports_is_all_null(client):
    r = client.get("/api/overview")
    assert r.status == 200
    assert r.json() == {"steps": {"provision": None, "migrate": None, "verify": None,
                                  "cleanup": None}, "latest": None}
    assert_security_headers(r)


def test_overview_steps_use_the_newest_real_run_of_each_command(client):
    state = client.state_dir
    _write_report(state, "1.json", _verify_report(10, 0), 1_000)
    _write_report(state, "2.json", _verify_report(7, 1), 2_000)  # newest real verify
    _write_report(state, "3.json", _migrate_report(5, dry_run=False), 3_000)
    _write_report(state, "4.json", _migrate_report(99, dry_run=True), 4_000)  # dry run: no step
    _write_report(state, "5.json", {"command": "cleanup", "dry_run": False, "exit_code": 0,
                                    "mailboxes": {"anna@example.net": {
                                        "errors": [], "deleted_app_passwords": 2}}}, 5_000)
    _write_report(state, "0.json", "{cut short", 500)  # unreadable and oldest: skipped
    (state / "reports" / "7.json").symlink_to(state / "reports" / "1.json")  # never followed
    _write_report(state, "8.json", {"command": "plan", "dry_run": True, "exit_code": 0,
                                    "mailboxes": {}}, 5_500)
    data = client.get("/api/overview").json()
    steps = data["steps"]
    assert steps["provision"] is None
    assert steps["verify"] == {"started": "2026-10-01T01:00:00+00:00",
                               "finished": "2026-10-01T02:00:00+00:00", "exit_code": 1,
                               "level": "warn", "text": "3 items missing"}
    assert steps["migrate"]["text"] == "5 items copied" and steps["migrate"]["level"] == "ok"
    assert steps["cleanup"]["text"] == "2 deleted"
    raw_latest = client.get("/api/reports/latest").json()
    assert raw_latest["command"] == "plan"  # the newest readable regular file, dry run or not
    assert data["latest"] == summarize_report(raw_latest)


def test_overview_with_only_garbage_reports(client):
    _write_report(client.state_dir, "1.json", "not json", 1_000)
    _write_report(client.state_dir, "2.json", json.dumps([1, 2, 3]), 2_000)
    data = client.get("/api/overview").json()
    assert all(step is None for step in data["steps"].values())
    assert data["latest"]["headline"] == {"level": "unknown", "text": "Run ended", "detail": ""}


def test_overview_reads_at_most_200_report_files(client, monkeypatch):
    for i in range(250):  # plan reports never complete the steps, so the scan cannot stop
        _write_report(client.state_dir, f"{i:04d}.json",
                      {"command": "plan", "dry_run": True, "mailboxes": {}}, 1_000 + i)
    reads: list[str] = []
    original = web.WebApp._read_report

    def counting(self, path):
        reads.append(path.name)
        return original(self, path)

    monkeypatch.setattr(web.WebApp, "_read_report", counting)
    data = client.get("/api/overview").json()
    assert len(reads) == web.MAX_REPORTS_SCANNED == 200
    assert reads[0] == "0249.json"  # newest first
    assert data["latest"]["command"] == "plan"


def test_overview_stops_reading_once_every_step_is_known(client, monkeypatch):
    state = client.state_dir
    _write_report(state, "1.json", {"command": "plan", "mailboxes": {}}, 1_000)
    _write_report(state, "2.json", _verify_report(10, 0), 2_000)
    _write_report(state, "3.json", _migrate_report(1, False), 3_000)
    _write_report(state, "4.json", {"command": "provision", "dry_run": False, "exit_code": 0,
                                    "mailboxes": {ANNA: {"errors": [],
                                                         "provision": "created"}}}, 4_000)
    _write_report(state, "5.json", {"command": "cleanup", "dry_run": False, "exit_code": 0,
                                    "mailboxes": {}}, 5_000)
    reads: list[str] = []
    original = web.WebApp._read_report
    monkeypatch.setattr(web.WebApp, "_read_report",
                        lambda self, path: reads.append(path.name) or original(self, path))
    steps = client.get("/api/overview").json()["steps"]
    assert steps["provision"]["text"] == "1 created, 0 existed"
    assert reads == ["5.json", "4.json", "3.json", "2.json"]  # 1.json is never read


def _progress_lines(clock: dict) -> tuple[web.JobOutput, Progress]:
    out = web.JobOutput(lambda t: t)
    return out, Progress(out=out, interval=1e9, clock=lambda: clock["t"])  # forced prints only


def test_progress_parses_exactly_what_report_progress_prints():
    """Guard: the printer (report.Progress) and the parser (JobOutput) cannot drift apart."""
    clock = {"t": 0.0}
    out, p = _progress_lines(clock)
    assert out.progress() is None
    p.start(f"{ANNA} mail", 150_374)
    p.start(f"{ANNA} calendar", 40)
    p.start("ben@example.com contacts", 12)
    clock["t"] = 60.0
    p.advance(f"{ANNA} mail", 310)
    p.phase(f"{ANNA} mail", "indexing Inbox: 800/3000 (pass 2)")
    p.advance(f"{ANNA} calendar", 40)
    p.advance("ben@example.com contacts", 12)
    p.maybe_print(force=True)
    p.finish("ben@example.com contacts")
    printed = out.tail(10)
    assert printed == [
        f"[{ANNA} mail] 310/150374 items, 310/min (indexing Inbox: 800/3000 (pass 2))",
        f"[{ANNA} calendar] 40/40 items, 40/min",
        "[ben@example.com contacts] 12/12 items, 12/min",
        "[ben@example.com contacts] 12/12 items, 12/min (finished)",
    ]
    assert out.progress() == {
        "done": 362, "total": 150_426, "rate": 350,  # the finished scope adds no rate
        "scopes": [
            {"mailbox": ANNA, "kind": "mail", "done": 310, "total": 150_374, "rate": 310,
             "phase": "indexing Inbox: 800/3000 (pass 2)", "finished": False},
            {"mailbox": ANNA, "kind": "calendar", "done": 40, "total": 40, "rate": 40,
             "phase": "", "finished": False},
            {"mailbox": "ben@example.com", "kind": "contacts", "done": 12, "total": 12,
             "rate": 12, "phase": "", "finished": True},
        ],
        "more_scopes": 0,
    }


def test_progress_ignores_lines_that_do_not_match():
    out = web.JobOutput(lambda t: t)
    lines = [
        "2026-10-01 01:00:00 INFO o365mig 0.1.0 migrate: 1 mailbox(es)",
        "[anna@example.com mail] 5/10 items, 3/min trailing",
        "[anna@example.com email] 5/10 items, 3/min",
        "[anna@example.com mail] 5/10 items",
        "[] 5/10 items, 3/min",
        "[",
        "[anna@example.com mail] 99999999999999999999/1 items, 0/min",
        "[anna@example.com mail] \u0663/10 items, 0/min",
        "[anna@example.com mail] 5/10 items, 3/min (" + "x" * 5000 + ")",
    ]
    for line in lines:
        out.write(line + "\n")
    assert out.progress() is None
    assert len(out.tail(100)) == len(lines)  # all stored as ordinary output


def test_progress_parsing_can_never_break_a_write(monkeypatch):
    def boom(self, line):
        raise RuntimeError("parser bug")

    monkeypatch.setattr(web.JobOutput, "_track_progress", boom)
    out = web.JobOutput(lambda t: t)
    assert out.write("[anna@example.com mail] 1/2 items, 1/min\n") > 0
    assert out.tail(1) == ["[anna@example.com mail] 1/2 items, 1/min"]


def test_progress_is_parsed_after_redaction_and_cleaning():
    out = web.JobOutput(lambda t: t.replace("hunter2", "***"))
    out.write("[anna@example.com mail] 1/2 items, 1/min (listing hunter2\u202e)\n")
    scope = out.progress()["scopes"][0]
    assert scope["phase"] == "listing ***"


def test_progress_scopes_are_capped():
    out = web.JobOutput(lambda t: t)
    for i in range(web.MAX_SCOPES + 1):  # the last one is ignored
        out.write(f"[user{i:05d}@example.com mail] 1/2 items, 1/min\n")
    out.write("[user00000@example.com mail] 2/2 items, 0/min (finished)\n")  # known: updated
    progress = out.progress()
    assert len(progress["scopes"]) == web.MAX_SCOPES_SHOWN == 600
    assert progress["more_scopes"] == web.MAX_SCOPES - 600
    assert progress["done"] == web.MAX_SCOPES + 1 and progress["total"] == 2 * web.MAX_SCOPES
    assert progress["rate"] == web.MAX_SCOPES - 1
    assert progress["scopes"][0] == {"mailbox": "user00000@example.com", "kind": "mail",
                                     "done": 2, "total": 2, "rate": 0, "phase": "",
                                     "finished": True}


def _fake_report_main(state_dir: Path, imap: int, *, progress: bool = False, code: int = 1):
    def fake_main(argv=None, *, stdout=None, stderr=None):
        if progress:
            stderr.write(f"[{ANNA} mail] 7/10 items, 7/min (finished)\n")
        rep = RunReport("verify", state_dir)
        rep.mailbox(ANNA, "anna@example.net")
        rep.set(ANNA, "mail", MailVerify(ANNA, [FolderVerify("INBOX", 10, 10, 0, 0, imap, 10,
                                                             imap != 10)]))
        rep.set(ANNA, "status", "ok")
        rep.write(code)
        return code
    return fake_main


def test_job_outcome_is_the_headline_of_the_report_it_wrote(client, monkeypatch):
    monkeypatch.setattr(cli, "main", _fake_report_main(client.state_dir, 7, progress=True))
    digest = _saved(client)
    job = wait_job(client, client.post("/api/jobs", {"command": "verify",
                                                     "selection_digest": digest}).json()["id"])
    assert job["exit_code"] == 1 and job["dry_run"] is False
    assert job["outcome"] == {"level": "warn", "text": "3 items have not arrived in mailcow"}
    assert job["progress"]["scopes"][0]["finished"] is True
    assert client.get("/api/jobs").json()[0]["outcome"] == job["outcome"]
    overview = client.get("/api/overview").json()
    assert overview["steps"]["verify"]["text"] == "3 items missing"
    for body in (json.dumps(job), json.dumps(overview), client.get("/api/jobs").body.decode()):
        for secret in SECRETS:
            assert secret not in body


def test_job_outcome_ignores_reports_older_than_the_job(client, monkeypatch):
    old = _write_report(client.state_dir, "20200101T000000Z.json",
                        _verify_report(10, 0, started="2020-01-01T00:00:00+00:00"), 1_000)
    recent_mtime = _write_report(client.state_dir, "20200101T000001Z.json",
                                 _verify_report(10, 0, started="2020-01-01T00:00:01+00:00"),
                                 time.time() + 60)  # touched later, but started long ago
    monkeypatch.setattr(cli, "main", lambda argv, **kw: 0)
    digest = _saved(client)
    job = wait_job(client, client.post("/api/jobs", {"command": "verify",
                                                     "selection_digest": digest}).json()["id"])
    assert job["exit_code"] == 0 and job["outcome"] is None
    assert old.exists() and recent_mtime.exists()


def test_job_outcome_of_another_command_is_not_taken(client, monkeypatch):
    def fake_main(argv=None, *, stdout=None, stderr=None):
        RunReport("plan", client.state_dir).write(0)
        return 0

    monkeypatch.setattr(cli, "main", fake_main)
    digest = _saved(client)
    job = wait_job(client, client.post("/api/jobs", {"command": "verify",
                                                     "selection_digest": digest}).json()["id"])
    assert job["outcome"] is None


def test_garbage_report_leaves_the_outcome_null(client, monkeypatch):
    def fake_main(argv=None, *, stdout=None, stderr=None):
        _write_report(client.state_dir, "garbage.json", "{not json", time.time())
        return 1

    monkeypatch.setattr(cli, "main", fake_main)
    digest = _saved(client)
    job = wait_job(client, client.post("/api/jobs", {"command": "migrate", "dry_run": True,
                                                     "selection_digest": digest}).json()["id"])
    assert job["exit_code"] == 1 and job["outcome"] is None and job["dry_run"] is True


def test_a_crashing_summary_never_stops_the_job_from_finishing(client, monkeypatch):
    def broken(report):
        raise RuntimeError(f"summary bug {API_KEY}")

    monkeypatch.setattr(cli, "main", _fake_report_main(client.state_dir, 10, code=0))
    monkeypatch.setattr(web, "summarize_report", broken)
    digest = _saved(client)
    job = wait_job(client, client.post("/api/jobs", {"command": "verify",
                                                     "selection_digest": digest}).json()["id"])
    assert job["exit_code"] == 0 and job["finished"] and job["outcome"] is None
    assert API_KEY not in json.dumps(job)


# -- review-gate follow-up: partial runs, cache, unreadable files, redaction -------------------

PARTIAL = {"only": "mail", "mailbox": ANNA, "mail_since": None}


def _with_scope(data: dict, scope: dict | None) -> dict:
    data = dict(data)
    if scope is None:
        data.pop("scope", None)
    else:
        data["scope"] = scope
    return data


def test_a_newer_clean_partial_verify_does_not_hide_a_failing_full_one(client):
    _write_report(client.state_dir, "1.json", _verify_report(7, 1), 1_000)  # full, failing
    _write_report(client.state_dir, "2.json", _with_scope(_verify_report(10, 0), PARTIAL),
                  2_000)  # newer, narrow, clean
    data = client.get("/api/overview").json()
    assert data["steps"]["verify"]["level"] == "warn"
    assert data["steps"]["verify"]["text"] == "3 items missing"
    assert data["latest"]["partial"] is True and data["latest"]["scope"] == "mail only · 1 mailbox"
    assert data["latest"]["headline"]["text"] == "Everything that was checked has arrived"


def test_a_newer_partial_verify_with_problems_decides_the_step(client):
    _write_report(client.state_dir, "1.json", _verify_report(10, 0), 1_000)  # full, clean
    _write_report(client.state_dir, "2.json", _with_scope(_verify_report(8, 1), PARTIAL), 2_000)
    step = client.get("/api/overview").json()["steps"]["verify"]
    assert step["level"] == "warn" and step["text"] == "2 items missing"


def test_only_partial_runs_fall_back_to_the_newest_one_in_grey(client):
    older = _with_scope(_verify_report(10, 0, started="2026-09-01T00:00:00+00:00"), PARTIAL)
    _write_report(client.state_dir, "1.json", older, 1_000)
    _write_report(client.state_dir, "2.json", _with_scope(_verify_report(10, 0), PARTIAL), 2_000)
    _write_report(client.state_dir, "3.json", _with_scope(_verify_report(10, 0), None), 500)
    step = client.get("/api/overview").json()["steps"]["verify"]
    assert step == {"started": "2026-10-01T01:00:00+00:00",
                    "finished": "2026-10-01T02:00:00+00:00", "exit_code": 0,
                    "level": "unknown", "text": "Partial check passed"}


def test_a_legacy_report_with_one_kind_counts_as_partial(client):
    _write_report(client.state_dir, "1.json", _with_scope(_verify_report(10, 0), None), 1_000)
    data = client.get("/api/overview").json()
    assert data["latest"]["partial"] is True and data["latest"]["scope"] == "mail only"
    assert data["steps"]["verify"]["text"] == "Partial check passed"


def test_an_unreadable_newer_report_marks_older_steps_and_latest(client):
    _write_report(client.state_dir, "1.json", _verify_report(7, 1), 1_000)
    _write_report(client.state_dir, "2.json", "{cut short", 2_000)
    data = client.get("/api/overview").json()
    assert data["steps"]["verify"] == {"started": "2026-10-01T01:00:00+00:00",
                                       "finished": "2026-10-01T02:00:00+00:00",
                                       "exit_code": 1, "level": "unknown",
                                       "text": "A newer report could not be read"}
    assert data["latest"]["headline"]["text"] == "The newest report could not be read"
    assert data["latest"]["headline"]["level"] == "unknown"
    assert client.get("/api/reports/latest").json()["command"] == "verify"  # still skips it


def _count_reads(monkeypatch) -> list[str]:
    reads: list[str] = []
    original = web._read_report_bytes

    def counting(path):
        reads.append(path.name)
        return original(path)

    monkeypatch.setattr(web, "_read_report_bytes", counting)
    return reads


def test_overview_caches_per_file_and_rereads_changes(client, monkeypatch):
    _write_report(client.state_dir, "1.json", _verify_report(7, 1), 1_000)
    path = _write_report(client.state_dir, "2.json", _migrate_report(3, False), 2_000)
    _write_report(client.state_dir, "3.json", "garbage", 500)  # unreadable results are cached
    reads = _count_reads(monkeypatch)
    first = client.get("/api/overview").json()
    assert sorted(reads) == ["1.json", "2.json", "3.json"]
    reads.clear()
    assert client.get("/api/overview").json() == first
    assert reads == []  # nothing changed: nothing read
    path.write_text(json.dumps(_migrate_report(4, False)), encoding="utf-8")
    os.utime(path, (2_001, 2_001))
    second = client.get("/api/overview").json()
    assert reads == ["2.json"] and second["steps"]["migrate"]["text"] == "4 items copied"
    assert second["latest"]["headline"]["text"] == "Migrate finished: 4 items copied"


def test_a_report_that_cannot_be_opened_is_retried_not_cached(client, monkeypatch):
    """Permissions can be fixed without the name, mtime or size changing (chmod only moves
    ctime), so a file that could not be opened must be tried again on the next overview."""
    _write_report(client.state_dir, "1.json", _verify_report(10, 0), 1_000)
    real = web._read_report_bytes

    def denied(path):
        raise PermissionError(13, "denied")

    monkeypatch.setattr(web, "_read_report_bytes", denied)
    first = client.get("/api/overview").json()
    assert first["latest"]["headline"]["text"] == "The newest report could not be read"
    assert first["steps"]["verify"] is None
    monkeypatch.setattr(web, "_read_report_bytes", real)
    second = client.get("/api/overview").json()
    assert second["steps"]["verify"]["text"] == "Everything arrived"
    assert second["latest"]["headline"]["level"] == "ok"


def test_overview_cache_is_pruned_to_the_current_files(client):
    app_paths = [_write_report(client.state_dir, f"{i}.json", _verify_report(10, 0), 1_000 + i)
                 for i in range(3)]
    client.get("/api/overview")
    app = _app_of(client)
    assert {key[0] for key in app._report_cache} == {"0.json", "1.json", "2.json"}
    app_paths[2].unlink()
    client.get("/api/overview")
    assert {key[0] for key in app._report_cache} == {"0.json", "1.json"}
    assert app._latest_summary[0][0] == "1.json"


def _app_of(client) -> web.WebApp:
    return client.server.app


def test_oversize_report_is_skipped_without_being_read(client, monkeypatch):
    monkeypatch.setattr(web, "MAX_REPORT_BYTES", 100)
    _write_report(client.state_dir, "1.json", _verify_report(10, 0), 1_000)  # > 100 bytes
    assert client.get("/api/reports/latest").status == 404
    data = client.get("/api/overview").json()
    assert data["latest"]["headline"]["text"] == "The newest report could not be read"
    assert data["steps"]["verify"] is None


def test_report_reads_never_follow_a_symlink(client, tmp_path):
    secret = tmp_path / "elsewhere.json"
    secret.write_text(json.dumps(_verify_report(10, 0)), encoding="utf-8")
    link = client.state_dir / "reports" / "x.json"
    link.parent.mkdir(parents=True, exist_ok=True)
    link.symlink_to(secret)
    with pytest.raises(OSError):  # the lstat/read race: a link swapped in after lstat
        web._read_report_bytes(link)
    fifo = client.state_dir / "reports" / "fifo.json"
    os.mkfifo(fifo)
    with pytest.raises(OSError):  # not a regular file on the descriptor, and no hang
        web._read_report_bytes(fifo)
    assert client.get("/api/reports/latest").status == 404


def test_a_job_without_a_report_never_inherits_the_previous_one(client, monkeypatch):
    monkeypatch.setattr(cli, "main", _fake_report_main(client.state_dir, 7))
    digest = _saved(client)
    first = wait_job(client, client.post("/api/jobs", {"command": "verify",
                                                       "selection_digest": digest}).json()["id"])
    assert first["outcome"]["level"] == "warn"
    monkeypatch.setattr(cli, "main", lambda argv, **kw: 2)  # e.g. the state lock was held
    second = wait_job(client, client.post("/api/jobs", {"command": "verify",
                                                        "selection_digest": digest}).json()["id"])
    assert second["exit_code"] == 2 and second["outcome"] is None


def test_secrets_in_reports_never_leave_through_overview_or_outcome(client, monkeypatch):
    leaky = _verify_report(7, 1)
    leaky["mailboxes"][ANNA]["errors"] = [f"boom {CLIENT_SECRET} {API_KEY} {TOKEN}"]
    leaky["mailboxes"][ANNA]["status"] = "failed"
    _write_report(client.state_dir, "1.json", {"command": f"x{API_KEY}", "exit_code": 1,
                                               "mailboxes": {}}, 1_000)
    _write_report(client.state_dir, "2.json", leaky, 2_000)  # newest: its errors are shown
    body = client.get("/api/overview").body.decode()
    assert "boom *** *** ***" in body
    for secret in SECRETS:
        assert secret not in body

    def leaking_summary(report):
        return {"headline": {"level": "bad", "text": f"oops {CLIENT_SECRET}"}}

    monkeypatch.setattr(cli, "main", _fake_report_main(client.state_dir, 10, code=0))
    monkeypatch.setattr(web, "summarize_report", leaking_summary)
    digest = _saved(client)
    job = wait_job(client, client.post("/api/jobs", {"command": "verify",
                                                     "selection_digest": digest}).json()["id"])
    assert job["outcome"] == {"level": "bad", "text": "oops ***"}
