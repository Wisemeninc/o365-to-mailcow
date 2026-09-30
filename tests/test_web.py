"""Local web UI server (ISC-155..161, 163, 166, 167): contract tests over real HTTP.

The server runs on an ephemeral port in a thread; requests go through ``http.client``.
Graph and MSAL are faked, mailcow is recorded with ``responses``.
"""

from __future__ import annotations

import http.client
import io
import json
import os
import re
import stat
import threading
import time
from dataclasses import dataclass

import pytest
import responses

from fakes_o365 import FakeGraph, make_config
from o365_to_mailcow import __version__, cli, web
from o365_to_mailcow.auth import AuthError
from o365_to_mailcow.config import MailboxMapping, load_config
from o365_to_mailcow.graph import GraphError

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
                         ("PUT", "/api/jobs")):
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
    }
    for secret in SECRETS:
        assert secret not in r.body.decode()


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
    assert r.json() == {"path": str(client.selection), "rows": []}
    rows = [
        {"source": " Alice@Contoso.com ", "destination": "Alice@Example.NET",
         "name": "Alice\x07 Liddell", "quota_mib": 2048},
        {"source": "bob@contoso.com", "destination": "robert@example.net",
         "name": 'Smith, Bob "B"', "quota_mib": None},
        {"source": "info@contoso.com", "destination": "info@example.net", "name": None,
         "quota_mib": None},
    ]
    expected = [
        {"source": "alice@contoso.com", "destination": "alice@example.net",
         "name": "Alice Liddell", "quota_mib": 2048},
        {"source": "bob@contoso.com", "destination": "robert@example.net",
         "name": 'Smith, Bob "B"', "quota_mib": None},
        {"source": "info@contoso.com", "destination": "info@example.net", "name": "",
         "quota_mib": None},
    ]
    r = client.put("/api/selection", {"path": "/ignored", "rows": rows})
    assert r.status == 200, r.body
    assert r.json() == {"path": str(client.selection), "rows": expected}
    assert client.get("/api/selection").json()["rows"] == expected
    assert stat.S_IMODE(client.selection.stat().st_mode) == 0o600
    assert client.selection.read_text(encoding="utf-8").startswith("#")
    assert not [p for p in client.selection.parent.iterdir() if p.name.endswith(".tmp")]
    # exactly the file `--mailboxes` accepts
    cfg = load_config(str(client.conf), mailboxes_csv=str(client.selection))
    assert cfg.mailboxes == (
        MailboxMapping("alice@contoso.com", "alice@example.net", "Alice Liddell", 2048),
        MailboxMapping("bob@contoso.com", "robert@example.net", 'Smith, Bob "B"', None),
        MailboxMapping("info@contoso.com", "info@example.net", None, None),
    )
    # an empty selection is a valid selection
    assert client.put("/api/selection", {"rows": []}).json()["rows"] == []
    assert client.get("/api/selection").json()["rows"] == []


@pytest.mark.parametrize("rows, fragment", [
    ([GOOD, {**GOOD, "source": "b@contoso.com", "destination": "A@Example.net"}],
     "duplicate destination"),
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
    assert r.status == 500
    assert "quota" in r.json()["error"]


# -- jobs ------------------------------------------------------------------------------

def test_job_runs_cli_main_and_captures_output_isc_160_161(client, monkeypatch):
    seen: dict = {}

    def fake_main(argv=None, *, stdout=None, stderr=None):
        seen["argv"] = list(argv)
        print("plan output line", file=stdout)
        stderr.write(f"leak {API_KEY} {CLIENT_SECRET} {TOKEN}\n")
        print("no trailing newline", end="", file=stdout)
        return 0

    monkeypatch.setattr(cli, "main", fake_main)
    assert client.put("/api/selection", {"rows": [GOOD]}).status == 200
    r = client.post("/api/jobs", {"command": "plan", "dry_run": True, "only": None,
                                  "mailbox": None, "sample": 0})
    assert r.status == 202, r.body
    job_id = r.json()["id"]
    job = wait_job(client, job_id)
    assert seen["argv"] == ["--config", str(client.conf), "--mailboxes", str(client.selection),
                            "plan", "--dry-run"]
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
    assert set(listing[0]) == {"id", "command", "args", "started", "finished", "exit_code"}


def test_job_options_become_cli_arguments_isc_160(client, monkeypatch):
    argvs: list[list[str]] = []
    monkeypatch.setattr(cli, "main", lambda argv, **kw: argvs.append(list(argv)) or 1)
    r = client.post("/api/jobs", {"command": "verify", "dry_run": False, "only": "mail",
                                  "mailbox": "Alice@Example.net", "sample": 5})
    job = wait_job(client, r.json()["id"])
    assert job["exit_code"] == 1
    # no selection saved yet: the config's own mailbox list applies
    assert argvs[0] == ["--config", str(client.conf), "verify", "--only", "mail",
                        "--mailbox=alice@example.net", "--sample", "5"]
    r = client.post("/api/jobs", {"command": "migrate", "sample": 5})  # sample: verify only
    wait_job(client, r.json()["id"])
    assert argvs[1] == ["--config", str(client.conf), "migrate"]
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
        first = client.post("/api/jobs", {"command": "migrate"})
        assert first.status == 202
        job_id = first.json()["id"]
        assert entered.wait(5)
        second = client.post("/api/jobs", {"command": "plan"})
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
    third = client.post("/api/jobs", {"command": "plan"})
    assert third.status == 202
    wait_job(client, third.json()["id"])


def test_job_crash_and_argparse_exit_are_recorded(client, monkeypatch):
    def crashing_main(argv=None, *, stdout=None, stderr=None):
        raise RuntimeError(f"exploded with {API_KEY}")

    monkeypatch.setattr(cli, "main", crashing_main)
    job = wait_job(client, client.post("/api/jobs", {"command": "plan"}).json()["id"])
    assert job["exit_code"] == 1
    assert job["output_tail"] == ["job failed: RuntimeError: exploded with ***"]

    def exiting_main(argv=None, *, stdout=None, stderr=None):
        raise SystemExit(2)

    monkeypatch.setattr(cli, "main", exiting_main)
    job = wait_job(client, client.post("/api/jobs", {"command": "plan"}).json()["id"])
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
    opts = cli.Options("plan", None, None, None, None, False, False, 0, 0)
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
    assert f"web UI: http://127.0.0.1:8099/#token={TOKEN}\n" in err.getvalue()
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
