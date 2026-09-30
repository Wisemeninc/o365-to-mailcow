"""mailcow API client: endpoints, payloads, error hygiene, endpoint allowlist."""

from __future__ import annotations

import json
import string

import pytest
import requests
import responses

from o365_to_mailcow.mailcow import (
    APP_PASSWORD_NAME,
    MailcowApi,
    MailcowError,
    generate_password,
)

BASE = "https://mail.example.net/api/v1/"
KEY = "api-key-value-456"
ALLOWED = ("get/mailbox/", "get/app-passwd/", "add/app-passwd", "delete/app-passwd")


@pytest.fixture
def api():
    return MailcowApi("mail.example.net", KEY)


def assert_only_allowed_endpoints():
    for call in responses.calls:
        path = call.request.url.removeprefix(BASE)
        assert call.request.url.startswith(BASE)
        assert path.startswith(ALLOWED), path  # ISC-95


@responses.activate
def test_mailbox_exists(api):
    responses.get(BASE + "get/mailbox/alice@example.net", json={"username": "alice@example.net"})
    responses.get(BASE + "get/mailbox/bob@example.net", json={})
    assert api.mailbox_exists("alice@example.net") is True
    assert api.mailbox_exists("bob@example.net") is False
    assert responses.calls[0].request.headers["X-API-Key"] == KEY
    assert responses.calls[0].request.headers["User-Agent"].startswith("o365-to-mailcow/")
    assert_only_allowed_endpoints()


@responses.activate
def test_list_app_passwords_handles_empty_object(api):
    responses.get(BASE + "get/app-passwd/all/a@x.net", json={})
    assert api.list_app_passwords("a@x.net") == []


@responses.activate
def test_create_app_password_payload_and_newest_id_isc_92(api):
    responses.post(BASE + "add/app-passwd",
                   json=[{"type": "success", "msg": ["app_passwd_added"], "log": []}])
    name = APP_PASSWORD_NAME + "-cafe0123"
    responses.get(BASE + "get/app-passwd/all/alice@example.net", json=[
        {"id": 3, "name": name}, {"id": 12, "name": name},
        {"id": 40, "name": "thunderbird"},
        {"id": 41, "name": name, "mailbox": "bob@example.net"},  # not ours: other mailbox
    ])
    mid, pw = api.create_app_password("alice@example.net", name=name)
    assert mid == "12"
    body = json.loads(responses.calls[0].request.body)
    assert body == {
        "username": "alice@example.net", "app_name": name, "app_passwd": pw,
        "app_passwd2": pw, "active": "1", "protocols": ["imap_access", "dav_access"],
    }
    assert len(pw) == 32


@responses.activate
def test_default_name_is_unique_per_run(api):
    responses.post(BASE + "add/app-passwd",
                   json=[{"type": "success", "msg": ["app_passwd_added"], "log": []}])
    seen: list[str] = []

    def listing(req):
        return 200, {}, json.dumps([{"id": len(seen) + 1, "name": seen[-1]}])

    def capture(req):
        seen.append(json.loads(req.body)["app_name"])
        return 200, {}, json.dumps([{"type": "success", "msg": ["ok"], "log": []}])

    responses.reset()
    responses.add_callback("POST", BASE + "add/app-passwd", callback=capture)
    responses.add_callback("GET", BASE + "get/app-passwd/all/alice@example.net",
                           callback=listing)
    api.create_app_password("alice@example.net")
    api.create_app_password("alice@example.net")
    assert len(set(seen)) == 2 and all(n.startswith(APP_PASSWORD_NAME + "-") for n in seen)
    assert_only_allowed_endpoints()


@responses.activate
def test_create_failure_never_leaks_password_or_log(api):
    responses.post(BASE + "add/app-passwd", json=[
        {"type": "danger", "msg": "password_complexity",
         "log": ["app_passwd", "add", {"app_passwd": "SECRET-ECHO"}]}])
    with pytest.raises(MailcowError) as exc:
        api.create_app_password("alice@example.net")
    assert "password_complexity" in str(exc.value)
    assert "SECRET-ECHO" not in str(exc.value)
    assert KEY not in str(exc.value)


@responses.activate
def test_delete_app_password_payload(api):
    responses.post(BASE + "delete/app-passwd", json=[{"type": "success", "msg": "ok"}])
    api.delete_app_password("12")
    assert json.loads(responses.calls[0].request.body) == {"items": ["12"]}


@responses.activate
def test_http_error_has_status_and_trimmed_body_without_key(api):
    responses.get(BASE + "get/mailbox/a@x.net", status=401, body="denied " + "x" * 500)
    with pytest.raises(MailcowError) as exc:
        api.mailbox_exists("a@x.net")
    assert exc.value.status == 401
    assert "denied" in str(exc.value) and KEY not in str(exc.value)
    assert len(str(exc.value)) < 260


def test_non_allowlisted_endpoint_refused_without_request(api):
    class Boom(requests.Session):
        def request(self, *a, **kw):
            raise AssertionError("no request may be sent")

    guarded = MailcowApi("mail.example.net", KEY, session=Boom())
    for path in ("add/mailbox", "delete/mailbox", "edit/app-passwd", "get/domain/all"):
        with pytest.raises(MailcowError, match="refusing"):
            guarded._request("POST", path, {})


def test_every_request_has_timeout_and_verify():
    seen: list[dict] = []

    class Capture(requests.Session):
        def request(self, method, url, **kw):
            seen.append(kw)
            resp = requests.Response()
            resp.status_code = 200
            resp._content = b"{}"
            return resp

    api = MailcowApi("mail.example.net", KEY, session=Capture())
    api.mailbox_exists("a@x.net")
    api.list_app_passwords("a@x.net")
    assert seen and all(kw["timeout"] and None not in kw["timeout"] for kw in seen)
    assert all(kw["verify"] is True for kw in seen)


def test_network_error_is_mailcow_error(api):
    class Down(requests.Session):
        def request(self, *a, **kw):
            raise requests.ConnectionError(f"cannot reach, key={KEY}")

    with pytest.raises(MailcowError) as exc:
        MailcowApi("mail.example.net", KEY, session=Down()).mailbox_exists("a@x.net")
    assert KEY not in str(exc.value)


def test_generate_password_meets_policy():
    for _ in range(50):
        pw = generate_password()
        assert len(pw) == 32
        assert set(pw) <= set(string.ascii_letters + string.digits + "-_")
        assert any(c.isdigit() for c in pw) and any(c in "-_" for c in pw)
        assert any(c.islower() for c in pw) and any(c.isupper() for c in pw)
