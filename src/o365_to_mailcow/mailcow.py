"""mailcow API client, restricted to the four endpoints this tool needs.

The API key can create and delete mailboxes, so the client refuses any path outside an
explicit allowlist (ISC-95). Nothing that carries a secret (API key, app password, the
``log`` field of mailcow responses which echoes the request) is ever put into an
exception or a log line.
"""

from __future__ import annotations

import logging
import secrets
import string
from urllib.parse import quote

import requests

from . import USER_AGENT

log = logging.getLogger(__name__)

APP_PASSWORD_NAME = "o365-migration"  # noqa: S105 - a label prefix, not a secret
PROTOCOLS = ["imap_access", "dav_access"]
ALLOWED_PREFIXES = ("get/mailbox/", "get/app-passwd/", "add/app-passwd", "delete/app-passwd")
TIMEOUT = (10.0, 60.0)


class MailcowError(Exception):
    """An API call failed. The message holds the status and at most 200 chars of body."""

    def __init__(self, status: int, message: str) -> None:
        super().__init__(f"mailcow API HTTP {status}: {message[:200]}")
        self.status = status


def generate_password() -> str:
    """Return a 32-character password from ``secrets.token_urlsafe``.

    Regenerated until it contains lower, upper, digit and one of ``-_`` so that any
    mailcow password policy (length, mixed case, numbers, special characters) accepts it.
    """
    while True:
        pw = secrets.token_urlsafe(24)  # 24 bytes -> exactly 32 characters
        if (
            any(c in string.ascii_lowercase for c in pw)
            and any(c in string.ascii_uppercase for c in pw)
            and any(c in string.digits for c in pw)
            and any(c in "-_" for c in pw)
        ):
            return pw


def run_app_password_name() -> str:
    """A per-run name (``o365-migration-<8 hex>``) so parallel runs never pick each
    other's password out of the listing, and ``cleanup`` can still find every one of ours
    by prefix."""
    return f"{APP_PASSWORD_NAME}-{secrets.token_hex(4)}"


def is_our_app_password(entry: dict) -> bool:
    name = str(entry.get("name", ""))
    return name == APP_PASSWORD_NAME or name.startswith(APP_PASSWORD_NAME + "-")


class MailcowApi:
    """Thin client for ``https://{host}/api/v1/``."""

    def __init__(self, host: str, api_key: str, session: requests.Session | None = None,
                 verify: bool = True) -> None:
        self._base = f"https://{host}/api/v1/"
        self._session = session or requests.Session()
        self._session.trust_env = False  # no proxy/netrc surprises for the API key
        self._headers = {"X-API-Key": api_key, "User-Agent": USER_AGENT,
                         "Accept": "application/json"}
        self._verify = verify

    # -- transport ---------------------------------------------------------------------

    def _request(self, method: str, path: str, json_body: object | None = None) -> object:
        if not path.startswith(ALLOWED_PREFIXES):
            raise MailcowError(0, f"refusing non-allowlisted endpoint {path.split('/')[0]}/...")
        try:
            resp = self._session.request(
                method, self._base + path, headers=self._headers, json=json_body,
                timeout=TIMEOUT, verify=self._verify, allow_redirects=False,
            )
        except requests.RequestException as exc:
            raise MailcowError(0, f"network error: {exc.__class__.__name__}") from exc
        if 300 <= resp.status_code < 400:  # never follow: the key must stay on this host
            raise MailcowError(resp.status_code, "redirect refused")
        if resp.status_code >= 400:
            raise MailcowError(resp.status_code, resp.text[:200])
        try:
            return resp.json()
        except ValueError as exc:
            raise MailcowError(resp.status_code, f"invalid JSON: {resp.text[:200]}") from exc

    @staticmethod
    def _check_result(data: object, action: str) -> None:
        """mailcow write calls answer with ``[{type, msg, log}]``; anything else is failure."""
        items = data if isinstance(data, list) else [data]
        for item in items:
            if not isinstance(item, dict) or item.get("type") != "success":
                msg = item.get("msg") if isinstance(item, dict) else item
                # never include item["log"]: it echoes the request, including passwords
                raise MailcowError(200, f"{action} failed: {msg!s}")

    # -- endpoints ---------------------------------------------------------------------

    def mailbox_exists(self, address: str) -> bool:
        data = self._request("GET", f"get/mailbox/{quote(address, safe='@')}")
        if isinstance(data, dict):
            return str(data.get("username", "")).lower() == address.lower()
        if isinstance(data, list):
            return any(
                isinstance(d, dict) and str(d.get("username", "")).lower() == address.lower()
                for d in data
            )
        return False

    def list_app_passwords(self, address: str) -> list[dict]:
        """App passwords of ``address``; entries naming another mailbox are dropped."""
        data = self._request("GET", f"get/app-passwd/all/{quote(address, safe='@')}")
        if isinstance(data, dict) and "id" in data:
            data = [data]
        if not isinstance(data, list):
            return []
        out: list[dict] = []
        for d in data:
            if not isinstance(d, dict) or "id" not in d:
                continue
            owner = d.get("mailbox") or d.get("username")
            if owner is not None and str(owner).lower() != address.lower():
                continue
            out.append(d)
        return out

    def create_app_password(self, address: str, name: str | None = None) -> tuple[str, str]:
        """Create an IMAP+DAV app password. Returns ``(mailcow_id, password)``."""
        name = name or run_app_password_name()
        pw = generate_password()
        body = {
            "username": address,
            "app_name": name,
            "app_passwd": pw,
            "app_passwd2": pw,
            "active": "1",
            "protocols": list(PROTOCOLS),
        }
        self._check_result(self._request("POST", "add/app-passwd", body), "add/app-passwd")
        matches = [p for p in self.list_app_passwords(address) if p.get("name") == name]
        if not matches:
            raise MailcowError(200, "app password created but not found in listing")
        newest = max(matches, key=lambda p: _id_key(p["id"]))
        log.info("created app password id=%s for %s", newest["id"], address)
        return str(newest["id"]), pw

    def delete_app_password(self, mailcow_id: str) -> None:
        # mailcow's json_api.php decodes the raw body as the list of ids for every
        # delete endpoint (`$_POST['items'] = $request`), so the body is a bare array.
        data = self._request("POST", "delete/app-passwd", [str(mailcow_id)])
        self._check_result(data, "delete/app-passwd")
        log.info("deleted app password id=%s", mailcow_id)


def _id_key(value: object) -> tuple[int, str]:
    text = str(value)
    return (int(text), text) if text.isdigit() else (-1, text)
