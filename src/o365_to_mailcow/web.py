"""Local web UI for o365mig (``o365mig web``, ISC-155..166).

A standard-library HTTP server with one embedded page and a small JSON API: list the
tenant's mailboxes, check which exist in mailcow, save the mailbox list, start commands,
watch their output and read the latest report. It is meant for the operator's own machine.

Security model
--------------
* **Bind.** 127.0.0.1 by default. Any other address logs a warning: the token is then the
  only protection and the traffic is plain HTTP, so anything beyond localhost belongs
  behind your own TLS reverse proxy with its own authentication.
* **Token.** Every ``/api/*`` request needs ``Authorization: Bearer <token>``, compared in
  constant time; without it the answer is 401 before any routing, so an unauthenticated
  client cannot even learn which API paths exist. The token is random per start
  (``secrets.token_urlsafe(32)``) unless ``O365MIG_WEB_TOKEN`` sets it, and reaches the
  browser in the URL *fragment*, which browsers never send to a server.
* **Browser isolation.** No CORS headers: other origins can neither read responses nor send
  the ``Authorization`` header (their preflight gets 405). The page gets a fresh CSP nonce
  per response (``default-src 'none'``, nonce-only script and style, ``connect-src
  'self'``); every response carries ``nosniff``, ``no-referrer``, ``no-store`` and
  ``X-Frame-Options: DENY``.
* **Surface.** The page is the only file served. Everything else is a fixed route table;
  unknown paths are 404, methods other than GET/POST/PUT 405, bodies over 1 MiB 413. The
  only file the server writes is ``<state_dir>/mailboxes.csv``.
* **Secrets.** Responses never carry the client secret, the mailcow API key or the token:
  ``/api/status`` lists only non-secret settings, and upstream error messages and job
  output pass through a redactor that knows all three (on top of the ``SecretFilter``
  inside ``cli.main``).
* **Jobs.** One at a time, in a background thread, through ``cli.main`` with the server's
  config file and the saved selection, so the page can do nothing the CLI cannot.
  ``cli.main`` takes the state-directory lock itself; the server never holds it.
"""

from __future__ import annotations

import contextlib
import csv
import hmac
import io
import ipaddress
import itertools
import json
import logging
import os
import re
import secrets
import socket
import socketserver
import stat
import tempfile
import threading
import time
from collections import deque
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from datetime import UTC, datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, TextIO
from urllib.parse import parse_qs

from . import __version__, cli
from . import config as config_mod
from .auth import AuthError, TokenProvider
from .config import Config, ConfigError, _read_mailboxes_csv
from .graph import GraphClient, GraphError
from .mailcow import MailcowApi, MailcowError
from .report import clean

log = logging.getLogger(__name__)

ENV_TOKEN = "O365MIG_WEB_TOKEN"  # noqa: S105 - name of a variable, not a secret
MIN_TOKEN_CHARS = 16
MAX_BODY = 1024 * 1024
MAX_DRAIN = 4 * MAX_BODY  # a refused body up to this size is read and dropped, so the
#                           client gets the 413 instead of a connection reset
TENANT_CACHE_SECONDS = 600.0
MAX_CHECK_ADDRESSES = 500
MAX_SELECTION_ROWS = 10_000
MAX_NAME_CHARS = 200
MAX_QUOTA_MIB = 1_000_000
MAX_SAMPLE = 10_000
MAX_OUTPUT_LINES = 50_000
MAX_LINE_CHARS = 4_000
MAX_PARTIAL_CHARS = 65_536
TAIL_LINES = 200
MAX_JOBS_KEPT = 20
JOB_COMMANDS = ("plan", "provision", "migrate", "verify", "cleanup")
JOB_OPTIONS = frozenset({"command", "dry_run", "only", "mailbox", "sample"})
ALLOWED_METHODS = "GET, POST, PUT"
NONCE_PLACEHOLDER = "__CSP_NONCE__"
SELECTION_FILE = "mailboxes.csv"
SELECTION_HEADER = "# source,destination,name,quota_mib"
USER_FIELDS = ("id,displayName,mail,userPrincipalName,userType,accountEnabled,"
               "assignedLicenses,proxyAddresses")
PERMISSION_HINT = ("listing tenant users needs the Microsoft Graph permission User.Read.All "
                   "(an application permission for auth_mode \"app\", a delegated one for "
                   "\"delegated\") with admin consent")
API_CSP = "default-src 'none'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'"
SECURITY_HEADERS: tuple[tuple[str, str], ...] = (
    ("X-Content-Type-Options", "nosniff"),
    ("Referrer-Policy", "no-referrer"),
    ("Cache-Control", "no-store"),
    ("X-Frame-Options", "DENY"),
    ("Cross-Origin-Resource-Policy", "same-origin"),
)

# RFC 5322 atext local part and an LDH domain with at least one dot. The local part may not
# start with '#' (a comment in the CSV), '-' (an option on the command line), '.' or '/'.
_LOCAL = r"(?![#./-])[a-z0-9!#$%&'*+/=?^_`{|}~.-]{1,64}"
_LABEL = r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?"
_ADDRESS = re.compile(_LOCAL + r"@(?:" + _LABEL + r"\.)+" + _LABEL)
_JOB_PATH = re.compile(r"/api/jobs/([A-Za-z0-9_-]{1,64})(/output)?")


def page_csp(nonce: str) -> str:
    return (f"default-src 'none'; script-src 'nonce-{nonce}'; style-src 'nonce-{nonce}'; "
            "connect-src 'self'; img-src 'self' data:; base-uri 'none'; form-action 'none'; "
            "frame-ancestors 'none'")


def is_loopback(host: str) -> bool:
    if host.lower() == "localhost":
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


def base_url(host: str, port: int) -> str:
    return f"http://[{host}]:{port}" if ":" in host else f"http://{host}:{port}"


def _now() -> str:
    return datetime.now(UTC).isoformat(timespec="seconds")


# -- responses and input validation ------------------------------------------------------

class HttpError(Exception):
    """Ends the request with ``status`` and the body ``{"error": message}``."""

    def __init__(self, status: int, message: str,
                 headers: tuple[tuple[str, str], ...] = ()) -> None:
        super().__init__(message)
        self.status = status
        self.message = message
        self.headers = headers


@dataclass(frozen=True)
class Response:
    status: int
    body: bytes
    content_type: str = "application/json; charset=utf-8"
    csp: str = API_CSP
    headers: tuple[tuple[str, str], ...] = ()


def json_response(status: int, payload: object,
                  headers: tuple[tuple[str, str], ...] = ()) -> Response:
    return Response(status, json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                    headers=headers)


def _allow(method: str, *allowed: str) -> None:
    if method not in allowed:
        raise HttpError(405, "method not allowed", (("Allow", ", ".join(allowed)),))


def _parse_json(body: bytes) -> Any:
    if not body:
        raise HttpError(400, "the request body must be JSON")
    try:
        return json.loads(body.decode("utf-8"))
    except (ValueError, RecursionError) as exc:  # includes UnicodeDecodeError
        raise HttpError(400, f"invalid JSON: {exc.__class__.__name__}") from exc


def normalise_address(value: object, what: str) -> str:
    """Lower-case, trimmed e-mail address, or 400."""
    if not isinstance(value, str):
        raise HttpError(400, f"{what} must be a string")
    address = value.strip().lower()
    if len(address) > 254 or not _ADDRESS.fullmatch(address):
        raise HttpError(400, f"{what} is not a valid e-mail address: {clean(value)[:100]!r}")
    return address


def _selection_row(row: object, n: int) -> dict[str, Any]:
    if not isinstance(row, dict):
        raise HttpError(400, f"row {n} must be an object")
    source = normalise_address(row.get("source"), f"row {n}: source")
    destination = normalise_address(row.get("destination"), f"row {n}: destination")
    raw_name = row.get("name")
    if raw_name is None:
        raw_name = ""
    if not isinstance(raw_name, str):
        raise HttpError(400, f"row {n}: name must be a string")
    name = clean(raw_name).strip()
    if len(name) > MAX_NAME_CHARS:
        raise HttpError(400, f"row {n}: name must be at most {MAX_NAME_CHARS} characters")
    quota = row.get("quota_mib")
    if quota is not None and (isinstance(quota, bool) or not isinstance(quota, int)
                              or not 1 <= quota <= MAX_QUOTA_MIB):
        raise HttpError(400, f"row {n}: quota_mib must be null or an integer from 1 to "
                             f"{MAX_QUOTA_MIB}")
    return {"source": source, "destination": destination, "name": name, "quota_mib": quota}


def tenant_rows(users: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    """Graph ``/users`` entries -> mailbox rows, sorted by address.

    Kept: entries with a ``mail`` address that are not guests. ``kind`` is ``user`` when a
    licence is assigned, else ``shared`` (shared, room and unlicensed mailboxes).
    ``aliases`` are the secondary ``smtp:`` proxy addresses; the primary is ``mail``.
    """
    rows: list[dict[str, Any]] = []
    for user in users:
        mail = user.get("mail")
        if not isinstance(mail, str) or not mail.strip() or user.get("userType") == "Guest":
            continue
        primary = mail.strip().lower()
        proxies = user.get("proxyAddresses") or []
        aliases = {p[5:].strip().lower() for p in proxies
                   if isinstance(p, str) and p.startswith("smtp:")} - {primary}
        rows.append({
            "id": str(user.get("id") or ""),
            "display_name": clean(user.get("displayName") or ""),
            "mail": clean(primary),
            "upn": clean(user.get("userPrincipalName") or ""),
            "kind": "user" if user.get("assignedLicenses") else "shared",
            "enabled": user.get("accountEnabled") is True,
            "aliases": sorted(clean(a) for a in aliases),
        })
    rows.sort(key=lambda r: r["mail"])
    return rows


# -- jobs ------------------------------------------------------------------------------

class JobOutput:
    """Stands in for a job's stdout and stderr (``print``, log handlers, progress).

    Thread-safe, capped at ``max_lines`` (the oldest lines are dropped and counted), and
    every line is redacted when it completes, so a secret split over two writes is still
    caught. An unfinished line is never shown, for the same reason.
    """

    def __init__(self, redact: Callable[[str], str], max_lines: int = MAX_OUTPUT_LINES) -> None:
        self._redact = redact
        self._lines: deque[str] = deque(maxlen=max_lines)
        self._dropped = 0
        self._partial = ""
        self._lock = threading.Lock()

    def write(self, text: str) -> int:
        with self._lock:
            *complete, self._partial = (self._partial + str(text)).split("\n")
            for line in complete:
                self._store(line)
            if len(self._partial) > MAX_PARTIAL_CHARS:  # a runaway line without newline
                self._store(self._partial)
                self._partial = ""
        return len(text)

    def flush(self) -> None:
        """Nothing is buffered outside the line store."""

    def isatty(self) -> bool:
        return False

    def close(self) -> None:
        """Store the last line even without a trailing newline."""
        with self._lock:
            if self._partial:
                self._store(self._partial)
                self._partial = ""

    def _store(self, line: str) -> None:  # caller holds the lock
        line = self._redact(line.rstrip("\r"))  # redact first: truncation must not split
        if len(line) > MAX_LINE_CHARS:
            line = line[:MAX_LINE_CHARS] + " [line truncated]"
        if len(self._lines) == self._lines.maxlen:
            self._dropped += 1
        self._lines.append(line)

    def tail(self, n: int = TAIL_LINES) -> list[str]:
        with self._lock:
            return list(itertools.islice(reversed(self._lines), max(n, 0)))[::-1]

    def text(self) -> str:
        with self._lock:
            lines = list(self._lines)
            dropped = self._dropped
        if dropped:
            lines.insert(0, f"[{dropped} earlier line(s) dropped]")
        return "".join(f"{line}\n" for line in lines)


@dataclass
class Job:
    id: str
    seq: int
    command: str
    args: list[str]
    output: JobOutput
    started: str = field(default_factory=_now)
    finished: str | None = None
    exit_code: int | None = None

    @property
    def running(self) -> bool:
        return self.finished is None

    def summary(self) -> dict[str, Any]:
        return {"id": self.id, "command": self.command, "args": list(self.args),
                "started": self.started, "finished": self.finished,
                "exit_code": self.exit_code}


# -- application -----------------------------------------------------------------------

def _toml_string(value: str) -> str:
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


def _toml_settings(ms: dict[str, str], mc: dict[str, str]) -> str:
    lines = ["# written by the web UI; wins over config.toml and the environment", "",
             "[microsoft]"]
    lines += [f"{k} = {_toml_string(str(v))}" for k, v in ms.items() if v]
    lines += ["", "[mailcow]"]
    lines += [f"{k} = {_toml_string(str(v))}" for k, v in mc.items() if v]
    return "\n".join(lines) + "\n"


class WebApp:
    """What the routes do. ``Handler`` only speaks HTTP and calls into this."""

    def __init__(self, cfg: Config, token: str, page_template: str,
                 config_path: str | None = None, out: TextIO | None = None) -> None:
        self.cfg = cfg
        self.token = token
        self.page_template = page_template
        self.config_path = config_path
        self.out = out  # where a delegated sign-in prompt for the tenant listing appears
        self.selection_path = cfg.state_dir / SELECTION_FILE
        self._lock = threading.Lock()  # guards jobs, the tenant cache and the token provider
        self._jobs: dict[str, Job] = {}
        self._seq = 0
        self._tenant: tuple[float, dict[str, Any]] | None = None
        self._tokens: TokenProvider | None = None
        self._secrets = cli.SecretFilter()
        for value in (cfg.client_secret, cfg.mailcow_api_key, token):
            self._secrets.add(value)

    def redact(self, text: str) -> str:
        return self._secrets.redact(text)

    def authorized(self, header: str | None) -> bool:
        scheme, _, value = (header or "").partition(" ")
        return scheme.lower() == "bearer" and hmac.compare_digest(
            value.strip().encode("utf-8"), self.token.encode("utf-8"))

    def _upstream_error(self, exc: Exception) -> str:
        text = f"{exc.__class__.__name__}: {self.redact(str(exc))[:200]}"
        if isinstance(exc, GraphError) and exc.status == 403:
            text += f" ({PERMISSION_HINT})"
        return text

    # -- page and status ---------------------------------------------------------------

    def page(self) -> Response:
        nonce = secrets.token_urlsafe(18)
        body = self.page_template.replace(NONCE_PLACEHOLDER, nonce).encode("utf-8")
        return Response(200, body, "text/html; charset=utf-8", page_csp(nonce))

    def status(self) -> dict[str, Any]:
        with self._lock:
            job = self._running()
            running = {"id": job.id, "command": job.command} if job else None
        return {"version": __version__, "auth_mode": self.cfg.auth_mode,
                "mailcow_host": self.cfg.mailcow_host, "state_dir": str(self.cfg.state_dir),
                "selection_path": str(self.selection_path), "running_job": running,
                "configured": self._configured()}

    # -- settings (saved to <state_dir>/settings.toml; secrets never returned) ----------

    def _configured(self) -> bool:
        cfg = self.cfg
        return bool(cfg.tenant_id and cfg.client_id and cfg.mailcow_host and cfg.mailcow_api_key
                    and (cfg.auth_mode == "delegated" or cfg.client_secret))

    def settings(self) -> dict[str, Any]:
        cfg = self.cfg
        return {
            "microsoft": {"tenant_id": cfg.tenant_id, "client_id": cfg.client_id,
                          "auth_mode": cfg.auth_mode,
                          "client_secret_set": bool(cfg.client_secret)},
            "mailcow": {"host": cfg.mailcow_host, "api_key_set": bool(cfg.mailcow_api_key)},
            "path": str(config_mod.settings_path(cfg.state_dir)),
            "configured": self._configured(),
        }

    def save_settings(self, body: object) -> dict[str, Any]:
        if not isinstance(body, dict):
            raise HttpError(400, "expected a JSON object")
        ms = body.get("microsoft") if isinstance(body.get("microsoft"), dict) else {}
        mc = body.get("mailcow") if isinstance(body.get("mailcow"), dict) else {}
        current = config_mod.read_settings(self.cfg.state_dir)
        out_ms = dict(current.get("microsoft", {}))
        out_mc = dict(current.get("mailcow", {}))

        def ident(value: object, what: str, limit: int = 200) -> str:
            if not isinstance(value, str) or not value.strip():
                raise HttpError(400, f"{what} is required")
            value = value.strip()
            if len(value) > limit or not all(0x21 <= ord(c) <= 0x7E for c in value):
                raise HttpError(400, f"{what} must be printable ASCII without spaces")
            return value

        out_ms["tenant_id"] = ident(ms.get("tenant_id"), "microsoft.tenant_id")
        out_ms["client_id"] = ident(ms.get("client_id"), "microsoft.client_id")
        auth_mode = str(ms.get("auth_mode", "app")).lower()
        if auth_mode not in ("app", "delegated"):
            raise HttpError(400, "microsoft.auth_mode must be 'app' or 'delegated'")
        out_ms["auth_mode"] = auth_mode
        host = ident(mc.get("host"), "mailcow.host", 253)
        if not config_mod._HOSTNAME.match(host):
            raise HttpError(400, "mailcow.host must be a bare hostname")
        out_mc["host"] = host.lower()
        # secrets: only replaced when a non-empty value is sent; never echoed back
        for section, key, sent in ((out_ms, "client_secret", ms.get("client_secret")),
                                   (out_mc, "api_key", mc.get("api_key"))):
            if isinstance(sent, str) and sent.strip():
                if len(sent) > 1000 or any(ord(c) < 0x20 for c in sent):
                    raise HttpError(400, f"{key} has an invalid value")
                section[key] = sent.strip()
        if auth_mode == "app" and not out_ms.get("client_secret") and not (
                os.environ.get(config_mod.ENV_CLIENT_SECRET) or self.cfg.client_secret):
            raise HttpError(400, "microsoft.client_secret is required for app-only sign-in")

        path = config_mod.settings_path(self.cfg.state_dir)
        path.parent.mkdir(parents=True, exist_ok=True)
        text = _toml_settings(out_ms, out_mc)
        tmp = path.with_name(path.name + ".tmp")
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(text)
        os.replace(tmp, path)
        self._reload()
        return self.settings()

    def _reload(self) -> None:
        """Re-read the effective configuration after settings changed."""
        cfg = config_mod.load_config(self.config_path, require_mailboxes=False,
                                     require_credentials=False)
        with self._lock:
            self.cfg = cfg
            self._tokens = None
            self._tenant = None
        for value in (cfg.client_secret, cfg.mailcow_api_key):
            self._secrets.add(value)

    def test_settings(self) -> dict[str, Any]:
        """Try a Microsoft sign-in, a Graph user listing and a mailcow API call."""
        result: dict[str, Any] = {}
        if not self._configured():
            raise HttpError(400, "save the settings first")
        if self.cfg.auth_mode == "delegated":
            result["microsoft"] = {"ok": None, "message": "delegated sign-in happens on the "
                                   "first job (device code in its output)"}
            result["graph_users"] = {"ok": None, "message": "checked after sign-in"}
        else:
            try:
                tokens = TokenProvider(self.cfg, out=self.out)
                tokens.get_token()
                result["microsoft"] = {"ok": True, "message": "signed in (client credentials)"}
                try:
                    next(iter(GraphClient(tokens).iter_pages(
                        "/users", params={"$select": "id", "$top": "1"})), None)
                    result["graph_users"] = {"ok": True, "message": "User.Read.All granted"}
                except (GraphError, OSError, ValueError) as exc:
                    result["graph_users"] = {"ok": False, "message": self._upstream_error(exc)}
            except (AuthError, OSError, ValueError) as exc:
                result["microsoft"] = {"ok": False, "message": self._upstream_error(exc)}
                result["graph_users"] = {"ok": None, "message": "not checked"}
        try:
            api = MailcowApi(self.cfg.mailcow_host, self.cfg.mailcow_api_key,
                             verify=self.cfg.mailcow_ca_file or True)
            api.mailbox_exists("probe@example.invalid")
            result["mailcow"] = {"ok": True, "message": "API key accepted"}
        except (MailcowError, OSError, ValueError) as exc:
            result["mailcow"] = {"ok": False, "message": self._upstream_error(exc)}
        return result

    # -- tenant and mailcow ------------------------------------------------------------

    def _token_provider(self) -> TokenProvider:
        with self._lock:
            tokens = self._tokens
        if tokens is None:  # built outside the lock: MSAL may contact the authority
            tokens = TokenProvider(self.cfg, out=self.out)
            with self._lock:
                self._tokens = tokens = self._tokens or tokens
        return tokens

    def tenant_mailboxes(self, refresh: bool) -> dict[str, Any]:
        with self._lock:
            cached = self._tenant
        if cached and not refresh and time.monotonic() - cached[0] < TENANT_CACHE_SECONDS:
            return cached[1]
        try:
            graph = GraphClient(self._token_provider())
            users = list(graph.iter_pages("/users", params={"$select": USER_FIELDS,
                                                            "$top": "999"}))
        except (GraphError, AuthError, OSError, ValueError) as exc:  # OSError: requests
            raise HttpError(502, self._upstream_error(exc)) from exc
        payload = {"fetched_at": _now(), "mailboxes": tenant_rows(users)}
        with self._lock:
            self._tenant = (time.monotonic(), payload)
        return payload

    def mailcow_check(self, body: object) -> dict[str, Any]:
        addresses = body.get("addresses") if isinstance(body, dict) else None
        if not isinstance(addresses, list) or len(addresses) > MAX_CHECK_ADDRESSES:
            raise HttpError(400, f"addresses must be a list of at most {MAX_CHECK_ADDRESSES}")
        wanted: dict[str, str] = {}  # as sent -> normalised; all validated before any call
        for item in addresses:
            normalised = normalise_address(item, "address")
            wanted[item] = normalised
        api = MailcowApi(self.cfg.mailcow_host, self.cfg.mailcow_api_key,
                         verify=self.cfg.mailcow_ca_file or True)
        known: dict[str, bool] = {}
        try:
            for address in wanted.values():
                if address not in known:
                    known[address] = api.mailbox_exists(address)
        except MailcowError as exc:
            raise HttpError(502, self._upstream_error(exc)) from exc
        return {"exists": {sent: known[address] for sent, address in wanted.items()}}

    # -- selection ---------------------------------------------------------------------

    def selection(self) -> dict[str, Any]:
        path = self.selection_path
        if not path.exists():
            return {"path": str(path), "rows": []}
        try:
            mappings = _read_mailboxes_csv(path)
        except (ConfigError, OSError, ValueError, csv.Error) as exc:
            raise HttpError(500, f"cannot read {path}: {clean(exc)[:300]}") from exc
        rows = [{"source": m.source, "destination": m.destination, "name": m.name or "",
                 "quota_mib": m.quota_mib} for m in mappings]
        return {"path": str(path), "rows": rows}

    def save_selection(self, body: object) -> dict[str, Any]:
        rows = body.get("rows") if isinstance(body, dict) else None
        if not isinstance(rows, list) or len(rows) > MAX_SELECTION_ROWS:
            raise HttpError(400, f"rows must be a list of at most {MAX_SELECTION_ROWS} rows")
        parsed: list[dict[str, Any]] = []
        sources: set[str] = set()
        destinations: set[str] = set()
        for n, row in enumerate(rows, 1):
            entry = _selection_row(row, n)
            if entry["source"] in sources:
                raise HttpError(400, f"row {n}: duplicate source {entry['source']}")
            if entry["destination"] in destinations:
                raise HttpError(400, f"row {n}: duplicate destination {entry['destination']}")
            sources.add(entry["source"])
            destinations.add(entry["destination"])
            parsed.append(entry)
        self._write_selection(parsed)
        return {"path": str(self.selection_path), "rows": parsed}

    def _write_selection(self, rows: list[dict[str, Any]]) -> None:
        """Atomic replace of the CSV that ``--mailboxes`` reads, mode 0600."""
        buf = io.StringIO()
        buf.write(SELECTION_HEADER + "\n")
        writer = csv.writer(buf, lineterminator="\n")
        for r in rows:
            quota = "" if r["quota_mib"] is None else r["quota_mib"]
            writer.writerow([r["source"], r["destination"], r["name"], quota])
        path = self.selection_path
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            fd, tmp = tempfile.mkstemp(prefix=".mailboxes-", suffix=".tmp", dir=path.parent)
            try:
                with os.fdopen(fd, "w", encoding="utf-8", newline="") as fh:
                    fh.write(buf.getvalue())
                    fh.flush()
                    os.fsync(fh.fileno())
                os.chmod(tmp, 0o600)
                os.replace(tmp, path)
            except BaseException:
                with contextlib.suppress(OSError):
                    os.unlink(tmp)
                raise
        except OSError as exc:
            raise HttpError(500, f"cannot write {path}: {exc.strerror or exc}") from exc

    # -- jobs --------------------------------------------------------------------------

    def _job_argv(self, body: object) -> tuple[str, list[str]]:
        """Validated request -> ``cli.main`` arguments. Only these five options exist."""
        if not isinstance(body, dict):
            raise HttpError(400, "the body must be a JSON object")
        unknown = sorted(str(k) for k in set(body) - JOB_OPTIONS)
        if unknown:
            raise HttpError(400, f"unknown option(s): {clean(', '.join(unknown))[:200]}")
        command = body.get("command")
        if not isinstance(command, str) or command not in JOB_COMMANDS:
            raise HttpError(400, f"command must be one of {', '.join(JOB_COMMANDS)}")
        dry_run = body.get("dry_run", False)
        if not isinstance(dry_run, bool):
            raise HttpError(400, "dry_run must be true or false")
        only = body.get("only")
        if only is not None and (not isinstance(only, str) or only not in cli.KINDS):
            raise HttpError(400, f"only must be null or one of {', '.join(cli.KINDS)}")
        mailbox = body.get("mailbox")
        mailbox = normalise_address(mailbox, "mailbox") if mailbox not in (None, "") else None
        sample = body.get("sample", 0)
        sample = 0 if sample is None else sample
        if isinstance(sample, bool) or not isinstance(sample, int) or not 0 <= sample <= MAX_SAMPLE:
            raise HttpError(400, f"sample must be an integer from 0 to {MAX_SAMPLE}")
        argv = ["--config", self.config_path] if self.config_path else []
        if self.selection_path.is_file():  # else the config file's own mailbox list applies
            argv += ["--mailboxes", str(self.selection_path)]
        argv.append(command)
        if dry_run:
            argv.append("--dry-run")
        if only:
            argv += ["--only", only]
        if mailbox:
            argv.append(f"--mailbox={mailbox}")  # '=' form: never parsed as an option
        if command == "verify" and sample:
            argv += ["--sample", str(sample)]
        return command, argv

    def _running(self) -> Job | None:  # caller holds the lock
        return next((job for job in self._jobs.values() if job.running), None)

    def running_job(self) -> Job | None:
        with self._lock:
            return self._running()

    def start_job(self, body: object) -> Job:
        command, argv = self._job_argv(body)
        with self._lock:
            running = self._running()
            if running is not None:
                raise HttpError(409, f"job {running.id} ({running.command}) is still running")
            self._seq += 1
            job = Job(secrets.token_hex(8), self._seq, command, argv, JobOutput(self.redact))
            self._jobs[job.id] = job
            finished = sorted((j for j in self._jobs.values() if not j.running),
                              key=lambda j: j.seq)
            for old in finished[:max(0, len(self._jobs) - MAX_JOBS_KEPT)]:
                del self._jobs[old.id]
            threading.Thread(target=self._run, args=(job,), name=f"job-{job.id}",
                             daemon=True).start()
        return job

    def _run(self, job: Job) -> None:
        code = 1
        try:
            code = cli.main(job.args, stdout=job.output, stderr=job.output)
        except SystemExit as exc:  # argparse refuses arguments by exiting
            code = exc.code if isinstance(exc.code, int) else (0 if exc.code is None else 1)
        except Exception as exc:  # a crash must still end the job record
            log.error("job %s (%s) crashed: %s", job.id, job.command, exc.__class__.__name__)
            job.output.write(f"job failed: {exc.__class__.__name__}: {exc}\n")
        finally:
            job.output.close()  # before `finished`: a finished job's output is complete
            with self._lock:
                job.exit_code = code
                job.finished = _now()

    def _job(self, job_id: str) -> Job:
        with self._lock:
            job = self._jobs.get(job_id)
        if job is None:
            raise HttpError(404, "not found")
        return job

    def list_jobs(self) -> list[dict[str, Any]]:
        with self._lock:
            return [j.summary() for j in sorted(self._jobs.values(), key=lambda j: -j.seq)]

    def job_detail(self, job_id: str) -> dict[str, Any]:
        job = self._job(job_id)
        with self._lock:  # state first, then output: "not running" implies complete output
            detail = {**job.summary(), "running": job.running}
        return {**detail, "output_tail": job.output.tail(TAIL_LINES)}

    def job_output(self, job_id: str) -> str:
        return self._job(job_id).output.text()

    # -- reports -----------------------------------------------------------------------

    def latest_report(self, command: str | None) -> Any:
        """Newest ``<state_dir>/reports/*.json`` (regular files only), optionally the newest
        of one command (``?command=verify``)."""
        if command is not None and command not in JOB_COMMANDS:
            raise HttpError(400, f"command must be one of {', '.join(JOB_COMMANDS)}")
        candidates: list[tuple[float, str, Path]] = []
        for path in (self.cfg.state_dir / "reports").glob("*.json"):
            try:
                st = path.lstat()
            except OSError:
                continue
            if stat.S_ISREG(st.st_mode):
                candidates.append((st.st_mtime, path.name, path))
        for _, _, path in sorted(candidates, reverse=True):
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, ValueError) as exc:  # e.g. cut short by a crash
                log.warning("skipping unreadable report %s: %s", path.name,
                            exc.__class__.__name__)
                continue
            if command is None or (isinstance(data, dict) and data.get("command") == command):
                return data
        raise HttpError(404, "no report found")


# -- HTTP ------------------------------------------------------------------------------

class Handler(BaseHTTPRequestHandler):
    """HTTP plumbing: body limits, the token check, routing and headers.

    Order per request: size gate (413) -> read the body -> page (``/``, no token: it holds
    no data) -> token check for ``/api/*`` (401) -> route table (404/405) -> ``WebApp``.
    Every response, including the stdlib's own error paths, leaves through ``_send`` and
    so carries the security headers.
    """

    server: WebServer
    timeout = 60  # per socket operation: a stalled client cannot hold a thread forever

    def version_string(self) -> str:
        return "o365mig-web"

    def do_GET(self) -> None:
        self._dispatch()

    def do_POST(self) -> None:
        self._dispatch()

    def do_PUT(self) -> None:
        self._dispatch()

    def _dispatch(self) -> None:
        try:
            response = self._respond()
        except HttpError as exc:
            response = json_response(exc.status, {"error": exc.message}, exc.headers)
        except (TimeoutError, ConnectionError):  # the client went away mid-request
            self.close_connection = True
            return
        except Exception:
            log.error("internal error on %s %s", self.command, self.path.partition("?")[0],
                      exc_info=True)
            response = json_response(500, {"error": "internal server error"})
        self._send(response)

    def _respond(self) -> Response:
        app = self.server.app
        path, _, query = self.path.partition("?")
        path = path.partition("#")[0]
        body = self._read_body()
        if path == "/":
            _allow(self.command, "GET")
            return app.page()
        if not path.startswith("/api/"):
            raise HttpError(404, "not found")
        if not app.authorized(self.headers.get("Authorization")):
            raise HttpError(401, "unauthorized")
        return self._api(path, parse_qs(query), body)

    def _api(self, path: str, query: dict[str, list[str]], body: bytes) -> Response:
        app, method = self.server.app, self.command
        if path == "/api/status":
            _allow(method, "GET")
            return json_response(200, app.status())
        if path == "/api/settings":
            _allow(method, "GET", "PUT")
            if method == "GET":
                return json_response(200, app.settings())
            return json_response(200, app.save_settings(_parse_json(body)))
        if path == "/api/settings/test":
            _allow(method, "POST")
            return json_response(200, app.test_settings())
        if path == "/api/tenant/mailboxes":
            _allow(method, "GET")
            return json_response(200, app.tenant_mailboxes(query.get("refresh") == ["1"]))
        if path == "/api/mailcow/check":
            _allow(method, "POST")
            return json_response(200, app.mailcow_check(_parse_json(body)))
        if path == "/api/selection":
            _allow(method, "GET", "PUT")
            if method == "GET":
                return json_response(200, app.selection())
            return json_response(200, app.save_selection(_parse_json(body)))
        if path == "/api/jobs":
            _allow(method, "GET", "POST")
            if method == "GET":
                return json_response(200, app.list_jobs())
            return json_response(202, {"id": app.start_job(_parse_json(body)).id})
        if path == "/api/reports/latest":
            _allow(method, "GET")
            command = (query.get("command") or [None])[-1]
            return json_response(200, app.latest_report(command))
        match = _JOB_PATH.fullmatch(path)
        if match:
            _allow(method, "GET")
            job_id, output = match.groups()
            if output:
                return Response(200, app.job_output(job_id).encode("utf-8"),
                                "text/plain; charset=utf-8")
            return json_response(200, app.job_detail(job_id))
        raise HttpError(404, "not found")

    def _read_body(self) -> bytes:
        if self.headers.get("Transfer-Encoding"):
            raise HttpError(411, "send the request body with a Content-Length")
        raw = (self.headers.get("Content-Length") or "").strip()
        if not raw:
            return b""
        if not (raw.isascii() and raw.isdigit()):
            raise HttpError(400, "invalid Content-Length")
        length = int(raw)
        if length > MAX_BODY:
            self.close_connection = True
            if length <= MAX_DRAIN:
                remaining = length
                while remaining > 0 and (chunk := self.rfile.read(min(remaining, 65536))):
                    remaining -= len(chunk)
            raise HttpError(413, "request body too large (limit 1 MiB)")
        body = self.rfile.read(length)
        if len(body) != length:
            raise HttpError(400, "incomplete request body")
        return body

    def _send(self, response: Response) -> None:
        self.send_response(response.status)
        self.send_header("Content-Type", response.content_type)
        self.send_header("Content-Length", str(len(response.body)))
        self.send_header("Content-Security-Policy", response.csp)
        for name, value in SECURITY_HEADERS + response.headers:
            self.send_header(name, value)
        self.end_headers()
        if self.command != "HEAD":
            try:
                self.wfile.write(response.body)
            except (BrokenPipeError, ConnectionResetError):
                self.close_connection = True

    def send_error(self, code: int, message: str | None = None,
                   explain: str | None = None) -> None:
        """The stdlib's own errors (unknown method, malformed request line, oversized
        headers) answer in JSON with the same headers as everything else."""
        if code == HTTPStatus.NOT_IMPLEMENTED:  # no do_<METHOD>: only GET, POST, PUT exist
            response = json_response(405, {"error": "method not allowed"},
                                     (("Allow", ALLOWED_METHODS),))
        else:
            text = message or HTTPStatus(code).phrase
            response = json_response(code, {"error": clean(text)[:200]})
        self.close_connection = True
        self._send(response)

    def log_request(self, code: int | str = "-", size: int | str = "-") -> None:
        if isinstance(code, int) and code < 400:
            return  # the page polls job status; successes would drown everything else
        super().log_request(code, size)

    def log_message(self, fmt: str, *args: Any) -> None:
        log.info("%s %s", self.address_string(), clean(self.server.app.redact(fmt % args)))


class WebServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, address: tuple[str, int], app: WebApp) -> None:
        self.app = app
        super().__init__(address, Handler)

    def server_bind(self) -> None:
        # HTTPServer.server_bind resolves the host's FQDN, which can stall without DNS;
        # the name is never used here.
        socketserver.TCPServer.server_bind(self)
        self.server_name, self.server_port = str(self.server_address[0]), self.server_address[1]


class WebServer6(WebServer):
    address_family = socket.AF_INET6


def make_server(cfg: Config, bind: str, port: int, token: str, page_path: Path, *,
                config_path: str | None = None, out: TextIO | None = None) -> WebServer:
    """Read the page and bind (port 0 picks a free port); ``serve_forever`` runs it.
    ``config_path`` is handed to every job; ``out`` receives a delegated sign-in prompt."""
    template = Path(page_path).read_text(encoding="utf-8")
    app = WebApp(cfg, token, template, config_path=config_path, out=out)
    server_class = WebServer6 if ":" in bind else WebServer
    return server_class((bind, port), app)


def serve(cfg: Config, bind: str, port: int, token: str, page_path: Path, *,
          config_path: str | None = None, out: TextIO | None = None) -> None:
    """Serve until Ctrl-C. A job still running then stops with the process; like any
    interrupted run it resumes when started again, and a temporary app password it leaves
    behind is removed by the next run for that mailbox or by ``cleanup``."""
    server = make_server(cfg, bind, port, token, page_path, config_path=config_path, out=out)
    try:
        with contextlib.suppress(KeyboardInterrupt):
            server.serve_forever()
    finally:
        server.server_close()
    job = server.app.running_job()
    if job is not None:
        log.warning("job %s (%s) was still running and has been stopped; start it again to "
                    "resume, and run cleanup if app passwords were left behind",
                    job.id, job.command)
