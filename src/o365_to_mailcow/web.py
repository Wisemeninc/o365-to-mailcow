"""Local web UI for o365mig (``o365mig web``, ISC-155..166).

A standard-library HTTP server with one embedded page and a small JSON API: list the
tenant's mailboxes, check which exist in mailcow, save the mailbox list, start commands,
watch their output (with live per-mailbox progress and, once done, the headline of the
report the job itself named in its ``report: <path>`` line) and read the reports, raw or summarised
for the page (``/api/overview``: the newest run of each step plus the latest report, judged by
``summary.py``). It is meant for the operator's own machine.

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
  server writes only ``<state_dir>/mailboxes.csv``, ``settings.toml`` and per-job selection
  snapshots; reports are only read (``/api/reports/latest``, ``/api/overview``, a finished
  job's outcome): opened with ``O_NOFOLLOW``, regular files of at most ``MAX_REPORT_BYTES``
  only, at most ``MAX_REPORTS_SCANNED`` files per overview or outcome lookup, and the
  overview caches what it learned per file (name, mtime, size) so polling re-reads nothing.
* **Secrets.** Responses never carry the client secret, the mailcow API key or the token:
  ``/api/status`` lists only non-secret settings, and upstream error messages, job output,
  job outcomes and every string of ``/api/overview`` pass through a redactor that knows all
  three (on top of the ``SecretFilter``
  inside ``cli.main``).
* **Jobs.** One at a time, in a background thread, through ``cli.main`` with the server's
  config file and the saved selection, so the page can do nothing the CLI cannot.
  ``cli.main`` takes the state-directory lock itself; the server never holds it.
"""

from __future__ import annotations

import contextlib
import csv
import hashlib
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
from .config import Config, ConfigError, _read_mailboxes_csv, valid_address
from .graph import GraphClient, GraphError
from .mailcow import MailcowApi, MailcowError
from .report import clean
from .summary import summarize_report, summarize_with_step, unreadable_summary

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
MAX_OUTPUT_BYTES = 8 * 1024 * 1024  # per job; oldest lines dropped beyond this
MAX_LINE_CHARS = 4_000
MAX_PARTIAL_CHARS = 65_536
TAIL_LINES = 200
MAX_JOBS_KEPT = 20
MAX_REPORTS_SCANNED = 200  # report files considered per overview / job-outcome lookup
MAX_REPORT_BYTES = 64 * 1024 * 1024  # a bigger report file is unreadable, never read
MAX_PROGRESS_LINE = 20_000  # longer lines are never parsed (the phase bound is below)
MAX_SCOPES = 6_000  # progress scopes tracked per job; new ones beyond this are ignored
MAX_SCOPES_SHOWN = 600
STEP_COMMANDS = ("provision", "migrate", "verify", "cleanup")
SCOPE_KINDS = ("mail", "calendar", "contacts")
JOB_COMMANDS = ("plan", "provision", "migrate", "verify", "cleanup")
_SNAPSHOT_PLACEHOLDER = "<selection-snapshot>"
JOB_OPTIONS = frozenset({"command", "dry_run", "only", "mailbox", "sample", "selection_digest",
                         "mail_since"})
_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
ALLOWED_METHODS = "GET, POST, PUT"
NONCE_PLACEHOLDER = "__CSP_NONCE__"
SELECTION_FILE = "mailboxes.csv"
SELECTION_HEADER = "# source,destination,name,quota_mib,aliases (semicolon-separated)"
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
# One line of report.Progress._line (plus " (finished)" from Progress.finish). Anchored, with
# bounded repetitions and ASCII digits only: it runs inside the migration's own print calls.
_PROGRESS_LINE = re.compile(
    r"\[(?P<mailbox>[^\s\[\]]{1,320}) (?P<kind>mail|calendar|contacts)\] "
    r"(?P<done>[0-9]{1,15})/(?P<total>[0-9]{1,15}) items, (?P<rate>[0-9]{1,15})/min"
    r"(?: \((?P<phase>.{1,16000})\))?")
# What cli.main prints to stderr in its `finally`: the job's own report, by path.
_REPORT_LINE = re.compile(r"report: (?P<path>[^\x00]{1,4096})")
_REPORT_NAME = re.compile(r"[0-9A-Za-z_-]{1,80}\.json")


def page_csp(nonce: str) -> str:
    return (f"default-src 'none'; script-src 'nonce-{nonce}'; style-src 'nonce-{nonce}'; "
            "connect-src 'self'; img-src 'self' data:; base-uri 'none'; form-action 'none'; "
            "frame-ancestors 'none'")


def host_only(value: str) -> str:
    """``host``, ``host:port``, ``[v6]`` or ``[v6]:port`` -> the bare host."""
    if value.startswith("["):
        return value[1:].partition("]")[0]
    return value.rsplit(":", 1)[0] if value.count(":") == 1 else value


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
    if len(address) > 254 or not _ADDRESS.fullmatch(address) or not valid_address(address):
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
    raw_aliases = row.get("aliases", [])
    if raw_aliases is None:
        raw_aliases = []
    if not isinstance(raw_aliases, list) or len(raw_aliases) > 100:
        raise HttpError(400, f"row {n}: aliases must be a list of at most 100 addresses")
    aliases = sorted({normalise_address(a, f"row {n}: alias") for a in raw_aliases})
    if destination in aliases:
        aliases.remove(destination)
    return {"source": source, "destination": destination, "name": name, "quota_mib": quota,
            "aliases": aliases}


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

    def __init__(self, redact: Callable[[str], str], max_lines: int = MAX_OUTPUT_LINES,
                 max_bytes: int = MAX_OUTPUT_BYTES) -> None:
        self._max_bytes = max_bytes
        self._bytes = 0
        self._redact = redact
        self._lines: deque[str] = deque(maxlen=max_lines)
        self._dropped = 0
        self._partial = ""
        self._scopes: dict[tuple[str, str], dict[str, Any]] = {}
        self._report_path: str | None = None  # from the job's last "report: <path>" line
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
        if line.startswith(("[", "report: ")) and len(line) <= MAX_PROGRESS_LINE:
            # parsed whole, before display truncation; never let it break the print() of the
            # migration that wrote it, and no logging here: the job's log handler writes
            # into this very object
            with contextlib.suppress(Exception):
                self._track(line)
        if len(line) > MAX_LINE_CHARS:
            line = line[:MAX_LINE_CHARS] + " [line truncated]"
        size = len(line.encode("utf-8"))
        if len(self._lines) == self._lines.maxlen:  # deque would drop the oldest silently
            self._bytes -= len(self._lines.popleft().encode("utf-8"))
            self._dropped += 1
        while self._lines and self._bytes + size > self._max_bytes:
            self._bytes -= len(self._lines.popleft().encode("utf-8"))
            self._dropped += 1
        self._lines.append(line)
        self._bytes += size

    def _track(self, line: str) -> None:  # caller holds the lock
        if line.startswith("report: "):
            match = _REPORT_LINE.fullmatch(line)
            if match is not None:
                self._report_path = match["path"]
            return
        self._track_progress(line)

    def _track_progress(self, line: str) -> None:  # caller holds the lock
        """Latest numbers per (mailbox, kind) from the tool's own progress lines."""
        match = _PROGRESS_LINE.fullmatch(line)
        if match is None:
            return
        key = (match["mailbox"], match["kind"])  # the whole address: the scope's identity
        current = self._scopes.get(key)
        if current is None and len(self._scopes) >= MAX_SCOPES:
            return
        phase = match["phase"] or ""
        finished = phase == "finished"
        if current is not None and current["finished"] and not finished:
            # Progress.maybe_print prints lines it built before releasing its lock, so a
            # stale unfinished line can arrive after the scope's "(finished)" line
            return
        self._scopes[key] = {
            "mailbox": clean(key[0])[:300], "kind": key[1], "done": int(match["done"]),
            "total": int(match["total"]), "rate": int(match["rate"]),
            "phase": "" if finished else clean(phase)[:300], "finished": finished}

    def report_path(self) -> str | None:
        """The path from the job's last ``report: <path>`` line (what ``cli.main`` prints
        when it wrote the run report), or ``None``."""
        with self._lock:
            return self._report_path

    def progress(self) -> dict[str, Any] | None:
        """Sums over all scopes plus the first ``MAX_SCOPES_SHOWN`` of them; ``None`` until
        the job printed its first progress line."""
        with self._lock:
            scopes = [dict(scope) for scope in self._scopes.values()]
        if not scopes:
            return None
        scopes.sort(key=lambda s: (s["mailbox"], SCOPE_KINDS.index(s["kind"])))
        return {"done": sum(s["done"] for s in scopes), "total": sum(s["total"] for s in scopes),
                "rate": sum(s["rate"] for s in scopes if not s["finished"]),
                # what the unfinished scopes still have to do: a finished scope that stopped
                # short has no rate, so its remainder must not count towards the time left
                "remaining": sum(max(s["total"] - s["done"], 0) for s in scopes
                                 if not s["finished"]),
                "scopes": scopes[:MAX_SCOPES_SHOWN],
                "more_scopes": max(len(scopes) - MAX_SCOPES_SHOWN, 0)}

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
    outcome: dict[str, str] | None = None  # headline of the report the job wrote, once done

    @property
    def running(self) -> bool:
        return self.finished is None

    @property
    def dry_run(self) -> bool:
        return "--dry-run" in self.args

    def summary(self) -> dict[str, Any]:
        return {"id": self.id, "command": self.command, "args": list(self.args),
                "started": self.started, "finished": self.finished,
                "exit_code": self.exit_code, "dry_run": self.dry_run,
                "outcome": dict(self.outcome) if self.outcome else None}


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
                 config_path: str | None = None, out: TextIO | None = None,
                 lock_settings: bool = False) -> None:
        self.cfg = cfg
        self.token = token
        self.lock_settings = lock_settings  # `web --lock-settings`: PUT /api/settings -> 403
        self._settings_lock = threading.Lock()
        self.page_template = page_template
        self.config_path = config_path
        self.out = out  # where a delegated sign-in prompt for the tenant listing appears
        self.selection_path = cfg.state_dir / SELECTION_FILE
        self._lock = threading.Lock()  # guards jobs, the tenant cache and the token provider
        self._jobs: dict[str, Job] = {}
        self._seq = 0
        self._report_cache_lock = threading.Lock()  # guards the two overview caches below
        self._report_cache: dict[tuple[str, int, int], _ReportMeta] = {}
        self._latest_summary: tuple[tuple[str, int, int], dict[str, Any]] | None = None
        self._tenant: tuple[float, dict[str, Any]] | None = None
        self._tokens: TokenProvider | None = None
        self._secrets = cli.SecretFilter()
        for value in (cfg.client_secret, cfg.mailcow_api_key, token):
            self._secrets.add(value)
        for stale in (cfg.state_dir / "jobs").glob("selection-*.csv"):  # earlier process
            with contextlib.suppress(OSError):
                stale.unlink()

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
        """Save connection settings. Rules that keep a token holder from redirecting a
        credential: a mailcow host is only ever stored together with an API key sent in
        the same request when it changes, and tenant/client ids only together with a
        client secret (app-only). Saved values never borrow a secret from the environment
        or the config file (see ``config.load_config``)."""
        if self.lock_settings:
            raise HttpError(403, "settings are locked (web --lock-settings)")
        if not isinstance(body, dict):
            raise HttpError(400, "expected a JSON object")
        ms = body.get("microsoft") if isinstance(body.get("microsoft"), dict) else {}
        mc = body.get("mailcow") if isinstance(body.get("mailcow"), dict) else {}

        def ident(value: object, what: str, limit: int = 200) -> str:
            if not isinstance(value, str) or not value.strip():
                raise HttpError(400, f"{what} is required")
            value = value.strip()
            if len(value) > limit or not all(0x21 <= ord(c) <= 0x7E for c in value):
                raise HttpError(400, f"{what} must be printable ASCII without spaces")
            return value

        def secret(value: object, what: str) -> str | None:
            if not isinstance(value, str) or not value.strip():
                return None
            if len(value) > 1000 or any(ord(c) < 0x20 or ord(c) == 0x7F for c in value):
                raise HttpError(400, f"{what} has an invalid value")
            try:
                value.encode("utf-8")  # lone surrogates cannot be written to the file
            except UnicodeEncodeError as exc:
                raise HttpError(400, f"{what} has an invalid value") from exc
            return value.strip()

        tenant_id = ident(ms.get("tenant_id"), "microsoft.tenant_id")
        client_id = ident(ms.get("client_id"), "microsoft.client_id")
        auth_mode = str(ms.get("auth_mode", "app")).lower()
        if auth_mode not in ("app", "delegated"):
            raise HttpError(400, "microsoft.auth_mode must be 'app' or 'delegated'")
        host = ident(mc.get("host"), "mailcow.host", 253).lower()
        if not config_mod.valid_hostname(host):
            raise HttpError(400, "mailcow.host must be a bare host name (no IP address)")
        client_secret = secret(ms.get("client_secret"), "client_secret")
        api_key = secret(mc.get("api_key"), "api_key")

        with self._settings_lock:
            try:
                current = config_mod.read_settings(self.cfg.state_dir)
            except ConfigError:
                current = {}  # a broken file is replaced, never a dead end
            cur_ms, cur_mc = current.get("microsoft", {}), current.get("mailcow", {})
            if self.running_job() is not None:
                # a running job holds a sign-in in memory and would rewrite its cache file
                raise HttpError(409, "a job is running; change the connection settings when "
                                     "it has finished")
            ids_changed = (tenant_id, client_id) != (cur_ms.get("tenant_id"),
                                                     cur_ms.get("client_id"))
            saved_host = cur_mc.get("host") or self.cfg.mailcow_host or ""
            host_changed = bool(saved_host) and host != saved_host
            mode_changed = auth_mode != cur_ms.get("auth_mode", auth_mode)
            # A change of the destination host, of the tenant/client ids or of the sign-in
            # mode re-pairs the *whole* connection: the saved client secret never carries
            # over, and every cached delegated sign-in is discarded, otherwise a token
            # holder could keep the operator's Graph credential and point the migration
            # at their own server.
            repair = host_changed or ids_changed or mode_changed
            if repair and auth_mode == "app" and client_secret is None:
                raise HttpError(400, "changing the mailcow host, the tenant/client ids or the "
                                     "sign-in mode requires entering the client secret again")
            if client_secret is None and auth_mode == "app" and not cur_ms.get("client_secret"):
                raise HttpError(400, "enter the client secret together with the tenant and "
                                     "client ids")
            if client_secret is None and not repair:
                client_secret = cur_ms.get("client_secret")
            if auth_mode != "app":
                client_secret = None  # delegated mode stores no secret, ever
            if api_key is None:
                if host_changed or ids_changed or not cur_mc.get("api_key"):
                    raise HttpError(400, "enter the mailcow API key together with the host "
                                         "(required again when the host or the tenant/client "
                                         "ids change)")
                api_key = cur_mc.get("api_key")
            if repair:
                self._discard_sign_ins()  # after every validation passed

            out_ms = {"tenant_id": tenant_id, "client_id": client_id, "auth_mode": auth_mode}
            if client_secret:
                out_ms["client_secret"] = client_secret
            out_mc = {"host": host, "api_key": api_key}
            path = config_mod.settings_path(self.cfg.state_dir)
            path.parent.mkdir(parents=True, exist_ok=True)
            fd, tmp = tempfile.mkstemp(prefix=".settings-", suffix=".tmp", dir=path.parent)
            try:
                os.fchmod(fd, 0o600)
                with os.fdopen(fd, "w", encoding="utf-8") as fh:
                    fh.write(_toml_settings(out_ms, out_mc))
                os.replace(tmp, path)
            except BaseException:
                with contextlib.suppress(OSError):
                    os.unlink(tmp)
                raise
            changed = [k for k, v in (("tenant_id", tenant_id), ("client_id", client_id),
                                      ("auth_mode", auth_mode), ("host", host))
                       if v != {**cur_ms, **cur_mc}.get(k)]
            if client_secret and client_secret != cur_ms.get("client_secret"):
                changed.append("client_secret")
            if api_key != cur_mc.get("api_key"):
                changed.append("api_key")
            log.warning("settings saved to %s; changed: %s; mailcow host now %s",
                        path, ", ".join(changed) or "nothing", host)
            self._reload()
        return self.settings()

    def _web_cache_path(self) -> Path:
        return self.cfg.state_dir / f"msal_cache_web_{config_mod.connection_id(self.cfg)}.bin"

    def _discard_sign_ins(self) -> None:
        """Delete every MSAL token cache in the state directory (web and job caches), so a
        delegated sign-in cannot outlive a change of destination or tenant."""
        for path in self.cfg.state_dir.glob("msal_cache*.bin"):
            with contextlib.suppress(OSError):
                path.unlink()
        with self._lock:
            self._tokens = None

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
                tokens = TokenProvider(self.cfg, cache_path=self._web_cache_path(), out=self.out)
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
            # own cache file: a job's TokenProvider writes msal_cache_<connection>.bin
            tokens = TokenProvider(self.cfg, cache_path=self._web_cache_path(), out=self.out)
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

    def _selection_digest(self) -> str | None:
        try:
            return hashlib.sha256(self.selection_path.read_bytes()).hexdigest()
        except OSError:
            return None

    def selection(self) -> dict[str, Any]:
        path = self.selection_path
        if not path.exists():
            return {"path": str(path), "rows": [], "digest": None}
        try:
            mappings = _read_mailboxes_csv(path)
        except ConfigError as exc:  # e.g. an alias edited by hand into something invalid
            raise HttpError(409, f"the saved selection is invalid; fix or re-save it: "
                                 f"{clean(exc)[:300]}") from exc
        except (OSError, ValueError, csv.Error) as exc:
            raise HttpError(500, f"cannot read {path}: {clean(exc)[:300]}") from exc
        rows = [{"source": m.source, "destination": m.destination, "name": m.name or "",
                 "quota_mib": m.quota_mib, "aliases": list(m.aliases)} for m in mappings]
        taken: set[str] = set()
        for n, row in enumerate(rows, 1):
            for address in (row["destination"], *row["aliases"]):
                if address in taken:
                    raise HttpError(409, f"the saved selection is invalid; fix or re-save "
                                         f"it: row {n}: {address} is used twice")
                taken.add(address)
        return {"path": str(path), "rows": rows, "digest": self._selection_digest()}

    def save_selection(self, body: object) -> dict[str, Any]:
        rows = body.get("rows") if isinstance(body, dict) else None
        if not isinstance(rows, list) or len(rows) > MAX_SELECTION_ROWS:
            raise HttpError(400, f"rows must be a list of at most {MAX_SELECTION_ROWS} rows")
        parsed: list[dict[str, Any]] = []
        sources: set[str] = set()
        taken: dict[str, int] = {}  # destination or alias -> row that claims it
        for n, row in enumerate(rows, 1):
            entry = _selection_row(row, n)
            if entry["source"] in sources:
                raise HttpError(400, f"row {n}: duplicate source {entry['source']}")
            sources.add(entry["source"])
            for address in (entry["destination"], *entry["aliases"]):
                if address in taken:
                    raise HttpError(400, f"row {n}: {address} is already used by row "
                                         f"{taken[address]} (as destination or alias)")
                taken[address] = n
            parsed.append(entry)
        self._write_selection(parsed)
        return {"path": str(self.selection_path), "rows": parsed,
                "digest": self._selection_digest()}

    def _write_selection(self, rows: list[dict[str, Any]]) -> None:
        """Atomic replace of the CSV that ``--mailboxes`` reads, mode 0600."""
        buf = io.StringIO()
        buf.write(SELECTION_HEADER + "\n")
        writer = csv.writer(buf, lineterminator="\n")
        for r in rows:
            quota = "" if r["quota_mib"] is None else r["quota_mib"]
            writer.writerow([r["source"], r["destination"], r["name"], quota,
                             ";".join(r.get("aliases", []))])
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

    def _job_argv(self, body: object) -> tuple[str, list[str], bytes | None]:
        """Validated request -> ``cli.main`` arguments plus the selection bytes a
        per-job snapshot is written from (``start_job`` substitutes its path)."""
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
        mail_since = body.get("mail_since")
        if mail_since not in (None, ""):
            if not isinstance(mail_since, str) or not _ISO_DATE.match(mail_since):
                raise HttpError(400, "mail_since must be a date like 2020-01-01")
            try:
                datetime.strptime(mail_since, "%Y-%m-%d")  # noqa: DTZ007 - date only
            except ValueError as exc:
                raise HttpError(400, "mail_since is not a real date") from exc
        else:
            mail_since = None
        sample = body.get("sample", 0)
        sample = 0 if sample is None else sample
        if isinstance(sample, bool) or not isinstance(sample, int) or not 0 <= sample <= MAX_SAMPLE:
            raise HttpError(400, f"sample must be an integer from 0 to {MAX_SAMPLE}")
        argv = ["--config", self.config_path] if self.config_path else []
        snapshot: bytes | None = None
        if command != "cleanup":
            # page-started jobs act on the saved selection alone (never the config's own
            # list), and only on the exact bytes the operator confirmed: those are copied
            # to a per-job file so a concurrent save cannot change what the job reads
            try:
                data = self.selection_path.read_bytes()
            except OSError as exc:
                raise HttpError(409, "no saved selection: load the tenant list, tick "
                                     "mailboxes and save the selection first") from exc
            digest = body.get("selection_digest")
            if digest != hashlib.sha256(data).hexdigest():
                if digest is None:
                    raise HttpError(409, "this page is older than the server: reload the page "
                                         "(F5) and start the job again")
                raise HttpError(409, "the saved selection changed since the page loaded it; "
                                     "reload the selection and start the job again")
            argv += ["--mailboxes", _SNAPSHOT_PLACEHOLDER, "--mailboxes-only"]
            snapshot = data
        elif self.selection_path.is_file():
            argv += ["--mailboxes", str(self.selection_path), "--mailboxes-only"]
        argv.append(command)
        if dry_run:
            argv.append("--dry-run")
        if only:
            argv += ["--only", only]
        if mailbox:
            argv.append(f"--mailbox={mailbox}")  # '=' form: never parsed as an option
        if command == "verify" and sample:
            argv += ["--sample", str(sample)]
        if mail_since and command in ("plan", "migrate", "verify"):
            argv += ["--mail-since", mail_since]
        return command, argv, snapshot

    def _running(self) -> Job | None:  # caller holds the lock
        return next((job for job in self._jobs.values() if job.running), None)

    def running_job(self) -> Job | None:
        with self._lock:
            return self._running()

    def start_job(self, body: object) -> Job:
        command, argv, data = self._job_argv(body)
        with self._settings_lock, self._lock:  # settings cannot change under a starting job
            running = self._running()
            if running is not None:
                raise HttpError(409, f"job {running.id} ({running.command}) is still running")
            if data is not None:  # written only for a job that really starts
                snapshot = self._write_snapshot(data)
                argv = [snapshot if a == _SNAPSHOT_PLACEHOLDER else a for a in argv]
            self._seq += 1
            job = Job(secrets.token_hex(8), self._seq, command, argv, JobOutput(self.redact))
            finished = sorted((j for j in self._jobs.values() if not j.running),
                              key=lambda j: j.seq)
            for old in finished[:max(0, len(self._jobs) + 1 - MAX_JOBS_KEPT)]:
                del self._jobs[old.id]
            thread = threading.Thread(target=self._run, args=(job,), name=f"job-{job.id}",
                                      daemon=True)
            try:
                thread.start()
            except RuntimeError as exc:  # cannot start a thread: never leave a phantom job
                raise HttpError(503, f"cannot start the job: {exc}") from exc
            self._jobs[job.id] = job
            log.info("job %s started: %s", job.id, " ".join(argv))
        return job

    def _write_snapshot(self, data: bytes) -> str:
        jobs_dir = self.cfg.state_dir / "jobs"
        jobs_dir.mkdir(mode=0o700, exist_ok=True)
        fd, snapshot = tempfile.mkstemp(prefix="selection-", suffix=".csv", dir=jobs_dir)
        with os.fdopen(fd, "wb") as fh:  # mkstemp creates the file 0600
            fh.write(data)
        return snapshot

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
            for i, arg in enumerate(job.args):  # the private selection snapshot, if any
                if arg == "--mailboxes" and i + 1 < len(job.args):
                    with contextlib.suppress(OSError):
                        snap = Path(job.args[i + 1]).resolve()
                        if snap.parent == (self.cfg.state_dir / "jobs").resolve():
                            snap.unlink()
            job.output.close()  # before `finished`: a finished job's output is complete
            with self._lock:  # the exit code first: reading the report must never cost it
                job.exit_code = code
            outcome = self._job_outcome(job)  # before `finished`: an outcome is final
            with self._lock:
                job.outcome = outcome
                job.finished = _now()

    def _job_outcome(self, job: Job) -> dict[str, str] | None:
        """Headline of the report this job wrote, named by its own ``report: <path>`` line:
        never another run's report (an external CLI run of the same command, an earlier
        job). Only a plain file name directly inside ``<state_dir>/reports`` is accepted, and
        the report must be of the job's command. Never raises: the job thread must always
        get to record its exit code."""
        try:
            raw = job.output.report_path()
            if raw is None:
                return None
            named = Path(raw)
            reports = self.cfg.state_dir / "reports"
            if (not _REPORT_NAME.fullmatch(named.name)
                    or named.parent.resolve() != reports.resolve()):
                return None
            data = self._read_report(reports / named.name)
            if not isinstance(data, dict) or data.get("command") != job.command:
                return None
            headline = summarize_report(self._redact_report(data))["headline"]
            return {"level": headline["level"], "text": self.redact(headline["text"])}
        except Exception as exc:  # e.g. an unresolvable path
            log.warning("job %s (%s): could not read its report: %s", job.id, job.command,
                        exc.__class__.__name__)
        return None

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
        return {**detail, "output_tail": job.output.tail(TAIL_LINES),
                "progress": job.output.progress()}

    def job_output(self, job_id: str) -> str:
        return self._job(job_id).output.text()

    # -- reports -----------------------------------------------------------------------

    def latest_report(self, command: str | None) -> Any:
        """Newest ``<state_dir>/reports/*.json`` (regular files only), optionally the newest
        of one command (``?command=verify``)."""
        if command is not None and command not in JOB_COMMANDS:
            raise HttpError(400, f"command must be one of {', '.join(JOB_COMMANDS)}")
        for _, data in self._reports():
            if command is None or (isinstance(data, dict) and data.get("command") == command):
                return data
        raise HttpError(404, "no report found")

    def overview(self) -> dict[str, Any]:
        """The page's start screen: one entry per step plus the summary of the newest file.

        Per step, real (non-dry-run) reports are scanned newest first: a full run decides
        the step, and so does a partial run (``--only``/``--mailbox``/``--mail-since``) that
        found problems; a clean partial run is only a fallback for when nothing older
        decides, so a narrow green check cannot hide an older failing full one. A step filled
        from a report older than an unreadable file says so instead of showing it as current.
        Reads at most ``MAX_REPORTS_SCANNED`` files, stops once every step is decided, and
        remembers per file what it learned, so a poll without new reports reads nothing.
        Every string is redacted on the way out."""
        files = self._report_files()[:MAX_REPORTS_SCANNED]
        decided: dict[str, dict[str, Any] | None] = dict.fromkeys(STEP_COMMANDS)
        fallback: dict[str, dict[str, Any]] = {}
        latest: dict[str, Any] | None = None
        newer_unreadable = False
        with self._report_cache_lock:
            self._prune_report_cache({f.key for f in files})
            for n, report_file in enumerate(files):
                meta = self._report_meta(report_file, newest=n == 0)
                if n == 0:
                    cached = self._latest_summary
                    latest = (cached[1] if meta.readable and cached is not None
                              and cached[0] == report_file.key else unreadable_summary())
                if not meta.readable:
                    newer_unreadable = True
                    continue
                command = meta.command
                if command not in decided or decided[command] is not None or meta.dry_run:
                    continue
                step = _stale(meta.step) if newer_unreadable else dict(meta.step)
                if not meta.partial or meta.level in ("warn", "bad"):
                    decided[command] = step  # a full run, or a partial one that found problems
                elif command not in fallback:
                    fallback[command] = step  # a clean partial run: only if nothing else
                if all(value is not None for value in decided.values()):
                    break
        steps = {command: decided[command] or fallback.get(command)
                 for command in STEP_COMMANDS}
        return self._redact_tree({"steps": steps, "latest": latest})

    def _prune_report_cache(self, keys: set[tuple[str, int, int]]) -> None:
        """Forget files that are gone, changed or no longer among the newest ones."""
        for key in [key for key in self._report_cache if key not in keys]:
            del self._report_cache[key]
        if self._latest_summary is not None and self._latest_summary[0] not in keys:
            self._latest_summary = None

    def _report_meta(self, report_file: _ReportFile, newest: bool) -> _ReportMeta:
        """What the overview needs from one file, from the cache when the file is unchanged;
        the newest file's full summary is kept too (caller holds the cache lock)."""
        meta = self._report_cache.get(report_file.key)
        have_summary = (self._latest_summary is not None
                        and self._latest_summary[0] == report_file.key)
        if meta is not None and (not newest or not meta.readable or have_summary):
            return meta
        data = self._read_report(report_file.path)
        if data is _UNAVAILABLE:
            # not cached: a permission or descriptor problem can clear up without the
            # file's name, mtime or size (the cache key) changing
            return _ReportMeta(readable=False)
        if data is not _UNREADABLE:
            try:  # redact the report whole first: summarising cuts strings to 300 characters,
                # and a longer secret cut in half would no longer be recognised afterwards
                data = self._redact_report(data)
            except RecursionError:  # nested deeper than the redactor can walk
                data = _UNREADABLE
        if data is _UNREADABLE:
            meta = _ReportMeta(readable=False)
        else:
            summary, step = summarize_with_step(data)
            meta = _ReportMeta(readable=True, command=summary["command"],
                               dry_run=summary["dry_run"], partial=summary["partial"],
                               level=summary["headline"]["level"], step=step)
            if newest:
                self._latest_summary = (report_file.key, summary)
        self._report_cache[report_file.key] = meta
        return meta

    def _redact_tree(self, value: Any) -> Any:
        """Every string value inside ``value`` (dicts and lists, recursively) through
        ``redact``. Dict keys are left alone: they are field names."""
        if isinstance(value, str):
            return self.redact(value)
        if isinstance(value, dict):
            return {key: self._redact_tree(item) for key, item in value.items()}
        if isinstance(value, list):
            return [self._redact_tree(item) for item in value]
        return value

    def _redact_report(self, data: Any) -> Any:
        """A report read from disk, redacted before it is summarised (the summary cuts
        strings to their display length, after which a longer secret would no longer
        match). The keys of ``mailboxes`` are data too (addresses) and are redacted as well;
        two that become equal stay two entries. Every other key is a field name."""
        data = self._redact_tree(data)
        mailboxes = data.get("mailboxes") if isinstance(data, dict) else None
        if not isinstance(mailboxes, dict):
            return data
        renamed: dict[Any, Any] = {}
        for key, entry in mailboxes.items():
            shown = self.redact(key) if isinstance(key, str) else key
            n = 1
            while shown in renamed:  # never let one entry replace another
                n += 1
                shown = f"{self.redact(key)} ({n})"
            renamed[shown] = entry
        return {**data, "mailboxes": renamed}

    def _report_files(self) -> list[_ReportFile]:
        """``<state_dir>/reports/*.json`` that are regular files (no symlinks), newest first
        by modification time, then by name."""
        candidates: list[tuple[float, str, _ReportFile]] = []
        for path in (self.cfg.state_dir / "reports").glob("*.json"):
            try:
                st = path.lstat()
            except OSError:
                continue
            if stat.S_ISREG(st.st_mode):
                candidates.append((st.st_mtime, path.name, _ReportFile(
                    path, st.st_mtime, (path.name, st.st_mtime_ns, st.st_size))))
        return [report_file for _, _, report_file in sorted(candidates, reverse=True)]

    def _reports(self, limit: int | None = None) -> Iterable[tuple[_ReportFile, Any]]:
        """``(file, parsed JSON)`` of the report files, newest first; unreadable files are
        skipped but count towards ``limit``, the number of files considered at most."""
        for report_file in self._report_files()[:limit]:
            data = self._read_report(report_file.path)
            if data is not _UNREADABLE and data is not _UNAVAILABLE:
                yield report_file, data

    def _read_report(self, path: Path) -> Any:
        """The parsed report, ``_UNAVAILABLE`` when the file could not be opened or read
        (or is too large), ``_UNREADABLE`` when its content is not JSON."""
        try:
            raw = _read_report_bytes(path)
        except OSError as exc:
            log.warning("skipping unreadable report %s: %s", clean(path.name)[:200],
                        exc.__class__.__name__)
            return _UNAVAILABLE
        try:
            return json.loads(raw.decode("utf-8"))
        except (ValueError, RecursionError) as exc:  # e.g. cut short by a crash
            log.warning("skipping unreadable report %s: %s", clean(path.name)[:200],
                        exc.__class__.__name__)
            return _UNREADABLE


_UNREADABLE = object()  # marks a report file whose content could not be parsed
_UNAVAILABLE = object()  # marks a report file that could not be opened or read right now


def _read_report_bytes(path: Path) -> bytes:
    """The file's bytes, read through one descriptor: a symlink or anything but a regular
    file of at most ``MAX_REPORT_BYTES`` (checked on the descriptor, so a file swapped in
    after ``lstat`` gains nothing) raises ``OSError``. ``O_NONBLOCK``: a FIFO swapped in
    cannot hang the request."""
    flags = os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0) | getattr(os, "O_NONBLOCK", 0)
    fd = os.open(path, flags)
    with os.fdopen(fd, "rb") as fh:
        st = os.fstat(fh.fileno())
        if not stat.S_ISREG(st.st_mode):
            raise OSError(f"{path.name} is not a regular file")
        if st.st_size > MAX_REPORT_BYTES:
            raise OSError(f"{path.name} is larger than {MAX_REPORT_BYTES} bytes")
        data = fh.read(MAX_REPORT_BYTES + 1)
    if len(data) > MAX_REPORT_BYTES:  # grew while being read
        raise OSError(f"{path.name} is larger than {MAX_REPORT_BYTES} bytes")
    return data


def _stale(step: dict[str, Any] | None) -> dict[str, Any]:
    """A step whose report is older than a report file that could not be read."""
    return {**(step or {}), "level": "unknown", "text": "A newer report could not be read"}


@dataclass(frozen=True)
class _ReportFile:
    path: Path
    mtime: float
    key: tuple[str, int, int]  # (name, st_mtime_ns, st_size): the overview cache key


@dataclass(frozen=True)
class _ReportMeta:
    """What the overview keeps per report file between requests (small on purpose)."""

    readable: bool
    command: str = ""
    dry_run: bool = False
    partial: bool = False
    level: str = "unknown"
    step: dict[str, Any] = field(default_factory=dict)


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
        self._check_host_and_origin()
        if path == "/":
            _allow(self.command, "GET")
            self._read_body()
            return app.page()
        if not path.startswith("/api/"):
            raise HttpError(404, "not found")
        if not app.authorized(self.headers.get("Authorization")):
            self.close_connection = True  # no body is read for an unauthenticated request
            raise HttpError(401, "unauthorized")
        body = self._read_body()
        response = self._api(path, parse_qs(query), body)
        if self.command != "GET":
            log.info("%s %s -> %s", self.command, path, response.status)
        return response

    def _check_host_and_origin(self) -> None:
        """DNS rebinding and cross-site defence on top of the token: the Host header must
        name this server (loopback, its bind address or a plain host name without a port
        mismatch), and a non-GET request that carries an Origin must come from it."""
        if len(self.headers.get_all("Host") or []) > 1:
            raise HttpError(400, "duplicate Host header")
        host = (self.headers.get("Host") or "").strip().lower()
        hostname = host_only(host)
        allowed = self.server.allowed_hosts
        if hostname and hostname not in allowed and not is_loopback(hostname):
            raise HttpError(421, "unexpected Host header")
        origin = self.headers.get("Origin")
        if origin and self.command != "GET":
            origin_authority = origin.split("://", 1)[-1].lower().rstrip("/")
            if origin_authority != host:  # the browser's Origin names the Host it used
                raise HttpError(403, "cross-origin request refused")

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
        if path == "/api/overview":
            _allow(method, "GET")
            return json_response(200, app.overview())
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

    def __init__(self, address: tuple[str, int], app: WebApp,
                 allowed_hosts: frozenset[str] = frozenset()) -> None:
        self.app = app
        # Host header values accepted besides loopback: the bind address and any names
        # the operator passes with --allow-host (e.g. the docker service name)
        self.allowed_hosts = frozenset(h.lower() for h in allowed_hosts) | {address[0].lower()}
        super().__init__(address, Handler)

    def server_bind(self) -> None:
        # HTTPServer.server_bind resolves the host's FQDN, which can stall without DNS;
        # the name is never used here.
        socketserver.TCPServer.server_bind(self)
        self.server_name, self.server_port = str(self.server_address[0]), self.server_address[1]


class WebServer6(WebServer):
    address_family = socket.AF_INET6


def make_server(cfg: Config, bind: str, port: int, token: str, page_path: Path, *,
                config_path: str | None = None, out: TextIO | None = None,
                lock_settings: bool = False,
                allowed_hosts: Iterable[str] = ()) -> WebServer:
    """Read the page and bind (port 0 picks a free port); ``serve_forever`` runs it.
    ``config_path`` is handed to every job; ``out`` receives a delegated sign-in prompt."""
    template = Path(page_path).read_text(encoding="utf-8")
    app = WebApp(cfg, token, template, config_path=config_path, out=out,
                 lock_settings=lock_settings)
    server_class = WebServer6 if ":" in bind else WebServer
    return server_class((bind, port), app, frozenset(allowed_hosts))


def serve(cfg: Config, bind: str, port: int, token: str, page_path: Path, *,
          config_path: str | None = None, out: TextIO | None = None,
          lock_settings: bool = False, allowed_hosts: Iterable[str] = ()) -> None:
    """Serve until Ctrl-C. A job still running then stops with the process; like any
    interrupted run it resumes when started again, and a temporary app password it leaves
    behind is removed by the next run for that mailbox or by ``cleanup``."""
    server = make_server(cfg, bind, port, token, page_path, config_path=config_path, out=out,
                         lock_settings=lock_settings, allowed_hosts=allowed_hosts)
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
