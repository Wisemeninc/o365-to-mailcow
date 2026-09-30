"""Configuration loading and validation.

Secrets are accepted from environment variables so they never need to live in the
config file. Values read from the file are still honoured for convenience, but the
file's permissions are checked and a warning is emitted when it is readable by others.
"""

from __future__ import annotations

import csv
import logging
import os
import stat
import tomllib
from dataclasses import dataclass, field
from pathlib import Path

log = logging.getLogger(__name__)

ENV_CONFIG = "O365MIG_CONFIG"
ENV_CLIENT_SECRET = "O365MIG_CLIENT_SECRET"  # noqa: S105 - name of a variable, not a secret
ENV_API_KEY = "O365MIG_MAILCOW_API_KEY"  # noqa: S105

GRAPH_HOST = "graph.microsoft.com"
LOGIN_HOST = "login.microsoftonline.com"

APP_SCOPES = ["https://graph.microsoft.com/.default"]
DELEGATED_SCOPES = [
    "Mail.Read.Shared",
    "Calendars.Read.Shared",
    "Contacts.Read.Shared",
    "User.Read",
]
# offline_access is added by MSAL automatically for delegated flows; listing it is an error.


class ConfigError(Exception):
    """Raised for any invalid or incomplete configuration. CLI maps this to exit code 2."""


@dataclass(frozen=True)
class MailboxMapping:
    source: str
    destination: str


@dataclass(frozen=True)
class Config:
    # Microsoft
    tenant_id: str
    client_id: str
    auth_mode: str  # "app" | "delegated"
    client_secret: str | None
    # mailcow
    mailcow_host: str
    mailcow_api_key: str
    # run
    state_dir: Path
    mailboxes: tuple[MailboxMapping, ...]
    parallel_mailboxes: int = 2
    max_message_bytes: int = 150 * 1024 * 1024
    calendar_exceptions_from_days: int = 730
    calendar_exceptions_to_days: int = 1095
    contacts_photos: bool = False
    calendar_attendees: str = "keep"  # "keep" (with SCHEDULE-AGENT=CLIENT) | "strip"
    verify_tls: bool = True  # exists only so tests can point at a local server; default True
    imap_port: int = 993
    log_level: str = "INFO"
    source_folder_skip: tuple[str, ...] = field(default_factory=tuple)

    @property
    def scopes(self) -> list[str]:
        return APP_SCOPES if self.auth_mode == "app" else list(DELEGATED_SCOPES)

    @property
    def authority(self) -> str:
        return f"https://{LOGIN_HOST}/{self.tenant_id}"


def _check_file_mode(path: Path) -> None:
    mode = stat.S_IMODE(path.stat().st_mode)
    if mode & 0o077:
        log.warning(
            "config file %s is readable by other users (mode %o); chmod 600 it", path, mode
        )


def _read_mailboxes_csv(path: Path) -> list[MailboxMapping]:
    rows: list[MailboxMapping] = []
    with path.open(newline="", encoding="utf-8") as fh:
        for raw in csv.reader(fh):
            if not raw or raw[0].strip().startswith("#"):
                continue
            src = raw[0].strip().lower()
            dst = (raw[1].strip().lower() if len(raw) > 1 and raw[1].strip() else src)
            if "@" not in src or "@" not in dst:
                raise ConfigError(f"invalid mailbox row in {path}: {raw!r}")
            rows.append(MailboxMapping(src, dst))
    return rows


def load_config(config_path: str | os.PathLike[str] | None, mailboxes_csv: str | None = None,
                env: dict[str, str] | None = None) -> Config:
    env = dict(os.environ if env is None else env)
    path_str = config_path or env.get(ENV_CONFIG)
    if not path_str:
        raise ConfigError(f"no config file: pass --config or set {ENV_CONFIG}")
    path = Path(path_str)
    if not path.is_file():
        raise ConfigError(f"config file not found: {path}")
    _check_file_mode(path)
    with path.open("rb") as fh:
        data = tomllib.load(fh)

    ms = data.get("microsoft", {})
    mc = data.get("mailcow", {})
    run = data.get("run", {})

    missing: list[str] = []

    def req(section: dict, key: str, name: str) -> str:
        val = section.get(key)
        if val in (None, ""):
            missing.append(name)
            return ""
        return str(val)

    tenant_id = req(ms, "tenant_id", "microsoft.tenant_id")
    client_id = req(ms, "client_id", "microsoft.client_id")
    auth_mode = str(ms.get("auth_mode", "app")).lower()
    client_secret = env.get(ENV_CLIENT_SECRET) or ms.get("client_secret") or None
    mailcow_host = req(mc, "host", "mailcow.host")
    api_key = env.get(ENV_API_KEY) or mc.get("api_key") or ""
    if not api_key:
        missing.append(f"mailcow.api_key (or {ENV_API_KEY})")
    if auth_mode == "app" and not client_secret:
        missing.append(f"microsoft.client_secret (or {ENV_CLIENT_SECRET})")
    if auth_mode not in ("app", "delegated"):
        raise ConfigError(f"microsoft.auth_mode must be 'app' or 'delegated', got {auth_mode!r}")

    mailboxes: list[MailboxMapping] = []
    csv_path = mailboxes_csv or run.get("mailboxes_csv")
    if csv_path:
        mailboxes.extend(_read_mailboxes_csv(Path(csv_path)))
    for entry in run.get("mailboxes", []):
        if isinstance(entry, str):
            mailboxes.append(MailboxMapping(entry.lower(), entry.lower()))
        elif isinstance(entry, dict) and "source" in entry:
            dest = entry.get("destination", entry["source"])
            mailboxes.append(MailboxMapping(entry["source"].lower(), dest.lower()))
        else:
            raise ConfigError(f"invalid run.mailboxes entry: {entry!r}")
    if not mailboxes:
        missing.append("run.mailboxes (or --mailboxes CSV)")

    if missing:
        raise ConfigError("missing required configuration: " + ", ".join(missing))

    if str(run.get("calendar_attendees", "keep")).lower() not in ("keep", "strip"):
        raise ConfigError("run.calendar_attendees must be 'keep' or 'strip'")
    if "/" in mailcow_host or ":" in mailcow_host:
        raise ConfigError("mailcow.host must be a bare hostname, e.g. mail.example.net")

    state_dir = Path(run.get("state_dir", "/state"))
    return Config(
        tenant_id=tenant_id,
        client_id=client_id,
        auth_mode=auth_mode,
        client_secret=client_secret,
        mailcow_host=mailcow_host,
        mailcow_api_key=api_key,
        state_dir=state_dir,
        mailboxes=tuple(mailboxes),
        parallel_mailboxes=int(run.get("parallel_mailboxes", 2)),
        max_message_bytes=int(run.get("max_message_bytes", 150 * 1024 * 1024)),
        calendar_exceptions_from_days=int(run.get("calendar_exceptions_from_days", 730)),
        calendar_exceptions_to_days=int(run.get("calendar_exceptions_to_days", 1095)),
        contacts_photos=bool(run.get("contacts_photos", False)),
        calendar_attendees=str(run.get("calendar_attendees", "keep")).lower(),
        verify_tls=bool(run.get("verify_tls", True)),
        imap_port=int(run.get("imap_port", 993)),
        log_level=str(run.get("log_level", "INFO")).upper(),
        source_folder_skip=tuple(run.get("source_folder_skip", [])),
    )
