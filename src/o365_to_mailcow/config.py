"""Configuration loading and validation.

Secrets are accepted from environment variables so they never need to live in the
config file. Values read from the file are still honoured for convenience, but the
file's permissions are checked and a warning is emitted when it is readable by others.
"""

from __future__ import annotations

import csv
import logging
import os
import re
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


# RFC 1123 hostname (labels of letters, digits, hyphens; dots between), no port, no path.
_HOSTNAME = re.compile(
    r"^(?=.{1,253}$)(?!-)[A-Za-z0-9-]{1,63}(?<!-)(\.(?!-)[A-Za-z0-9-]{1,63}(?<!-))+$"
)
_LOCAL_PART = re.compile(r"^[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+(\.[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+)*$")


def valid_hostname(value: str) -> bool:
    """RFC 1123 host name with a non-numeric top label: never an IPv4 literal."""
    return bool(_HOSTNAME.match(value)) and not value.rsplit(".", 1)[-1].isdigit()


def valid_address(value: str) -> bool:
    """``local@host`` with a real host name (used for mailbox sources and destinations)."""
    local, _, domain = value.partition("@")
    return bool(local) and len(local) <= 64 and bool(_LOCAL_PART.match(local)) \
        and valid_hostname(domain)


def _bool(section: dict, key: str, default: bool) -> bool:
    value = section.get(key, default)
    if not isinstance(value, bool):
        raise ConfigError(f"run.{key} must be true or false, got {value!r}")
    return value


def _int(section: dict, key: str, default: int, lo: int, hi: int) -> int:
    value = section.get(key, default)
    if isinstance(value, bool) or not isinstance(value, int) or not lo <= value <= hi:
        raise ConfigError(f"run.{key} must be an integer between {lo} and {hi}, got {value!r}")
    return value


@dataclass(frozen=True)
class MailboxMapping:
    source: str
    destination: str
    name: str | None = None        # display name for `provision` (else the local part)
    quota_mib: int | None = None   # quota for `provision` (else run.provision_quota_mib)


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
    mailcow_ca_file: str | None = None  # private CA bundle for the mailcow host (API, DAV, IMAP)
    parallel_mailboxes: int = 2
    max_message_bytes: int = 150 * 1024 * 1024
    calendar_exceptions_from_days: int = 730
    calendar_exceptions_to_days: int = 1095
    contacts_photos: bool = False
    calendar_attendees: str = "keep"  # "keep" (with SCHEDULE-AGENT=CLIENT) | "strip"
    imap_port: int = 993
    provision_quota_mib: int = 3072  # default quota for mailboxes created by `provision`
    provision_tls_enforce: bool = False  # mailcow tls_enforce_in/out on created mailboxes
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
            cols = [c.strip() for c in raw] + ["", "", "", ""]
            src = cols[0].lower()
            dst = cols[1].lower() or src
            if "@" not in src or "@" not in dst:
                raise ConfigError(f"invalid mailbox row in {path}: {raw!r}")
            quota: int | None = None
            if cols[3]:
                if not cols[3].isdigit() or not 1 <= int(cols[3]) <= 1_000_000:
                    raise ConfigError(f"invalid quota (MiB) in {path}: {raw!r}")
                quota = int(cols[3])
            rows.append(MailboxMapping(src, dst, cols[2] or None, quota))
    return rows


SETTINGS_FILE = "settings.toml"  # written by the web UI into state_dir; wins over file and env


def settings_path(state_dir: Path) -> Path:
    return Path(state_dir) / SETTINGS_FILE


def read_settings(state_dir: Path) -> dict[str, dict[str, str]]:
    """The web UI's saved settings ({"microsoft": {...}, "mailcow": {...}}) or {}."""
    path = settings_path(state_dir)
    if not path.is_file():
        return {}
    try:
        with path.open("rb") as fh:
            data = tomllib.load(fh)
    except (OSError, tomllib.TOMLDecodeError) as exc:
        raise ConfigError(f"{path} is not valid TOML: {exc}") from exc
    return {k: v for k, v in data.items() if isinstance(v, dict)}


def load_config(config_path: str | os.PathLike[str] | None, mailboxes_csv: str | None = None,
                env: dict[str, str] | None = None, *, require_mailboxes: bool = True,
                require_credentials: bool = True, mailboxes_only: bool = False) -> Config:
    """Precedence for tenant/client ids, hosts and secrets: values saved by the web UI
    (``<state_dir>/settings.toml``) > environment variables > the config file.

    A saved host is only ever paired with a *saved* key, and saved tenant/client ids only
    with a saved client secret: the overlay never borrows a secret from the environment or
    the file, so redirecting the host cannot carry a credential elsewhere."""
    env = dict(os.environ if env is None else env)
    path_str = config_path or env.get(ENV_CONFIG)
    if not path_str:
        raise ConfigError(f"no config file: pass --config or set {ENV_CONFIG}")
    path = Path(path_str)
    if not path.is_file():
        raise ConfigError(f"config file not found: {path}")
    _check_file_mode(path)
    try:
        with path.open("rb") as fh:
            data = tomllib.load(fh)
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"{path} is not valid TOML: {exc}") from exc

    ms = dict(data.get("microsoft", {}))
    mc = dict(data.get("mailcow", {}))
    run = data.get("run", {})
    state_dir = Path(run.get("state_dir", "/state"))

    # values saved through the web UI override both the file and the environment
    saved = read_settings(state_dir)
    saved_ms, saved_mc = saved.get("microsoft", {}), saved.get("mailcow", {})
    for key in ("tenant_id", "client_id", "auth_mode", "client_secret"):
        if saved_ms.get(key):
            ms[key] = saved_ms[key]
    for key in ("host", "api_key"):
        if saved_mc.get(key):
            mc[key] = saved_mc[key]
    if saved_ms or saved_mc:
        log.warning("using connection settings saved by the web UI from %s (mailcow host %s)",
                    settings_path(state_dir), mc.get("host", ""))

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
    if saved_ms.get("tenant_id") or saved_ms.get("client_id"):
        client_secret = saved_ms.get("client_secret") or None  # paired with saved ids only
    else:
        client_secret = env.get(ENV_CLIENT_SECRET) or ms.get("client_secret") or None
    mailcow_host = req(mc, "host", "mailcow.host")
    if saved_mc.get("host"):
        api_key = saved_mc.get("api_key") or ""  # paired with the saved host only
    else:
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
    for entry in ([] if mailboxes_only and csv_path else run.get("mailboxes", [])):
        if isinstance(entry, str):
            src, dst = entry, entry
        elif isinstance(entry, dict) and isinstance(entry.get("source"), str):
            src, dst = entry["source"], entry.get("destination", entry["source"])
        else:
            raise ConfigError(f"invalid run.mailboxes entry: {entry!r}")
        if not isinstance(dst, str) or "@" not in src or "@" not in dst:
            raise ConfigError(f"invalid run.mailboxes entry: {entry!r}")
        name = entry.get("name") if isinstance(entry, dict) else None
        quota = entry.get("quota_mib") if isinstance(entry, dict) else None
        if (name is not None and not isinstance(name, str)) or (
            quota is not None and (isinstance(quota, bool) or not isinstance(quota, int)
                                   or not 1 <= quota <= 1_000_000)):
            raise ConfigError(f"invalid run.mailboxes entry: {entry!r}")
        mailboxes.append(MailboxMapping(src.strip().lower(), dst.strip().lower(), name, quota))
    if not mailboxes and require_mailboxes:
        missing.append("run.mailboxes (or --mailboxes CSV)")
    # Two mappings onto one destination would share one mailbox's app password and
    # IMAP folders from two threads; refuse rather than race (Silas M3).
    seen_dst: set[str] = set()
    seen_src: set[str] = set()
    for m in mailboxes:
        if m.destination in seen_dst:
            raise ConfigError(f"destination mailbox {m.destination} is listed more than once")
        if m.source in seen_src:
            raise ConfigError(f"source mailbox {m.source} is listed more than once")
        seen_dst.add(m.destination)
        seen_src.add(m.source)

    if missing and require_credentials:
        raise ConfigError("missing required configuration: " + ", ".join(missing))
    if missing and not require_credentials:  # the web UI can start and be configured
        missing = [m for m in missing if m.startswith("run.")]
        if missing:
            raise ConfigError("missing required configuration: " + ", ".join(missing))

    if str(run.get("calendar_attendees", "keep")).lower() not in ("keep", "strip"):
        raise ConfigError("run.calendar_attendees must be 'keep' or 'strip'")
    if mailcow_host and not valid_hostname(mailcow_host):
        raise ConfigError("mailcow.host must be a bare hostname, e.g. mail.example.net")
    for m in mailboxes:
        if not valid_address(m.source) or not valid_address(m.destination):
            raise ConfigError(f"invalid mailbox address in {m.source} -> {m.destination}")
    ca_file = mc.get("ca_file")
    if ca_file is not None:
        if not isinstance(ca_file, str) or not Path(ca_file).is_file():
            raise ConfigError(f"mailcow.ca_file must name a readable PEM file, got {ca_file!r}")

    return Config(
        tenant_id=tenant_id,
        client_id=client_id,
        auth_mode=auth_mode,
        client_secret=client_secret,
        mailcow_host=mailcow_host,
        mailcow_api_key=api_key,
        mailcow_ca_file=ca_file,
        state_dir=state_dir,
        mailboxes=tuple(mailboxes),
        parallel_mailboxes=_int(run, "parallel_mailboxes", 2, 1, 4),
        max_message_bytes=_int(run, "max_message_bytes", 150 * 1024 * 1024,
                               1024, 1024 * 1024 * 1024),
        calendar_exceptions_from_days=_int(run, "calendar_exceptions_from_days", 730, 0, 36500),
        calendar_exceptions_to_days=_int(run, "calendar_exceptions_to_days", 1095, 0, 36500),
        contacts_photos=_bool(run, "contacts_photos", False),
        calendar_attendees=str(run.get("calendar_attendees", "keep")).lower(),
        imap_port=_int(run, "imap_port", 993, 1, 65535),
        provision_quota_mib=_int(run, "provision_quota_mib", 3072, 1, 1_000_000),
        provision_tls_enforce=_bool(run, "provision_tls_enforce", False),
        log_level=str(run.get("log_level", "INFO")).upper(),
        source_folder_skip=tuple(run.get("source_folder_skip", [])),
    )
