"""Command line entry point: ``o365mig plan|migrate|verify|cleanup|provision|web``.

Exit codes (ISC-115): 0 everything succeeded, 1 any item failed / was skipped (verify)
or a mailbox errored, 2 configuration or sign-in error.
"""

from __future__ import annotations

import argparse
import csv
import fcntl
import logging
import os
import re
import secrets
import sys
import threading
from collections.abc import Callable, Iterable
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from importlib import resources
from pathlib import Path
from typing import Any, TextIO

from . import __version__
from .auth import AuthError, TokenProvider
from .calendar_sync import CalendarMigrator
from .config import ENV_CONFIG, Config, ConfigError, MailboxMapping, load_config
from .contacts_sync import ContactsMigrator
from .dav import SogoDav
from .graph import GraphClient
from .imap_dest import ImapDestination
from .mail import MailMigrator
from .mailcow import (
    APP_PASSWORD_NAME,
    MailcowApi,
    MailcowError,
    generate_password,
    is_our_app_password,
)
from .report import Progress, RunReport, clean, utc_stamp, verify_summary
from .state import State

log = logging.getLogger("o365_to_mailcow")

KINDS = ("mail", "calendar", "contacts")
COMMANDS = {
    "plan": "show what would be migrated (read-only; creates nothing)",
    "migrate": "migrate mail, calendars and contacts (re-runnable)",
    "verify": "compare source and destination counts per folder and calendar",
    "cleanup": f"delete every '{APP_PASSWORD_NAME}-*' app password",
    "provision": "create missing destination mailboxes in mailcow (opt-in; never part of "
                 "migrate)",
    "web": "serve the local web UI: pick mailboxes, run the commands above, read reports",
}


# -- arguments -------------------------------------------------------------------------

def _port(text: str) -> int:
    if not (text.isascii() and text.isdigit()) or not 1 <= int(text) <= 65535:
        raise argparse.ArgumentTypeError(f"invalid port {text!r} (1-65535)")
    return int(text)


def build_parser() -> argparse.ArgumentParser:
    common = argparse.ArgumentParser(add_help=False)
    s = argparse.SUPPRESS  # so options work before and after the command
    common.add_argument("--config", default=s,
                        help="config TOML (default: $O365MIG_CONFIG)")
    common.add_argument("--mailboxes", metavar="CSV", default=s,
                        help="CSV of source[,destination] addresses (adds to config)")
    common.add_argument("--mailboxes-only", action="store_true", default=s,
                        help="use only the --mailboxes CSV, ignoring run.mailboxes in the config")
    common.add_argument("--only", choices=KINDS, default=s,
                        help="restrict to one kind of data")
    common.add_argument("--mailbox", metavar="ADDRESS", default=s,
                        help="restrict to one mailbox (source or destination address)")
    common.add_argument("--dry-run", action="store_true", default=s,
                        help="list and count only; no writes anywhere")
    common.add_argument("--keep-app-passwords", action="store_true", default=s,
                        help="do not delete the temporary app passwords at the end")
    common.add_argument("-v", "--verbose", action="count", default=s,
                        help="debug logging")
    parser = argparse.ArgumentParser(
        prog="o365mig", parents=[common],
        description="Migrate Microsoft 365 mail, calendars and contacts into mailcow "
                    "via Microsoft Graph.")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = parser.add_subparsers(dest="command", required=True, metavar="COMMAND")
    for name, text in COMMANDS.items():
        p = sub.add_parser(name, parents=[common], help=text, description=text)
        if name == "verify":
            p.add_argument("--sample", type=int, metavar="N", default=0,
                           help="re-download N random migrated messages per mailbox and "
                                "compare SHA-256 with the destination copy")
        if name == "web":
            p.add_argument("--bind", default="127.0.0.1", metavar="ADDRESS",
                           help="listen address (default 127.0.0.1; anything else exposes "
                                "the UI to the network)")
            p.add_argument("--port", type=_port, default=8080, help="port (default 8080)")
            p.add_argument("--lock-settings", action="store_true",
                           help="refuse changes to the connection settings from the page")
            p.add_argument("--allow-host", action="append", default=[], metavar="NAME",
                           help="extra Host header value to accept (loopback and the bind "
                                "address are always accepted)")
    return parser


@dataclass(frozen=True)
class Options:
    command: str
    config: str | None
    mailboxes_csv: str | None
    only: str | None
    mailboxes_only: bool
    mailbox: str | None
    dry_run: bool
    keep_app_passwords: bool
    verbose: int
    sample: int

    @classmethod
    def from_args(cls, ns: argparse.Namespace) -> Options:
        return cls(
            command=ns.command, config=getattr(ns, "config", None),
            mailboxes_csv=getattr(ns, "mailboxes", None), only=getattr(ns, "only", None),
            mailboxes_only=getattr(ns, "mailboxes_only", False),
            mailbox=getattr(ns, "mailbox", None), dry_run=getattr(ns, "dry_run", False),
            keep_app_passwords=getattr(ns, "keep_app_passwords", False),
            verbose=getattr(ns, "verbose", 0) or 0, sample=getattr(ns, "sample", 0) or 0,
        )

    @property
    def kinds(self) -> tuple[str, ...]:
        return (self.only,) if self.only else KINDS


def select_mailboxes(cfg: Config, address: str | None) -> list[MailboxMapping]:
    if not address:
        return list(cfg.mailboxes)
    wanted = address.strip().lower()
    chosen = [m for m in cfg.mailboxes if wanted in (m.source, m.destination)]
    if not chosen:
        raise ConfigError(f"--mailbox {address} is not in the configured mailbox list")
    return chosen


# -- logging ---------------------------------------------------------------------------

class SecretFilter(logging.Filter):
    """Last line of defence: replace any registered secret value with ``***``."""

    def __init__(self) -> None:
        super().__init__()
        self._secrets: set[str] = set()
        self._lock = threading.Lock()

    def add(self, value: str | None) -> None:
        if value and len(value) >= 6:
            with self._lock:
                self._secrets.add(value)

    def redact(self, text: str) -> str:
        with self._lock:
            secrets_ = list(self._secrets)
        for value in secrets_:
            text = text.replace(value, "***")
        return text

    def filter(self, record: logging.LogRecord) -> bool:
        msg = record.getMessage()
        redacted = self.redact(msg)
        if redacted != msg:
            record.msg, record.args = redacted, ()
        if record.exc_info:  # tracebacks are formatted later; redact them now instead
            record.exc_text = self.redact(
                logging.Formatter().formatException(record.exc_info))
            record.exc_info = None
        return True


def setup_logging(cfg: Config, verbose: int, secret_filter: SecretFilter,
                  stream: TextIO | None = None) -> tuple[Path, list[logging.Handler]]:
    logs = cfg.state_dir / "logs"
    logs.mkdir(parents=True, exist_ok=True)
    path = logs / f"{utc_stamp()}.log"
    os.close(os.open(path, os.O_WRONLY | os.O_CREAT | os.O_APPEND, 0o600))
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(threadName)s %(name)s: %(message)s")
    level = logging.DEBUG if verbose else getattr(logging, cfg.log_level, logging.INFO)
    handlers: list[logging.Handler] = [
        logging.StreamHandler(sys.stderr if stream is None else stream),
        logging.FileHandler(path, encoding="utf-8")]
    root = logging.getLogger()
    for h in handlers:
        h.setFormatter(fmt)
        h.setLevel(level)
        h.addFilter(secret_filter)
        root.addHandler(h)
    root.setLevel(min(root.level or logging.WARNING, level))
    # these log raw protocol traffic (message bodies, tokens) at DEBUG
    for noisy in ("imapclient", "msal", "urllib3"):
        logging.getLogger(noisy).setLevel(logging.WARNING)
    return path, handlers


def auth_hint(cfg: Config, exc: Exception) -> str:
    """ISC-126: explain the usual cause of a refused device-code flow."""
    text = str(exc)
    if cfg.auth_mode == "delegated" and re.search(
            r"device|authorization_pending|AADSTS", text, re.IGNORECASE):
        return (
            f"Microsoft sign-in failed: {text}\n"
            "The device-code sign-in was refused. Most tenants block it with a Conditional "
            "Access policy on 'authentication flows' (Conditions > Authentication flows > "
            "Device code flow). Either exclude this administrator from that policy for the "
            "duration of the migration, or use app-only sign-in: set auth_mode = \"app\" in "
            "[microsoft], add the application permissions Mail.Read, Calendars.Read and "
            "Contacts.Read with admin consent, and provide O365MIG_CLIENT_SECRET."
        )
    return f"Microsoft sign-in failed: {text}"


# -- runner ----------------------------------------------------------------------------

class Runner:
    """Executes one command against the selected mailboxes."""

    def __init__(self, cfg: Config, opts: Options, mailboxes: list[MailboxMapping],
                 secret_filter: SecretFilter, out: Any = None, err: Any = None) -> None:
        self.cfg = cfg
        self.opts = opts
        self.mailboxes = mailboxes
        self.secret_filter = secret_filter
        self.out = sys.stdout if out is None else out
        self.err = sys.stderr if err is None else err
        self.report = RunReport(opts.command, cfg.state_dir, dry_run=opts.dry_run,
                                redactor=secret_filter.redact)
        self.progress = Progress(self.err)
        self._tokens: TokenProvider | None = None
        self._state: State | None = None
        self._init_lock = threading.Lock()  # mailbox threads share one State/TokenProvider

    # -- shared resources --------------------------------------------------------------

    def say(self, line: str = "") -> None:
        print(self.secret_filter.redact(line), file=self.out, flush=True)

    @property
    def tokens(self) -> TokenProvider:
        with self._init_lock:
            if self._tokens is None:  # a device-code prompt goes where the output goes
                self._tokens = TokenProvider(self.cfg, out=self.out)
            return self._tokens

    @property
    def state(self) -> State:
        with self._init_lock:
            if self._state is None:
                self._state = State(self.cfg.state_dir / "state.db")
            return self._state

    def api(self, allow_provision: bool = False) -> MailcowApi:
        return MailcowApi(self.cfg.mailcow_host, self.cfg.mailcow_api_key,
                          verify=self.cfg.mailcow_ca_file or True,
                          allow_provision=allow_provision)

    def graph(self) -> GraphClient:
        """One client per mailbox, so Graph's 4-in-flight limit applies per mailbox."""
        return GraphClient(self.tokens)

    def sign_in(self) -> None:
        """Acquire a token once in the main thread (device-code prompt, early failure)."""
        self.tokens.get_token()

    def close(self) -> None:
        if self._state is not None:
            self._state.close()

    def parallel(self, fn: Callable[[MailboxMapping], None]) -> None:
        workers = max(1, min(self.cfg.parallel_mailboxes, len(self.mailboxes)))
        with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="mbx") as pool:
            for fut in [pool.submit(self._guarded, fn, m) for m in self.mailboxes]:
                fut.result()

    def _guarded(self, fn: Callable[[MailboxMapping], None], m: MailboxMapping) -> None:
        self.report.mailbox(m.source, m.destination)
        try:
            fn(m)
        except Exception as exc:  # one mailbox must never take the others down
            log.error("%s: %s: %s", m.source, exc.__class__.__name__, exc)
            log.debug("traceback", exc_info=True)
            self.report.error(m.source, f"{exc.__class__.__name__}: {exc}")
            self.report.set(m.source, "status", "failed")

    # -- app passwords -----------------------------------------------------------------

    def purge_app_passwords(self, api: MailcowApi, address: str, include_named: bool) -> int:
        """Delete app passwords recorded in state for ``address`` (and, with
        ``include_named``, every one named o365-migration). Returns how many were deleted."""
        listed = api.list_app_passwords(address)
        listed_ids = {str(p["id"]) for p in listed}
        recorded = {mid for mbx, mid in self.state.app_passwords() if mbx == address}
        targets = recorded & listed_ids
        if include_named:
            targets |= {str(p["id"]) for p in listed if is_our_app_password(p)}
        deleted = 0
        for mid in sorted(targets):
            if self.opts.dry_run:
                self.say(f"  would delete app password id={mid} of {address}")
                continue
            api.delete_app_password(mid)
            self.state.forget_app_password(address, mid)
            deleted += 1
        if not self.opts.dry_run:
            for mid in recorded - listed_ids:  # already gone in mailcow
                self.state.forget_app_password(address, mid)
        return deleted

    def with_app_password(self, api: MailcowApi, m: MailboxMapping,
                          fn: Callable[[str], None]) -> None:
        """Create a temporary app password, run ``fn(password)``, always delete it."""
        try:
            leftovers = self.purge_app_passwords(api, m.destination, include_named=False)
        except MailcowError as exc:  # a stuck leftover must not block the migration itself
            leftovers = 0
            log.warning("%s: could not delete leftover app password(s): %s; "
                        "run 'o365mig cleanup' later", m.destination, exc)
        if leftovers:
            log.info("%s: deleted %d leftover app password(s)", m.destination, leftovers)
        pw_id, password = api.create_app_password(m.destination)
        self.secret_filter.add(password)
        self.state.record_app_password(m.destination, pw_id)  # before first use (ISC-125)
        try:
            fn(password)
        finally:
            if self.opts.keep_app_passwords:
                log.warning("%s: keeping app password id=%s (--keep-app-passwords)",
                            m.destination, pw_id)
            else:
                try:
                    api.delete_app_password(pw_id)
                    self.state.forget_app_password(m.destination, pw_id)
                except MailcowError as exc:
                    log.error("%s: could not delete app password id=%s: %s; "
                              "run 'o365mig cleanup'", m.destination, pw_id, exc)
                    self.report.error(m.source, f"app password id={pw_id} not deleted: {exc}")

    def dest_factory(self, m: MailboxMapping, password: str) -> Callable[[], ImapDestination]:
        cfg = self.cfg
        return lambda: ImapDestination(cfg.mailcow_host, cfg.imap_port, m.destination,
                                       password, verify=cfg.mailcow_ca_file or True)

    def dav(self, m: MailboxMapping, password: str) -> SogoDav:
        return SogoDav(self.cfg.mailcow_host, m.destination, password,
                       verify=self.cfg.mailcow_ca_file or True)

    def exists(self, api: MailcowApi, m: MailboxMapping) -> bool:
        if api.mailbox_exists(m.destination):
            return True
        msg = f"destination mailbox {m.destination} does not exist in mailcow; skipped"
        log.error("%s: %s", m.source, msg)
        self.report.set(m.source, "status", "missing")
        self.report.error(m.source, msg)
        return False

    def mailbox_failed(self, source: str) -> bool:
        entry = self.report.data["mailboxes"].get(source, {})
        return entry.get("status") in ("failed", "missing") or bool(entry.get("errors"))


# -- commands --------------------------------------------------------------------------

def _human_bytes(n: int) -> str:
    size = float(n)
    for unit in ("B", "KiB", "MiB", "GiB"):
        if size < 1024:
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} TiB"


def cmd_plan(r: Runner) -> int:
    """Read-only: Graph listing plus mailcow ``get`` calls. No add, no DAV, no IMAP."""
    r.sign_in()

    def one(m: MailboxMapping) -> None:
        exists = r.exists(r.api(), m)
        graph = r.graph()
        lines = [f"{m.source} -> {m.destination}"
                 + ("" if exists else "  [MISSING IN MAILCOW: will be skipped]")]
        if "mail" in r.opts.kinds:
            plan = MailMigrator(r.cfg, graph, r.state, None, m, True, r.progress).plan()
            r.report.set(m.source, "mail_plan", plan)
            size = (_human_bytes(plan.total_bytes) if plan.bytes_known
                    else f"at least {_human_bytes(plan.total_bytes)}")
            lines.append(f"  mail: {plan.total_messages} messages, {size}, "
                         f"{len(plan.folders) - plan.skipped_folders} folders "
                         f"({plan.skipped_folders} skipped)")
            for f in plan.folders:
                what = f"skip: {f.skip_reason}" if f.skip else f"-> {f.dest_name}"
                lines.append(f"    {clean(f.source_path)} ({f.total}) {what}")
        for kind, cls in (("calendar", CalendarMigrator), ("contacts", ContactsMigrator)):
            if kind not in r.opts.kinds:
                continue
            cplan = cls(r.cfg, graph, r.state, None, m, True, r.progress).plan()
            r.report.set(m.source, f"{kind}_plan", cplan)
            total = sum(c.count for c in cplan.collections)
            lines.append(f"  {kind}: {total} items in {len(cplan.collections)} collection(s)")
            for c in cplan.collections:
                lines.append(f"    {clean(c.name)} ({c.count}) -> {c.slug}")
            for sk in cplan.skipped:
                lines.append(f"    skip (not owned by mailbox): {clean(sk)}")
        r.say("\n".join(lines))
        if exists:
            r.report.set(m.source, "status", "ok")

    r.parallel(one)
    return 1 if any(r.mailbox_failed(m.source) for m in r.mailboxes) else 0


def _run_migrators(r: Runner, m: MailboxMapping, password: str | None) -> None:
    dry = r.opts.dry_run
    graph = r.graph()
    summary: list[str] = []
    failed = 0
    if "mail" in r.opts.kinds:
        factory = r.dest_factory(m, password) if password else None
        res = MailMigrator(r.cfg, graph, r.state, factory, m, dry, r.progress).migrate()
        r.report.set(m.source, "mail", res)
        failed += res.failed + len(res.errors)
        verb = "would append" if dry else "appended"
        n = res.total("would_append") if dry else res.total("appended")
        delta_folders = sum(1 for f in res.folders if f.delta_pass)
        summary.append(
            f"mail: {verb} {n}, already done {res.total('already_done')}, "
            f"found by Message-ID {res.total('dedup_hits')}, failed {res.total('failed')}, "
            f"skipped too large {res.total('skipped_too_large')}, "
            f"removed at source (not applied) {res.total('removed_in_source')}"
            + (f"; {delta_folders} folder(s) checked for changes since the last run only"
               if delta_folders else ""))
        # folder errors are also collected in res.errors (with the source path); print each once
        for e in res.errors:
            summary.append(f"  error: {clean(e)}")
        for f in res.folders:
            if f.error and not any(f.error in e for e in res.errors):
                summary.append(f"  error: {clean(f.dest_name)}: {clean(f.error)}")
    dav = r.dav(m, password) if password else None
    for kind, cls in (("calendar", CalendarMigrator), ("contacts", ContactsMigrator)):
        if kind not in r.opts.kinds:
            continue
        cres = cls(r.cfg, graph, r.state, dav, m, dry, r.progress).migrate()
        r.report.set(m.source, kind, cres)
        failed += cres.failed
        put = sum(c.would_put if dry else c.put for c in cres.collections)
        summary.append(
            f"{kind}: {'would put' if dry else 'put'} {put}, unchanged "
            f"{sum(c.unchanged for c in cres.collections)}, failed "
            f"{sum(c.failed for c in cres.collections)}, skipped {len(cres.skipped)}")
        for w in cres.warnings:
            summary.append(f"  warning: {clean(w)}")
        for e in cres.errors:
            summary.append(f"  error: {clean(e)}")
    r.report.set(m.source, "status", "failed" if failed else "ok")
    r.say(f"{m.source} -> {m.destination}\n  " + "\n  ".join(summary))


def cmd_migrate(r: Runner) -> int:
    r.sign_in()
    stop = threading.Event()
    ticker = threading.Thread(target=r.progress.run_ticker, args=(stop,), daemon=True)
    ticker.start()

    def one(m: MailboxMapping) -> None:
        api = r.api()
        if not r.exists(api, m):
            return
        if r.opts.dry_run:
            _run_migrators(r, m, None)
        else:
            r.with_app_password(api, m, lambda pw: _run_migrators(r, m, pw))

    try:
        r.parallel(one)
    finally:
        stop.set()
    return 1 if any(r.mailbox_failed(m.source) for m in r.mailboxes) else 0


def cmd_verify(r: Runner) -> int:
    if r.opts.dry_run:
        r.say("verify needs a temporary app password; --dry-run is not supported for verify")
        return 2
    r.sign_in()

    def one(m: MailboxMapping) -> None:
        api = r.api()
        if not r.exists(api, m):
            return

        def run(password: str) -> None:
            graph = r.graph()
            if "mail" in r.opts.kinds:
                mig = MailMigrator(r.cfg, graph, r.state, r.dest_factory(m, password), m,
                                   False, r.progress)
                r.report.set(m.source, "mail", mig.verify(sample=r.opts.sample))
            dav = r.dav(m, password)
            for kind, cls in (("calendar", CalendarMigrator), ("contacts", ContactsMigrator)):
                if kind in r.opts.kinds:
                    r.report.set(m.source, kind,
                                 cls(r.cfg, graph, r.state, dav, m, False, r.progress).verify())
            r.report.set(m.source, "status", "ok")

        r.with_app_password(api, m, run)

    r.parallel(one)
    lines, problems = verify_summary(r.report.data["mailboxes"])
    r.say("\n".join(lines))
    return 1 if problems else 0


def cmd_cleanup(r: Runner) -> int:
    addresses: list[str] = [m.destination for m in r.mailboxes]
    if not r.opts.mailbox:  # also anything recorded in state for unlisted mailboxes
        addresses += sorted({mbx for mbx, _ in r.state.app_passwords()} - set(addresses))
    api = r.api()
    total, errors = 0, 0
    for address in addresses:
        try:
            n = r.purge_app_passwords(api, address, include_named=True)
        except MailcowError as exc:
            errors += 1
            log.error("%s: cleanup failed: %s", address, exc)
            r.report.error(address, str(exc))
            continue
        total += n
        r.report.set(address, "deleted_app_passwords", n)
        r.say(f"{address}: deleted {n} app password(s)")
    r.say(f"cleanup: {'dry run, ' if r.opts.dry_run else ''}deleted {total} app password(s) "
          f"across {len(addresses)} mailbox(es)")
    return 1 if errors else 0


def cmd_provision(r: Runner) -> int:
    """Create the destination mailboxes that do not exist yet, then their aliases. Passwords
    are generated, written to a 0600 ledger in the state directory before each create and
    never logged; users must change them at first login. Domains are never created."""
    api = r.api(allow_provision=True)
    domains_ok: dict[str, bool] = {}
    ledger: _PasswordLedger | None = None  # opened before the first create, never lost
    ready: set[str] = set()  # destinations that exist, were created, or would be (dry run)
    errors = 0
    try:
        for m in r.mailboxes:
            try:
                if api.mailbox_exists(m.destination):
                    r.say(f"{m.destination}: exists")
                    r.report.set(m.source, "provision", "exists")
                    ready.add(m.destination)
                    continue
                domain = m.destination.partition("@")[2]
                if domain not in domains_ok:
                    domains_ok[domain] = api.domain_exists(domain)
                if not domains_ok[domain]:
                    errors += 1
                    msg = f"domain {domain} does not exist in mailcow; add it in the UI first"
                    r.say(f"{m.destination}: {msg}")
                    r.report.error(m.source, msg)
                    continue
                name = m.name or m.destination.partition("@")[0]
                quota = m.quota_mib or r.cfg.provision_quota_mib
                if r.opts.dry_run:
                    r.say(f"{m.destination}: would create ({name!r}, {quota} MiB)")
                    r.report.set(m.source, "provision", "would create")
                    ready.add(m.destination)
                    continue
                password = generate_password()
                r.secret_filter.add(password)
                if ledger is None:
                    ledger = _PasswordLedger(r.cfg.state_dir)
                ledger.write(m.destination, password, "pending")  # on disk before the call
                try:
                    api.create_mailbox(m.destination, name, quota, password,
                                       tls_enforce=r.cfg.provision_tls_enforce)
                except MailcowError:
                    # the call may have completed server-side (timeout after commit): keep
                    # the row when the mailbox now exists, otherwise mark it failed
                    if api.mailbox_exists(m.destination):
                        ledger.write(m.destination, password, "created-unconfirmed")
                        ready.add(m.destination)
                    else:
                        ledger.write(m.destination, password, "failed")
                    raise
                ledger.write(m.destination, password, "created")
                ready.add(m.destination)
                r.say(f"{m.destination}: created ({name!r}, {quota} MiB)")
                r.report.set(m.source, "provision", "created")
            except MailcowError as exc:
                errors += 1
                log.error("%s: provisioning failed: %s", m.destination, exc)
                r.report.error(m.source, f"provisioning failed: {exc}")
    finally:
        if ledger is not None:
            ledger.close()
            r.say(f"initial passwords written to {ledger.path} (mode 0600); users must "
                  "change them at first login. Delete the file once distributed.")
    errors += _provision_aliases(r, api, domains_ok, ready)
    return 1 if errors else 0


def _provision_aliases(r: Runner, api: MailcowApi, domains_ok: dict[str, bool],
                       ready: set[str]) -> int:
    """Create each mailbox's extra addresses as mailcow aliases pointing at it. Only for
    destinations that exist here (or were just created, or would be in a dry run); aliases
    in a domain mailcow does not host and addresses that already exist are skipped."""
    wanted = [(alias, m) for m in r.mailboxes for alias in m.aliases]
    if not wanted:
        return 0
    errors = 0
    try:
        existing = {k.lower(): v.lower() for k, v in api.list_aliases().items()}
    except MailcowError as exc:
        r.say(f"aliases: cannot list existing aliases: {exc}")
        r.report.error("aliases", f"cannot list aliases: {exc}")
        return 1
    created = would_create = 0
    for alias, m in wanted:
        domain = alias.partition("@")[2]
        try:
            if domain not in domains_ok:
                domains_ok[domain] = api.domain_exists(domain)
            if not domains_ok[domain]:
                r.say(f"  alias {alias}: skipped, domain {domain} is not hosted in mailcow")
                continue
            if alias in existing:
                if m.destination in existing[alias].split(","):
                    r.say(f"  alias {alias}: exists -> {m.destination}")
                else:
                    r.say(f"  alias {alias}: exists but points elsewhere; left unchanged")
                continue
            if m.destination not in ready:  # never forward to a mailbox that is not here
                r.say(f"  alias {alias}: skipped, {m.destination} is not a mailbox here")
                continue
            if api.mailbox_exists(alias):
                r.say(f"  alias {alias}: skipped, a mailbox with that address exists")
                continue
            if r.opts.dry_run:
                would_create += 1
                r.say(f"  alias {alias}: would create -> {m.destination}")
                continue
            api.create_alias(alias, m.destination)
            existing[alias] = m.destination
            created += 1
            r.say(f"  alias {alias}: created -> {m.destination}")
        except MailcowError as exc:
            errors += 1
            log.error("alias %s: %s", alias, exc)
            r.report.error(m.source, f"alias {alias}: {exc}")
    r.report.set("aliases", "created" if not r.opts.dry_run else "would_create",
                 created if not r.opts.dry_run else would_create)
    return errors


class _PasswordLedger:
    """Append-only, fsync'd CSV of generated initial passwords: opened with a unique name
    before the first mailbox is created, every row written before and after the API call,
    so an interrupted run never loses a password it already set."""

    def __init__(self, state_dir: Path) -> None:
        base = state_dir / f"provisioned-{utc_stamp()}"
        for n in range(1, 1000):
            path = Path(f"{base}.csv" if n == 1 else f"{base}-{n}.csv")
            try:
                fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL
                             | getattr(os, "O_NOFOLLOW", 0), 0o600)
                break
            except FileExistsError:
                continue
        else:  # pragma: no cover - a thousand runs in one second
            raise OSError("cannot create a unique provisioned-passwords file")
        self.path = path
        self._fh = os.fdopen(fd, "w", encoding="utf-8", newline="")
        self._csv = csv.writer(self._fh)
        self._csv.writerow(["mailbox", "initial_password", "status"])
        self._flush()

    def write(self, address: str, password: str, status: str) -> None:
        self._csv.writerow([address, password, status])
        self._flush()

    def _flush(self) -> None:
        self._fh.flush()
        os.fsync(self._fh.fileno())

    def close(self) -> None:
        self._fh.close()


HANDLERS: dict[str, Callable[[Runner], int]] = {
    "plan": cmd_plan, "migrate": cmd_migrate, "verify": cmd_verify, "cleanup": cmd_cleanup,
    "provision": cmd_provision,
}


# -- main ------------------------------------------------------------------------------

def cmd_web(cfg: Config, opts: Options, bind: str, port: int, err: TextIO,
            lock_settings: bool = False, allowed_hosts: list[str] | None = None) -> int:
    """Serve the web UI until Ctrl-C (security model: see ``web.py``). Takes no state lock:
    each job started from the page runs ``main`` and takes it then."""
    allowed_hosts = [h.rsplit(":", 1)[0] if h.count(":") == 1 else h
                     for h in (allowed_hosts or [])]
    from . import web  # imported here because web imports this module

    token = os.environ.get(web.ENV_TOKEN) or secrets.token_urlsafe(32)
    if len(token) < web.MIN_TOKEN_CHARS:
        print(f"configuration error: {web.ENV_TOKEN} must be at least "
              f"{web.MIN_TOKEN_CHARS} characters", file=err)
        return 2
    config_path = opts.config or os.environ.get(ENV_CONFIG)
    secret_filter = SecretFilter()
    for value in (cfg.client_secret, cfg.mailcow_api_key, token):
        secret_filter.add(value)
    handler = logging.StreamHandler(err)
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    handler.addFilter(secret_filter)
    saved = (web.log.level, web.log.propagate)
    web.log.addHandler(handler)
    web.log.setLevel(logging.INFO)
    web.log.propagate = False  # server messages stay out of the jobs' logs, and vice versa
    try:
        if not web.is_loopback(bind):
            web.log.warning("the web UI on %s is reachable from the network; the token is its "
                            "only protection and the connection is not encrypted", bind)
        if os.environ.get("O365MIG_WEB_TOKEN"):
            print(f"web UI: {web.base_url(bind, port)}/#token=<O365MIG_WEB_TOKEN from .env>",
                  file=err, flush=True)
        else:
            print(f"web UI: {web.base_url(bind, port)}/#token={token}", file=err, flush=True)
        page = resources.files(__package__) / "web_static" / "index.html"
        with resources.as_file(page) as page_path:
            web.serve(cfg, bind, port, token, page_path, out=err,
                      config_path=str(Path(config_path).resolve()) if config_path else None,
                      lock_settings=lock_settings, allowed_hosts=allowed_hosts)
    except OSError as exc:
        print(f"web UI: cannot start on {bind}:{port}: {exc}", file=err)
        return 2
    finally:
        web.log.removeHandler(handler)
        web.log.setLevel(saved[0])
        web.log.propagate = saved[1]
    return 0


def main(argv: Iterable[str] | None = None, *, stdout: TextIO | None = None,
         stderr: TextIO | None = None) -> int:
    """Run one command. Everything it prints, logs or reports as progress goes to
    ``stdout``/``stderr`` (default: the process's streams); the web UI passes a buffer."""
    out = sys.stdout if stdout is None else stdout
    err = sys.stderr if stderr is None else stderr
    args = list(sys.argv[1:] if argv is None else argv)
    if args and args[0] == "o365mig":  # `docker run IMAGE o365mig ...` with ENTRYPOINT o365mig
        args = args[1:]
    ns = build_parser().parse_args(args)
    opts = Options.from_args(ns)
    try:
        cfg = load_config(opts.config, opts.mailboxes_csv, mailboxes_only=opts.mailboxes_only,
                          require_mailboxes=opts.command != "web",
                          require_credentials=opts.command != "web")
        mailboxes = select_mailboxes(cfg, opts.mailbox)
        cfg.state_dir.mkdir(parents=True, exist_ok=True)
    except (ConfigError, OSError) as exc:
        print(f"configuration error: {exc}", file=err)
        return 2
    if opts.command == "web":
        return cmd_web(cfg, opts, ns.bind, ns.port, err, lock_settings=ns.lock_settings,
                       allowed_hosts=ns.allow_host)

    lock = _acquire_lock(cfg.state_dir)
    if lock is None:
        print("another o365mig run is using this state directory; wait for it to finish",
              file=err)
        return 2
    secret_filter = SecretFilter()
    secret_filter.add(cfg.client_secret)
    secret_filter.add(cfg.mailcow_api_key)
    log_path, handlers = setup_logging(cfg, opts.verbose, secret_filter, err)
    runner = Runner(cfg, opts, mailboxes, secret_filter, out=out, err=err)
    code = 1
    try:
        log.info("o365mig %s %s: %d mailbox(es), log %s", __version__, opts.command,
                 len(mailboxes), log_path)
        code = HANDLERS[opts.command](runner)
    except AuthError as exc:
        print(auth_hint(cfg, exc), file=err)
        code = 2
    except KeyboardInterrupt:
        print("interrupted; state is saved, re-run the same command to resume", file=err)
        code = 1
    finally:
        try:
            path = runner.report.write(code)
            print(f"report: {path}", file=err)
        except OSError as exc:
            print(f"could not write report: {exc}", file=err)
        runner.close()
        root = logging.getLogger()
        for h in handlers:
            root.removeHandler(h)
            h.close()
        lock.close()
    return code


def _acquire_lock(state_dir: Path):
    """Hold an exclusive lock on ``<state_dir>/.lock`` for the whole run, so two containers
    sharing one state volume cannot purge each other's live app passwords."""
    path = state_dir / ".lock"
    fd = os.open(path, os.O_WRONLY | os.O_CREAT, 0o600)
    handle = os.fdopen(fd, "w")
    try:
        fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        handle.close()
        return None
    return handle


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
