"""Command line entry point: ``o365mig plan|migrate|verify|cleanup``.

Exit codes (ISC-115): 0 everything succeeded, 1 any item failed / was skipped (verify)
or a mailbox errored, 2 configuration or sign-in error.
"""

from __future__ import annotations

import argparse
import fcntl
import logging
import os
import re
import sys
import threading
from collections.abc import Callable, Iterable
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from . import __version__
from .auth import AuthError, TokenProvider
from .calendar_sync import CalendarMigrator
from .config import Config, ConfigError, MailboxMapping, load_config
from .contacts_sync import ContactsMigrator
from .dav import SogoDav
from .graph import GraphClient
from .imap_dest import ImapDestination
from .mail import MailMigrator
from .mailcow import APP_PASSWORD_NAME, MailcowApi, MailcowError, is_our_app_password
from .report import Progress, RunReport, clean, utc_stamp, verify_summary
from .state import State

log = logging.getLogger("o365_to_mailcow")

KINDS = ("mail", "calendar", "contacts")
COMMANDS = {
    "plan": "show what would be migrated (read-only; creates nothing)",
    "migrate": "migrate mail, calendars and contacts (re-runnable)",
    "verify": "compare source and destination counts per folder and calendar",
    "cleanup": f"delete every '{APP_PASSWORD_NAME}-*' app password",
}


# -- arguments -------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    common = argparse.ArgumentParser(add_help=False)
    s = argparse.SUPPRESS  # so options work before and after the command
    common.add_argument("--config", default=s,
                        help="config TOML (default: $O365MIG_CONFIG)")
    common.add_argument("--mailboxes", metavar="CSV", default=s,
                        help="CSV of source[,destination] addresses (adds to config)")
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
    return parser


@dataclass(frozen=True)
class Options:
    command: str
    config: str | None
    mailboxes_csv: str | None
    only: str | None
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


def setup_logging(cfg: Config, verbose: int,
                  secret_filter: SecretFilter) -> tuple[Path, list[logging.Handler]]:
    logs = cfg.state_dir / "logs"
    logs.mkdir(parents=True, exist_ok=True)
    path = logs / f"{utc_stamp()}.log"
    os.close(os.open(path, os.O_WRONLY | os.O_CREAT | os.O_APPEND, 0o600))
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(threadName)s %(name)s: %(message)s")
    level = logging.DEBUG if verbose else getattr(logging, cfg.log_level, logging.INFO)
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stderr),
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
                 secret_filter: SecretFilter, out: Any = None) -> None:
        self.cfg = cfg
        self.opts = opts
        self.mailboxes = mailboxes
        self.secret_filter = secret_filter
        self.out = out or sys.stdout
        self.report = RunReport(opts.command, cfg.state_dir, dry_run=opts.dry_run,
                                redactor=secret_filter.redact)
        self.progress = Progress()
        self._tokens: TokenProvider | None = None
        self._state: State | None = None
        self._init_lock = threading.Lock()  # mailbox threads share one State/TokenProvider

    # -- shared resources --------------------------------------------------------------

    def say(self, line: str = "") -> None:
        print(line, file=self.out, flush=True)

    @property
    def tokens(self) -> TokenProvider:
        with self._init_lock:
            if self._tokens is None:
                self._tokens = TokenProvider(self.cfg)
            return self._tokens

    @property
    def state(self) -> State:
        with self._init_lock:
            if self._state is None:
                self._state = State(self.cfg.state_dir / "state.db")
            return self._state

    def api(self) -> MailcowApi:
        return MailcowApi(self.cfg.mailcow_host, self.cfg.mailcow_api_key)

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
                                       password)

    def dav(self, m: MailboxMapping, password: str) -> SogoDav:
        return SogoDav(self.cfg.mailcow_host, m.destination, password)

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
        summary.append(
            f"mail: {verb} {n}, already done {res.total('already_done')}, "
            f"found by Message-ID {res.total('dedup_hits')}, failed {res.total('failed')}, "
            f"skipped too large {res.total('skipped_too_large')}, "
            f"removed at source (not applied) {res.total('removed_in_source')}")
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


HANDLERS: dict[str, Callable[[Runner], int]] = {
    "plan": cmd_plan, "migrate": cmd_migrate, "verify": cmd_verify, "cleanup": cmd_cleanup,
}


# -- main ------------------------------------------------------------------------------

def main(argv: Iterable[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if args and args[0] == "o365mig":  # `docker run IMAGE o365mig ...` with ENTRYPOINT o365mig
        args = args[1:]
    opts = Options.from_args(build_parser().parse_args(args))
    try:
        cfg = load_config(opts.config, opts.mailboxes_csv)
        mailboxes = select_mailboxes(cfg, opts.mailbox)
        cfg.state_dir.mkdir(parents=True, exist_ok=True)
    except (ConfigError, OSError) as exc:
        print(f"configuration error: {exc}", file=sys.stderr)
        return 2

    lock = _acquire_lock(cfg.state_dir)
    if lock is None:
        print("another o365mig run is using this state directory; wait for it to finish",
              file=sys.stderr)
        return 2
    secret_filter = SecretFilter()
    secret_filter.add(cfg.client_secret)
    secret_filter.add(cfg.mailcow_api_key)
    log_path, handlers = setup_logging(cfg, opts.verbose, secret_filter)
    runner = Runner(cfg, opts, mailboxes, secret_filter)
    code = 1
    try:
        log.info("o365mig %s %s: %d mailbox(es), log %s", __version__, opts.command,
                 len(mailboxes), log_path)
        code = HANDLERS[opts.command](runner)
    except AuthError as exc:
        print(auth_hint(cfg, exc), file=sys.stderr)
        code = 2
    except KeyboardInterrupt:
        print("interrupted; state is saved, re-run the same command to resume",
              file=sys.stderr)
        code = 1
    finally:
        try:
            path = runner.report.write(code)
            print(f"report: {path}", file=sys.stderr)
        except OSError as exc:
            print(f"could not write report: {exc}", file=sys.stderr)
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
