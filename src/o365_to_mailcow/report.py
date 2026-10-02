"""Progress output and the JSON run report.

``Progress`` prints one line per active mailbox scope with done/total and items per
minute, at most every ``interval`` seconds from ``advance`` calls and at least every
``interval`` seconds from the ticker thread (ISC-118).

``RunReport`` collects per-mailbox results and writes
``<state_dir>/reports/<UTC timestamp>.json`` (ISC-114). ``verify_summary`` turns verify
results into lines plus a problem count; it never produces an all-clear line while any
skipped or failed category is non-zero (ISC-128).
"""

from __future__ import annotations

import dataclasses
import json
import os
import re
import sys
import threading
import time
from collections.abc import Callable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, TextIO

FAILED_ITEMS_LIMIT = 100  # failed/skipped items per mailbox section stored in a report
FAILED_ITEMS_SHOWN = 20  # of those, lines printed per section


@dataclasses.dataclass
class CollectionPlan:
    """One Graph calendar or contact folder and where it would go in SOGo."""

    source_id: str
    name: str
    slug: str
    count: int
    skip: bool = False
    skip_reason: str | None = None


@dataclasses.dataclass
class CollectionsPlan:
    mailbox: str
    kind: str  # "calendar" | "contacts"
    collections: list[CollectionPlan] = dataclasses.field(default_factory=list)
    skipped: list[str] = dataclasses.field(default_factory=list)


@dataclasses.dataclass
class CollectionResult:
    source_id: str
    name: str
    slug: str
    listed: int = 0
    put: int = 0
    unchanged: int = 0
    failed: int = 0
    would_put: int = 0
    error: str | None = None


@dataclasses.dataclass
class CollectionsResult:
    mailbox: str
    kind: str
    dry_run: bool
    collections: list[CollectionResult] = dataclasses.field(default_factory=list)
    skipped: list[str] = dataclasses.field(default_factory=list)
    fallbacks: list[str] = dataclasses.field(default_factory=list)
    warnings: list[str] = dataclasses.field(default_factory=list)
    errors: list[str] = dataclasses.field(default_factory=list)
    duration_s: float = 0.0
    # items currently failed or skipped (state.failed_items); not part of ``failed``
    failed_items: list[dict] = dataclasses.field(default_factory=list)
    failed_items_total: int = 0

    @property
    def failed(self) -> int:
        return (sum(c.failed for c in self.collections)
                + sum(1 for c in self.collections if c.error) + len(self.errors))


@dataclasses.dataclass
class CollectionVerify:
    name: str
    slug: str
    graph_count: int
    done: int
    failed: int
    dav_count: int
    expected: int
    mismatch: bool


@dataclasses.dataclass
class CollectionsVerify:
    mailbox: str
    kind: str
    collections: list[CollectionVerify] = dataclasses.field(default_factory=list)
    skipped: list[str] = dataclasses.field(default_factory=list)
    fallbacks: list[str] = dataclasses.field(default_factory=list)
    errors: list[str] = dataclasses.field(default_factory=list)
    failed_items: list[dict] = dataclasses.field(default_factory=list)
    failed_items_total: int = 0


_UNSAFE = re.compile(r"[\x00-\x1f\x7f\u200b-\u200f\u2028-\u202e\u2066-\u2069]")


def clean(text: object) -> str:
    """Strip control and bidi/zero-width characters from tenant-supplied strings before
    they reach a terminal or a log line (no escape-sequence or log-forging tricks)."""
    return _UNSAFE.sub("", str(text))


def utc_stamp(now: datetime | None = None) -> str:
    """Filesystem-safe UTC timestamp, e.g. ``20260930T201500Z``."""
    return (now or datetime.now(UTC)).strftime("%Y%m%dT%H%M%SZ")


# -- progress --------------------------------------------------------------------------

class NullProgress:
    """Progress sink that does nothing (tests, dry library use)."""

    def start(self, key: str, total: int) -> None:
        """Register ``total`` more items for ``key``."""

    def adjust_total(self, key: str, delta: int) -> None:
        """Correct the total once the real amount of work is known."""

    def advance(self, key: str, n: int = 1) -> None:
        """Record ``n`` processed items."""

    def phase(self, key: str, text: str) -> None:
        """Describe what ``key`` is doing while ``done`` cannot move (listing, indexing)."""

    def finish(self, key: str) -> None:
        """Mark ``key`` complete."""


class Progress(NullProgress):
    """Thread-safe progress printer."""

    def __init__(self, out: TextIO | None = None, interval: float = 30.0,
                 clock: Callable[[], float] = time.monotonic) -> None:
        self._out = out or sys.stderr
        self._interval = interval
        self._clock = clock
        self._lock = threading.Lock()
        self._scopes: dict[str, dict[str, float]] = {}
        self._phases: dict[str, str] = {}
        self._last_print = clock()

    def start(self, key: str, total: int) -> None:
        with self._lock:
            scope = self._scopes.setdefault(
                key, {"done": 0, "total": 0, "started": self._clock(), "finished": 0})
            scope["total"] += total

    def adjust_total(self, key: str, delta: int) -> None:
        with self._lock:
            scope = self._scopes.get(key)
            if scope is not None:
                scope["total"] = max(0, scope["total"] + delta)

    def advance(self, key: str, n: int = 1) -> None:
        with self._lock:
            scope = self._scopes.setdefault(
                key, {"done": 0, "total": 0, "started": self._clock(), "finished": 0})
            scope["done"] += n
        self.maybe_print()

    def phase(self, key: str, text: str) -> None:
        with self._lock:
            if text:
                self._phases[key] = text
            else:
                self._phases.pop(key, None)
        self.maybe_print()

    def finish(self, key: str) -> None:
        with self._lock:
            scope = self._scopes.get(key)
            if scope is None:
                return
            scope["finished"] = 1
            self._phases.pop(key, None)
            line = self._line(key, scope)
        print(f"{line} (finished)", file=self._out, flush=True)

    def _line(self, key: str, scope: dict[str, float]) -> str:
        elapsed = max(self._clock() - scope["started"], 1e-9)
        rate = scope["done"] / elapsed * 60.0
        line = f"[{key}] {int(scope['done'])}/{int(scope['total'])} items, {rate:.0f}/min"
        phase = self._phases.get(key)
        return f"{line} ({phase})" if phase else line

    def maybe_print(self, force: bool = False) -> bool:
        """Print all active scopes if ``interval`` elapsed (or ``force``)."""
        with self._lock:
            now = self._clock()
            if not force and now - self._last_print < self._interval:
                return False
            self._last_print = now
            lines = [self._line(k, s) for k, s in self._scopes.items() if not s["finished"]]
        for line in lines:
            print(line, file=self._out, flush=True)
        return bool(lines)

    def run_ticker(self, stop: threading.Event) -> None:
        """Thread target: print at least every ``interval`` seconds until ``stop`` is set."""
        while not stop.wait(self._interval):
            self.maybe_print(force=True)


# -- run report ------------------------------------------------------------------------

def _plain(obj: Any) -> Any:
    if dataclasses.is_dataclass(obj) and not isinstance(obj, type):
        return {k: _plain(v) for k, v in dataclasses.asdict(obj).items()}
    if isinstance(obj, dict):
        return {str(k): _plain(v) for k, v in obj.items()}
    if isinstance(obj, list | tuple | set):
        return [_plain(v) for v in obj]
    if isinstance(obj, Path):
        return str(obj)
    return obj


class RunReport:
    """Accumulates one command's results; thread-safe; written once at the end."""

    def __init__(self, command: str, state_dir: Path, dry_run: bool = False,
                 now: Callable[[], datetime] = lambda: datetime.now(UTC),
                 redactor: Callable[[str], str] | None = None,
                 scope: dict[str, Any] | None = None) -> None:
        self._now = now
        self._redact = redactor or (lambda text: text)
        self._state_dir = Path(state_dir)
        self._lock = threading.Lock()
        self._t0 = time.monotonic()
        started = now()
        self._stamp = utc_stamp(started)
        self.data: dict[str, Any] = {
            "command": command,
            "dry_run": dry_run,
            "started": started.isoformat(),
            "mailboxes": {},
        }
        if scope is not None:  # what this run was narrowed to (--only/--mailbox/--mail-since)
            self.data["scope"] = dict(scope)

    def mailbox(self, source: str, destination: str) -> None:
        with self._lock:
            self.data["mailboxes"].setdefault(source, {
                "destination": destination, "status": "pending", "errors": [],
            })

    def set(self, source: str, section: str, value: Any) -> None:
        with self._lock:
            self.data["mailboxes"].setdefault(source, {"errors": []})[section] = _plain(value)

    def error(self, source: str, message: str) -> None:
        message = clean(self._redact(message))
        with self._lock:
            self.data["mailboxes"].setdefault(source, {"errors": []}).setdefault(
                "errors", []).append(message)

    def write(self, exit_code: int) -> Path:
        with self._lock:
            self.data["finished"] = self._now().isoformat()
            self.data["duration_s"] = round(time.monotonic() - self._t0, 3)
            self.data["exit_code"] = exit_code
            reports = self._state_dir / "reports"
            reports.mkdir(parents=True, exist_ok=True)
            path = reports / f"{self._stamp}.json"
            n = 1
            while path.exists():
                n += 1
                path = reports / f"{self._stamp}-{n}.json"
            fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                json.dump(self.data, fh, indent=2, sort_keys=True)
                fh.write("\n")
            return path


# -- verify summary --------------------------------------------------------------------

def failed_item_lines(section: Any) -> list[str]:
    """One line per failed/skipped item of a plain-dict report section, at most
    ``FAILED_ITEMS_SHOWN``, plus how many more the report holds. Reports written before
    items were labelled have no ``failed_items``; anything unreadable is skipped."""
    if not isinstance(section, dict):
        return []
    items = section.get("failed_items")
    if not isinstance(items, list):
        return []
    lines: list[str] = []
    for item in items:
        if len(lines) >= FAILED_ITEMS_SHOWN:
            break
        if not isinstance(item, dict):
            continue

        def text(key: str, item: dict = item) -> str:
            value = item.get(key)
            return clean(value) if isinstance(value, str) else ""

        title = text("title") or "(title not recorded)"
        hint = text("hint")
        reason = text("error") or text("status")
        lines.append(f"    not copied: {text('place')} · {title}"
                     + (f" ({hint})" if hint else "") + f": {reason}")
    total = section.get("failed_items_total")
    if not isinstance(total, int) or isinstance(total, bool):
        total = len(items)
    if total > len(lines):
        lines.append(f"    … and {total - len(lines)} more (the report file lists up to "
                     f"{FAILED_ITEMS_LIMIT})")
    return lines


def verify_summary(mailboxes: dict[str, dict[str, Any]]) -> tuple[list[str], int]:
    """Render verify results (plain dicts from ``RunReport``) as lines + problem count.

    Problems (each makes verify exit 1): failed items, skipped items (too large, skipped
    folders holding messages), count mismatches, sample mismatches, errors, missing
    destination mailboxes, address-book fallbacks. Shared calendars that are skipped by
    design (ISC-83) and messages removed at the source are listed but are not problems.
    """
    lines: list[str] = []
    problems = 0
    listed_by_design = 0

    def problem(text: str, count: int = 1) -> None:
        nonlocal problems
        problems += count
        lines.append(f"  ! {text}")

    for source, entry in mailboxes.items():
        lines.append(f"{clean(source)} -> {clean(entry.get('destination', source))}")
        if entry.get("status") == "missing":
            problem("destination mailbox does not exist in mailcow")
        for err in entry.get("errors", []):
            problem(f"error: {clean(err)}")
        mail = entry.get("mail")
        if mail:
            lines.append("  mail folder                          graph  skip  fail  "
                         "expect  imap")
            for f in mail.get("folders", []):
                mark = "MISMATCH" if f["mismatch"] else "ok"
                lines.append(
                    f"    {clean(f['dest_name'])[:34]:<34} {f['graph_total']:>6} {f['skipped']:>5} "
                    f"{f['failed']:>5} {f['expected']:>7} {f['imap_count']:>5}  {mark}")
                if f["failed"]:
                    problem(f"mail {f['dest_name']}: {f['failed']} failed", f["failed"])
                if f["skipped"]:
                    problem(f"mail {f['dest_name']}: {f['skipped']} skipped (too large)",
                            f["skipped"])
                if f["mismatch"]:
                    problem(f"mail {f['dest_name']}: expected {f['expected']}, "
                            f"IMAP has {f['imap_count']}")
                if f.get("surplus_total"):
                    shown = f.get("surplus") or []
                    lines.append(f"      {f['surplus_total']} surplus cop"
                                 f"{'y' if f['surplus_total'] == 1 else 'ies'} of migrated "
                                 f"messages ({len(shown)} listed in the report JSON with "
                                 "every copy's UID; 'tool_uids' are this tool's appends). "
                                 "Remove with doveadm expunge; this tool never deletes.")
                if f.get("note"):
                    lines.append(f"      note: {f['note']}")
            for sk in mail.get("skipped_folders", []):
                text = (f"mail folder skipped: {clean(sk['path'])} ({sk['reason']}), "
                        f"{sk['total']} messages")
                if sk["total"]:
                    problem(text, sk["total"])
                else:
                    lines.append(f"  - {text}")
                    listed_by_design += 1
            if mail.get("sample_requested"):
                lines.append(
                    f"  sample: {mail['sample_checked']} compared, "
                    f"{len(mail['sample_mismatches'])} mismatched, "
                    f"{mail.get('sample_regenerated', 0)} re-rendered by Exchange (same "
                    f"headers), {mail['sample_unverifiable']} without Message-ID/gone at source")
                for mm in mail["sample_mismatches"]:
                    problem(f"sample content mismatch: {mm}")
            for err in mail.get("errors", []):
                problem(f"mail: {clean(err)}")
            # listed for information: the folders' failed/skipped counts already count them
            lines.extend(failed_item_lines(mail))
        for kind in ("calendar", "contacts"):
            sec = entry.get(kind)
            if not sec:
                continue
            for c in sec.get("collections", []):
                mark = "MISMATCH" if c["mismatch"] else "ok"
                lines.append(
                    f"  {kind} {clean(c['name'])[:30]:<30} graph {c['graph_count']:>5} "
                    f"fail {c['failed']:>4} expect {c['expected']:>5} "
                    f"dav {c['dav_count']:>5}  {mark}")
                if c["failed"]:
                    problem(f"{kind} {c['name']}: {c['failed']} failed", c["failed"])
                if c["mismatch"]:
                    problem(f"{kind} {c['name']}: expected {c['expected']}, "
                            f"DAV has {c['dav_count']}")
            for sk in sec.get("skipped", []):
                lines.append(f"  - {kind} skipped by design: {clean(sk)}")
                listed_by_design += 1
            for fb in sec.get("fallbacks", []):
                problem(f"{kind} fallback (address book not created): {clean(fb)}")
            for err in sec.get("errors", []):
                problem(f"{kind}: {clean(err)}")
            lines.extend(failed_item_lines(sec))
    if problems:
        lines.append(f"VERIFY FAILED: {problems} skipped/failed/mismatched item(s) listed above")
    elif listed_by_design:
        lines.append(f"All counts match; nothing failed. {listed_by_design} item(s) skipped by "
                     "design are listed above (empty system folders, calendars shared by others).")
    else:
        lines.append("All counts match; nothing skipped, nothing failed.")
    return lines, problems
