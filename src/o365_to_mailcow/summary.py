"""The web page's view model of a run report (``GET /api/overview``, job outcomes).

The page only renders: every judgement (what counts as missing, which level a mailbox gets,
what the headline sentence says) is made here, in plain functions without I/O, so it can be
unit-tested. ``summarize_report`` turns a report dict into the summary the page shows;
``summarize_with_step`` also returns the short entry of the overview's step strip, computed
from the same numbers instead of being parsed back out of the headline.

Verify outcomes come **only** from ``report.verify_summary``: after ``verify`` an entry's
``status: "ok"`` means the sections were collected, not that the counts match. The per-kind
numbers below explain *where* the problems are; whether there are any is the tool's call.

Partial runs (``--only``, ``--mailbox``, a per-run ``--mail-since``; recorded in the report's
``scope``, inferred from the kinds present for older reports) never claim more than they
checked: a clean partial run keeps a green headline that says what it covered, but its step
is ``unknown`` so it cannot stand in for a full run.

A migrate that skipped oversized messages or left something out (calendar ``warnings``,
contacts ``fallbacks``) is ``warn``, never green: verify would count the same things.

Construction is bounded, not only the output: at most ``MAX_TEXTS`` strings are converted
per report list, detail lists stop at ``MAX_DETAIL`` and count the rest, and mailboxes
beyond ``MAX_MAILBOXES`` keep only the numbers the headline needs.

Reports are read from disk, so every field is treated as untrusted: wrong types never raise,
every string is cleaned and cut to ``MAX_TEXT`` characters, counts above ``MAX_COUNT`` are
not believed, and data that cannot be read is reported as ``unknown`` — never as ``ok`` or
as "0 missing".
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from dataclasses import dataclass, field
from typing import Any

from .report import clean, verify_summary

MAX_TEXT = 300
MAX_MAILBOXES = 500
MAX_DETAIL = 200
MAX_COUNT = 10**15  # matches the 15-digit bound of the web server's progress parser
MAX_TEXTS = 1_000  # strings converted per report list; the rest is only counted
KINDS = ("mail", "calendar", "contacts")
TITLES = {"mail": "Mail", "calendar": "Calendar", "contacts": "Contacts"}
LEVELS = ("ok", "warn", "bad", "unknown")
_CHECKED = {"mail": ("folder", "folders"), "calendar": ("calendar", "calendars"),
            "contacts": ("address book", "address books")}
_DISPLAY = {"plan": "Plan", "provision": "Provision", "migrate": "Migrate",
            "verify": "Verify", "cleanup": "Clean up"}
_KIND_SCOPED = ("plan", "migrate", "verify")  # commands whose sections show the kinds run

JsonDict = dict[str, Any]


# -- reading untrusted values -------------------------------------------------------------

def _text(value: object) -> str:
    """Report-derived string for display: no control/bidi characters, at most 300 chars."""
    try:
        return clean(value)[:MAX_TEXT]
    except Exception:  # a str() that raises (exotic objects never come from JSON)
        return ""


def _opt_text(value: object) -> str | None:
    return None if value is None else _text(value)


def _exit_code(value: object) -> int | None:
    """A process exit code (-255..255); anything else, absurd integers included, is None."""
    if isinstance(value, int) and not isinstance(value, bool) and -255 <= value <= 255:
        return value
    return None


def _number(value: object) -> int | float | None:
    if isinstance(value, bool) or not isinstance(value, int | float):
        return None
    return value if value == value and abs(value) != float("inf") else None


def _n(value: int) -> str:
    return f"{value:,}"


def _plural(count: int, one: str, many: str) -> str:
    return f"{count:,} {one if count == 1 else many}"


def _items(count: int) -> str:
    return _plural(count, "item", "items")


def _items_are(count: int) -> str:
    return f"{count:,} item is" if count == 1 else f"{count:,} items are"


def _parts(items: Iterable[tuple[int, str]]) -> list[str]:
    """``(3, "missing"), (0, "failed")`` -> ``["3 missing"]`` (labels that do not inflect)."""
    return [f"{count:,} {label}" for count, label in items if count > 0]


class _Texts:
    """The first ``MAX_TEXTS`` strings of a report list, cleaned, plus the list's length:
    a report with millions of error strings costs a count, not millions of conversions."""

    def __init__(self, items: list[str], total: int) -> None:
        self.items = items
        self.total = total

    def __len__(self) -> int:
        return self.total

    def __iter__(self) -> Iterator[str]:
        return iter(self.items)


class _Capped:
    """A detail list that keeps at most ``MAX_DETAIL`` entries and counts the rest."""

    def __init__(self) -> None:
        self.items: list[Any] = []
        self.total = 0

    def add(self, item: Any) -> None:
        self.total += 1
        if len(self.items) < MAX_DETAIL:
            self.items.append(item)

    def add_texts(self, texts: _Texts, prefix: str = "") -> None:
        for text in texts:
            if len(self.items) >= MAX_DETAIL:
                break
            self.items.append(_text(prefix + text) if prefix else text)
        self.total += len(texts)

    def __len__(self) -> int:
        return self.total

    def __iter__(self) -> Iterator[Any]:
        return iter(self.items)

    @property
    def more(self) -> int:
        return self.total - len(self.items)


class _Reader:
    """Reads one section of a report; remembers whether anything had the wrong shape, so
    the caller can refuse to call data it could not read ``ok``."""

    def __init__(self) -> None:
        self.suspect = False

    def section(self, value: object) -> JsonDict:
        if isinstance(value, dict):
            return value
        self.suspect = True
        return {}

    def num(self, row: JsonDict, key: str, *, required: bool = True,
            signed: bool = False) -> int:
        """A whole number up to ``MAX_COUNT``; anything else is 0 and marks the data suspect.
        ``required=False``: an absent value (older reports) is a plain 0. ``signed``: may be
        negative (verify's ``expected`` after source deletions); other counts may not."""
        value = row.get(key)
        if value is None and not required:
            return 0
        if isinstance(value, bool) or not isinstance(value, int | float):
            self.suspect = True
            return 0
        if isinstance(value, float) and not value.is_integer():  # also nan and inf
            self.suspect = True
            return 0
        if abs(value) > MAX_COUNT or (value < 0 and not signed):
            self.suspect = True
            return 0
        return int(value)

    def flag(self, row: JsonDict, key: str) -> bool:
        value = row.get(key, False)
        if isinstance(value, bool):
            return value
        self.suspect = True
        return bool(value)

    def rows(self, section: JsonDict, key: str, *, required: bool = True) -> list[JsonDict]:
        value = section.get(key)
        if value is None and not required:
            return []
        if not isinstance(value, list):
            self.suspect = True
            return []
        rows = [row for row in value if isinstance(row, dict)]
        if len(rows) != len(value):
            self.suspect = True
        return rows

    def texts(self, section: JsonDict, key: str) -> _Texts:
        """A list of messages (errors, skipped names); a lone value is kept, but suspect."""
        value = section.get(key)
        if value is None:
            return _Texts([], 0)
        if not isinstance(value, list):
            self.suspect = True
            return _Texts([_text(value)], 1)
        return _Texts([_text(item) for item in value[:MAX_TEXTS]], len(value))


# -- building blocks of a mailbox row ------------------------------------------------------

@dataclass
class _Detail:
    cards: list[JsonDict] = field(default_factory=list)
    differences: _Capped = field(default_factory=_Capped)
    skipped: _Capped = field(default_factory=_Capped)
    sample: JsonDict | None = None
    problems: _Capped = field(default_factory=_Capped)
    errors: _Capped = field(default_factory=_Capped)

    def view(self) -> JsonDict:
        return {
            "cards": self.cards,
            "differences": self.differences.items,
            "more_differences": self.differences.more,
            "skipped": self.skipped.items,
            "more_skipped": self.skipped.more,
            "sample": self.sample,
            "problems": self.problems.items,
            "errors": self.errors.items,
        }


@dataclass
class _Mailbox:
    """One mailbox row plus the numbers the headline adds up."""

    key: str
    destination: str | None
    level: str = "unknown"
    text: str = "Unreadable entry"
    cells: dict[str, JsonDict | None] = field(
        default_factory=lambda: dict.fromkeys(KINDS))
    detail: _Detail = field(default_factory=_Detail)
    counts: dict[str, int] = field(default_factory=dict)
    first_error: str = ""

    def result(self, level: str, text: str) -> None:
        self.level, self.text = level, text

    def add(self, **counts: int) -> None:
        for name, value in counts.items():
            self.counts[name] = self.counts.get(name, 0) + value

    def view(self) -> JsonDict:
        return {"key": self.key, "destination": self.destination, "cells": self.cells,
                "result": {"level": self.level, "text": self.text},
                "detail": self.detail.view()}


def _total(boxes: list[_Mailbox], name: str) -> int:
    return sum(box.counts.get(name, 0) for box in boxes)


def _first_error(boxes: Iterable[_Mailbox]) -> str:
    return next((box.first_error for box in boxes if box.first_error), "")


def _entries(mailboxes: JsonDict) -> list[tuple[str, object]]:
    return sorted(((_text(key), entry) for key, entry in mailboxes.items()),
                  key=lambda item: item[0])


def _cell(primary: str, secondary: str, level: str) -> JsonDict:
    return {"primary": primary, "secondary": secondary, "level": level}


def _card(kind: str, rows: list[list[str]], level: str, verdict: str, note: str = "") -> JsonDict:
    return {"title": TITLES[kind], "rows": rows, "verdict": {"level": level, "text": verdict},
            "note": note}


# -- verify ----------------------------------------------------------------------------------

def _verify_kind(box: _Mailbox, kind: str, value: object) -> bool:
    """Fill ``box`` with one verify section; returns whether its data was readable."""
    reader = _Reader()
    sec = reader.section(value)
    title = TITLES[kind]
    source = found = missing = extra = failed = too_large = in_skipped = surplus = 0
    rows = reader.rows(sec, "folders" if kind == "mail" else "collections")
    for row in rows:
        if kind == "mail":
            row_source, row_found = reader.num(row, "graph_total"), reader.num(row, "imap_count")
            row_large = reader.num(row, "skipped")
            surplus += reader.num(row, "surplus_total", required=False)
            name, why = row.get("dest_name"), row.get("note")
        else:
            row_source, row_found = reader.num(row, "graph_count"), reader.num(row, "dav_count")
            row_large = 0
            name, why = row.get("name"), None
        # expected = source - skipped - failed: negative when those items were deleted at the
        # source since; only the derived missing/extra are clamped
        expected = reader.num(row, "expected", signed=True)
        row_failed = reader.num(row, "failed")
        mismatch = reader.flag(row, "mismatch")
        if mismatch != (row_found != expected):
            # the tool derives the flag from these two counts; a row where they disagree was
            # not written by it, and verify_summary would trust the flag
            reader.suspect = True
        row_missing, row_extra = max(expected - row_found, 0), max(row_found - expected, 0)
        source, found, failed = source + row_source, found + row_found, failed + row_failed
        missing, extra, too_large = missing + row_missing, extra + row_extra, too_large + row_large
        if mismatch or row_failed or row_large:
            delta = _parts(((row_missing, "missing"), (row_failed, "failed"),
                            (row_large, "too large"), (row_extra, "more than expected")))
            box.detail.differences.add({
                "kind": title, "name": _text(name), "source": _n(row_source),
                "destination": _n(row_found), "delta": " · ".join(delta) or "count mismatch",
                "why": _text(why) if why else ""})
    notes = [_plural(len(rows), *_CHECKED[kind]) + " checked"]
    sample_mismatches = 0
    if kind == "mail":
        for row in reader.rows(sec, "skipped_folders", required=False):
            total = reader.num(row, "total")
            in_skipped += total
            box.detail.skipped.add({
                "kind": title, "name": _text(row.get("path")), "reason": _text(row.get("reason")),
                "items": _n(total), "level": "warn" if total else "ok"})
        if surplus:
            notes.append(_plural(surplus, "surplus copy", "surplus copies"))
        box.detail.sample = _sample(reader, sec)
        if box.detail.sample is not None:  # verify_summary counts each one as a problem
            sample_mismatches = len(reader.texts(sec, "sample_mismatches"))
    else:
        by_design = reader.texts(sec, "skipped")  # shared calendars: not copied by design
        for name in by_design:
            box.detail.skipped.add({
                "kind": title, "name": name, "reason": "skipped by design (not owned by "
                "this mailbox)", "items": "", "level": "ok"})
        box.detail.skipped.total += len(by_design) - len(by_design.items)
        if kind == "calendar" and by_design:
            notes.append(_plural(len(by_design), "shared calendar", "shared calendars")
                         + " not copied")
    errors = reader.texts(sec, "errors")
    fallbacks = reader.texts(sec, "fallbacks")
    box.detail.errors.add_texts(errors, f"{kind}: ")
    parts = _parts(((missing, "missing"), (failed, "failed"), (too_large, "too large"),
                    (in_skipped, "in skipped folders"), (extra, "more than expected")))
    if sample_mismatches:
        parts.append(_plural(sample_mismatches, "sample mismatch", "sample mismatches"))
    if errors or fallbacks:
        parts.append(_plural(len(errors) + len(fallbacks), "error", "errors"))
    if reader.suspect:
        level, secondary = "unknown", " · ".join(parts) or "some counts could not be read"
    else:
        level, secondary = ("warn", " · ".join(parts)) if parts else ("ok", "complete")
    box.cells[kind] = _cell(f"{_n(found)} of {_n(source)}", secondary, level)
    box.detail.cards.append(_card(
        kind, [["In Microsoft 365", _n(source)], ["In mailcow", _n(found)]], level,
        "Complete" if level == "ok" else secondary, " · ".join(notes)))
    box.add(source=source, found=found, missing=missing, extra=extra, failed=failed,
            too_large=too_large, in_skipped=in_skipped)
    return not reader.suspect


def _sample(reader: _Reader, mail: JsonDict) -> JsonDict | None:
    if not reader.num(mail, "sample_requested", required=False):
        return None
    checked = reader.num(mail, "sample_checked")
    mismatches = len(reader.texts(mail, "sample_mismatches"))
    regenerated = reader.num(mail, "sample_regenerated", required=False)
    unverifiable = reader.num(mail, "sample_unverifiable", required=False)
    notes = []
    if mismatches:
        notes.append(_plural(mismatches, "message differs", "messages differ"))
    if regenerated:
        notes.append(f"{_n(regenerated)} re-rendered by Exchange (same headers)")
    if unverifiable:
        notes.append(f"{_n(unverifiable)} without Message-ID or gone at the source")
    return {"level": "warn" if mismatches else "ok",
            "text": f"{_n(max(checked - mismatches, 0))} of {_n(checked)} sampled "
                    + ("message matches" if checked == 1 else "messages match"),
            "detail": (", ".join(notes) + ".") if notes else ""}


def _verify_box(key: str, entry: object) -> _Mailbox:
    if not isinstance(entry, dict):
        box = _Mailbox(key, None)
        box.add(unknown=1)
        return box
    box = _Mailbox(key, _opt_text(entry.get("destination")))
    readable = True
    errors = _Reader()
    box.detail.errors.add_texts(errors.texts(entry, "errors"))
    ran = [kind for kind in KINDS if entry.get(kind) is not None]
    for kind in ran:
        readable = _verify_kind(box, kind, entry.get(kind)) and readable
    box.first_error = next((e for e in box.detail.errors if e), "")
    try:
        lines, problems = verify_summary({key: entry})
        for line in lines:  # its own list is the CLI's cost; ours stops at the cap
            if isinstance(line, str) and line.lstrip().startswith("!"):
                if len(box.detail.problems.items) < MAX_DETAIL:
                    box.detail.problems.add(_text(line.strip()[1:].strip()))
                else:
                    box.detail.problems.total += 1
    except Exception:  # malformed rows: verify_summary indexes keys directly
        box.result("unknown", "Results could not be read")
        box.add(unknown=1)
        return box
    problems = problems if isinstance(problems, int) and problems >= 0 else 0
    box.add(problems=problems)
    status = entry.get("status")
    if status == "missing":
        box.result("bad", "No mailbox in mailcow")
    elif status in ("failed", "pending"):
        box.result("bad", "Did not finish")
    elif problems == 0 and (not readable or errors.suspect or status != "ok" or not ran):
        box.result("unknown", "Results could not be read" if ran else "Nothing was checked")
    elif problems == 0:
        box.result("ok", "Complete")
    elif box.counts.get("missing", 0):
        box.result("warn", f"{_n(box.counts['missing'])} missing")
    else:
        box.result("warn", _plural(problems, "problem", "problems"))
    box.add(bad=box.level == "bad", unknown=box.level == "unknown",
            skipped_rows=len(box.detail.skipped))
    return box


def _verify(boxes: list[_Mailbox], exit_code: int | None) -> tuple[JsonDict, str]:
    found, source_n = _n(_total(boxes, "found")), _total(boxes, "source")
    missing, problems = _total(boxes, "missing"), _total(boxes, "problems")
    if not boxes:
        return _headline("unknown", "Verify checked no mailboxes"), "Nothing checked"
    if exit_code == 2 or _total(boxes, "bad"):
        return _headline("bad", "Verify could not finish", _first_error(boxes)), \
            "Could not finish"
    if problems == 0 and exit_code == 0 and not _total(boxes, "unknown"):
        detail = f"{found} of {_items_are(source_n)} in mailcow."
        if _total(boxes, "skipped_rows"):
            detail += " Items skipped by design are listed per mailbox."
        return _headline("ok", "Everything arrived", detail), "Everything arrived"
    detail = f"{found} of {_items_are(source_n)} at the destination."
    others = _parts(((_total(boxes, "failed"), "failed"), (_total(boxes, "too_large"), "too large"),
                     (_total(boxes, "in_skipped"), "in skipped folders"),
                     (_total(boxes, "extra"), "more than expected")))
    if others:
        detail += " " + ", ".join(others) + "."
    if missing:
        text = (f"{_n(missing)} item{' has' if missing == 1 else 's have'} not arrived in "
                "mailcow")
        return _headline("warn", text, detail), _plural(missing, "item missing", "items missing")
    if problems:
        return (_headline("warn", f"Verify found {_plural(problems, 'problem', 'problems')}",
                          detail), _plural(problems, "problem", "problems"))
    return _ended("Verify", exit_code)


# -- migrate ---------------------------------------------------------------------------------

def _migrate_kind(box: _Mailbox, kind: str, value: object, dry_run: bool) -> bool:
    reader = _Reader()
    sec = reader.section(value)
    title = TITLES[kind]
    copied = already = failed = too_large = would = 0
    for row in reader.rows(sec, "folders" if kind == "mail" else "collections"):
        if kind == "mail":
            row_copied = reader.num(row, "appended", required=not dry_run)
            already += (reader.num(row, "already_done", required=False)
                        + reader.num(row, "dedup_hits", required=False))
            too_large += reader.num(row, "skipped_too_large", required=False)
            would += reader.num(row, "would_append", required=dry_run)
            name = row.get("dest_name")
        else:
            row_copied = reader.num(row, "put", required=not dry_run)
            already += reader.num(row, "unchanged", required=False)
            would += reader.num(row, "would_put", required=dry_run)
            name = row.get("name")
        copied += row_copied
        error = _text(row.get("error")) if row.get("error") else ""
        row_failed = reader.num(row, "failed") + (1 if error else 0)
        failed += row_failed
        if row_failed:
            box.detail.differences.add({
                "kind": title, "name": _text(name), "source": "", "destination": "",
                "delta": f"{_n(row_failed)} failed", "why": error})
    errors = reader.texts(sec, "errors")
    box.detail.errors.add_texts(errors, f"{kind}: ")
    # known omissions verify would count as problems: an event's recurrence exception or a
    # contact photo left out (warnings), an address book refused and redirected (fallbacks)
    warning_texts, fallback_texts = reader.texts(sec, "warnings"), reader.texts(sec, "fallbacks")
    box.detail.problems.add_texts(warning_texts, f"{kind}: ")
    box.detail.problems.add_texts(fallback_texts, f"{kind}: ")
    mail_errors = 0
    if kind == "mail":
        skipped = reader.texts(sec, "skipped_folders")
        for name in skipped:
            box.detail.skipped.add({"kind": title, "name": name, "reason": "", "items": "",
                                    "level": "ok"})
        box.detail.skipped.total += len(skipped) - len(skipped.items)
        if sec.get("stopped") is True:
            box.add(stopped=1)
        mail_errors = len(errors)  # MailResult keeps errors apart from the failed count
    else:  # like CollectionsResult.failed: section errors count as failures
        failed += len(errors)
    warnings = len(warning_texts) + len(fallback_texts)
    if dry_run:
        primary = f"{_n(would)} to copy"
        parts = _parts(((already, "already there"), (too_large, "too large")))
        rows = [["To copy", _n(would)], ["Already there", _n(already)]]
    else:
        primary = f"{_n(copied)} copied"
        parts = _parts(((already, "already there"), (failed, "failed"),
                        (too_large, "too large")))
        rows = [["Copied", _n(copied)], ["Already there", _n(already)], ["Failed", _n(failed)]]
    if kind == "mail":
        rows.append(["Too large", _n(too_large)])
    if warnings:
        parts.append(_plural(warnings, "warning", "warnings"))
        rows.append(["Warnings", _n(warnings)])
    if mail_errors:
        parts.append(_plural(mail_errors, "error", "errors"))
        rows.append(["Errors", _n(mail_errors)])
    if failed or errors:
        level = "bad"
    elif too_large or warnings:
        level = "warn"
    else:
        level = "unknown" if reader.suspect else "ok"
    secondary = " · ".join(parts)
    box.cells[kind] = _cell(primary, secondary, level)
    verdict = secondary or {"bad": "Failed", "unknown": "Some counts could not be read",
                            "ok": "Nothing to copy" if dry_run else "Copied"}[level]
    box.detail.cards.append(_card(kind, rows, level, verdict))
    box.add(copied=copied, already=already, failed=failed, too_large=too_large, would=would,
            warnings=warnings)
    return not reader.suspect


def _to_check(box: _Mailbox) -> list[str]:
    return [*_parts(((box.counts.get("too_large", 0), "too large"),)),
            *([_plural(box.counts["warnings"], "warning", "warnings")]
              if box.counts.get("warnings") else [])]


def _migrate_box(key: str, entry: object, dry_run: bool) -> _Mailbox:
    if not isinstance(entry, dict):
        box = _Mailbox(key, None)
        box.add(unknown=1)
        return box
    box = _Mailbox(key, _opt_text(entry.get("destination")))
    errors = _Reader()
    box.detail.errors.add_texts(errors.texts(entry, "errors"))
    ran = [kind for kind in KINDS if entry.get(kind) is not None]
    readable = not errors.suspect
    for kind in ran:
        readable = _migrate_kind(box, kind, entry.get(kind), dry_run) and readable
    box.first_error = next((e for e in box.detail.errors if e), "")
    status, failed = entry.get("status"), box.counts.get("failed", 0)
    if status == "missing":
        box.result("bad", "No mailbox in mailcow")
    elif status == "pending":
        box.result("bad", "Did not finish")
    elif box.counts.get("stopped"):
        box.result("bad", "Stopped early")
    elif status == "failed" or failed or box.detail.errors:
        box.result("bad", f"{_n(failed)} failed" if failed else "Failed")
    elif not readable or status != "ok" or not ran:
        box.result("unknown", "Results could not be read" if ran else "Nothing was copied")
    elif _to_check(box):
        box.result("warn", " · ".join(_to_check(box)))
    else:
        box.result("ok", "Dry run" if dry_run else "Copied")
    box.add(bad=box.level == "bad", unknown=box.level == "unknown",
            incomplete=box.text in ("No mailbox in mailcow", "Did not finish", "Stopped early"))
    return box


def _migrate(boxes: list[_Mailbox], exit_code: int | None,
             dry_run: bool) -> tuple[JsonDict, str]:
    copied, already = _total(boxes, "copied"), _total(boxes, "already")
    too_large, warnings = _total(boxes, "too_large"), _total(boxes, "warnings")
    would = _total(boxes, "would")
    if not boxes:
        return _ended("Migrate", exit_code)
    if exit_code == 2 or _total(boxes, "bad"):
        failed = _total(boxes, "failed")
        if dry_run:
            detail = (f"Dry run: {_items(would)} would be copied, {_n(failed)} failed.")
        else:
            detail = (f"{_items(copied)} copied, {_n(already)} already there, "
                      f"{_n(failed)} failed.")
        if exit_code == 2 or _total(boxes, "incomplete"):
            return _headline("bad", "Migrate did not finish", detail), "Did not finish"
        bad = _plural(_total(boxes, "bad"), "mailbox", "mailboxes")
        return (_headline("bad", f"Migrate finished with failures in {bad}", detail),
                f"Failures in {bad}")
    if exit_code != 0 or _total(boxes, "unknown"):
        return _ended("Migrate", exit_code)
    if dry_run:
        extra = _parts(((too_large, "too large to copy"),))
        if warnings:
            extra.append(_plural(warnings, "warning", "warnings"))
        text = f"Dry run: {_items(would)} would be copied"
        if extra:
            return _headline("warn", f"{text}, {', '.join(extra)}"), f"{_n(would)} to copy"
        return _headline("ok", text), f"{_n(would)} to copy"
    detail = f"{_n(already)} {'was' if already == 1 else 'were'} already there."
    if too_large:
        detail += f" {_n(too_large)} too large to copy."
    if warnings:
        detail += f" {_plural(warnings, 'warning', 'warnings')}; see the mailbox details."
    if too_large or warnings:
        return (_headline("warn", f"Migrate finished with warnings: {_items(copied)} copied",
                          detail), f"{_n(copied)} copied · {_n(too_large + warnings)} to check")
    return (_headline("ok", f"Migrate finished: {_items(copied)} copied", detail),
            f"{_items(copied)} copied")


# -- plan ------------------------------------------------------------------------------------

def _plan_box(key: str, entry: object) -> _Mailbox:
    if not isinstance(entry, dict):
        box = _Mailbox(key, None)
        box.add(unknown=1)
        return box
    box = _Mailbox(key, _opt_text(entry.get("destination")))
    reader = _Reader()
    box.detail.errors.add_texts(reader.texts(entry, "errors"))
    planned = 0
    for kind in KINDS:
        value = entry.get(f"{kind}_plan")
        if value is None:
            continue
        sec = reader.section(value)
        if kind == "mail":
            count = reader.num(sec, "total_messages")
        else:
            count = sum(reader.num(c, "count") for c in reader.rows(sec, "collections")
                        if not reader.flag(c, "skip"))
        planned += 1
        box.cells[kind] = _cell(f"{_n(count)} to copy", "",
                                "unknown" if reader.suspect else "ok")
        box.detail.cards.append(_card(kind, [["Items", _n(count)]],
                                      "unknown" if reader.suspect else "ok",
                                      "Ready" if not reader.suspect else "Unreadable"))
        box.add(items=count)
    box.first_error = next((e for e in box.detail.errors if e), "")
    status = entry.get("status")
    if status == "missing":
        box.result("bad", "No mailbox in mailcow")
        box.add(missing_box=1)
    elif status in ("failed", "pending") or box.detail.errors:
        box.result("bad", "Failed")
        box.add(bad=1)
    elif reader.suspect or status != "ok" or not planned:
        box.result("unknown", "Results could not be read")
        box.add(unknown=1)
    else:
        box.result("ok", "Ready")
    return box


def _plan(boxes: list[_Mailbox], exit_code: int | None) -> tuple[JsonDict, str]:
    missing = _total(boxes, "missing_box")
    if not boxes:
        return _ended("Plan", exit_code)
    if exit_code == 2 or _total(boxes, "bad"):
        first = _first_error(b for b in boxes if b.counts.get("bad")) or _first_error(boxes)
        return _headline("bad", "Plan could not finish", first), "Could not finish"
    if missing:
        text = f"{_plural(missing, 'mailbox does', 'mailboxes do')} not exist in mailcow yet"
        return (_headline("warn", text, "Run Provision to create them."),
                f"{_n(missing)} missing in mailcow")
    if exit_code != 0 or _total(boxes, "unknown"):
        return _ended("Plan", exit_code)
    items = _total(boxes, "items")
    return (_headline("ok", f"Plan: {_items(items)} in "
                            f"{_plural(len(boxes), 'mailbox', 'mailboxes')}"),
            _plural(items, "item to copy", "items to copy"))


# -- provision and cleanup -----------------------------------------------------------------

_PROVISIONED = {"created": "Created", "exists": "Already exists",
                "would create": "Would be created"}


def _provision_box(key: str, entry: object) -> _Mailbox:
    box = _Mailbox(key, None)
    if not isinstance(entry, dict):
        box.add(unknown=1)
        return box
    reader = _Reader()
    errors = reader.texts(entry, "errors")
    box.detail.errors.add_texts(errors)
    box.first_error = next((e for e in errors if e), "")
    outcome = entry.get("provision")
    # only when every error was looked at: beyond the first MAX_TEXTS the kind is unknown
    alias_only = (bool(errors) and len(errors) == len(errors.items)
                  and all(e.startswith("alias ") for e in errors))
    if outcome not in _PROVISIONED or (errors and not alias_only) or reader.suspect:
        box.result("bad", "Failed")
        box.add(bad=1)
    elif alias_only:  # the mailbox is there; only some of its extra addresses failed
        box.result("warn", f"{_PROVISIONED[outcome]} · alias failed")
        box.add(alias_problems=len(errors), created=outcome != "exists",
                existed=outcome == "exists")
    else:
        box.result("ok", _PROVISIONED[outcome])
        box.add(created=outcome != "exists", existed=outcome == "exists")
    return box


@dataclass
class _AliasNote:
    """The ``"aliases"`` pseudo-entry of a provision report."""

    note: str
    error: str | None = None  # the aliases could not be listed or created
    unreadable: bool = False


def _aliases_note(entry: object) -> _AliasNote:
    reader = _Reader()
    sec = reader.section(entry)
    errors = [e for e in reader.texts(sec, "errors") if e]
    if errors:
        return _AliasNote(f"Aliases: {errors[0]}", error=errors[0])
    for key, label in (("created", "Aliases created"),
                       ("would_create", "Aliases that would be created")):
        if key in sec:
            count = reader.num(sec, key)
            if reader.suspect:
                break
            return _AliasNote(f"{label}: {_n(count)}")
    return _AliasNote("Aliases: the result could not be read", unreadable=True)


def _provision(boxes: list[_Mailbox], exit_code: int | None, dry_run: bool,
               aliases: _AliasNote | None) -> tuple[JsonDict, str]:
    failed, unknown = _total(boxes, "bad"), _total(boxes, "unknown")
    alias_problems = _total(boxes, "alias_problems")
    created, existed = _n(_total(boxes, "created")), _n(_total(boxes, "existed"))
    if failed:
        return (_headline("bad", f"Provision failed for "
                                 f"{_plural(failed, 'mailbox', 'mailboxes')}",
                          _first_error(b for b in boxes if b.counts.get("bad"))),
                f"{_n(failed)} failed")
    if aliases is not None and aliases.error is not None:
        return _headline("bad", "Provision could not create the aliases", aliases.error), \
            "Aliases failed"
    if exit_code == 2:
        return _headline("bad", "Provision could not finish"), "Could not finish"
    if alias_problems:
        return (_headline("warn", "Provision finished, "
                                  + _plural(alias_problems, "alias problem", "alias problems"),
                          _first_error(boxes)),
                _plural(alias_problems, "alias problem", "alias problems"))
    if exit_code != 0 or unknown or not boxes or (aliases is not None and aliases.unreadable):
        return _ended("Provision", exit_code)
    if dry_run:
        return (_headline("ok", f"Dry run: {created} would be created, {existed} already "
                                "exist"), f"{created} would be created")
    return (_headline("ok", f"Provision: {created} created, {existed} already existed"),
            f"{created} created, {existed} existed")


def _cleanup_box(key: str, entry: object, dry_run: bool) -> _Mailbox:
    box = _Mailbox(key, key)  # cleanup reports are keyed by destination address
    if not isinstance(entry, dict):
        box.add(unknown=1)
        return box
    reader = _Reader()
    errors = reader.texts(entry, "errors")
    box.detail.errors.add_texts(errors)
    box.first_error = next((e for e in errors if e), "")
    deleted = reader.num(entry, "deleted_app_passwords", required=bool(not errors))
    if errors:
        box.result("bad", "Failed")
        box.add(bad=1)
    elif reader.suspect:
        box.add(unknown=1)
    elif dry_run:  # a dry run lists what it would delete and records 0 deletions
        box.result("ok", "Dry run")
    else:
        box.result("ok", _plural(deleted, "app password", "app passwords") + " deleted")
        box.add(deleted=deleted)
    return box


def _cleanup(boxes: list[_Mailbox], exit_code: int | None,
             dry_run: bool) -> tuple[JsonDict, str]:
    failed = _total(boxes, "bad")
    if failed:
        return (_headline("bad", f"Clean up failed for "
                                 f"{_plural(failed, 'mailbox', 'mailboxes')}",
                          _first_error(boxes)), f"{_n(failed)} failed")
    if exit_code == 2:
        return _headline("bad", "Clean up could not finish"), "Could not finish"
    if exit_code != 0 or _total(boxes, "unknown"):
        return _ended("Clean up", exit_code)
    if dry_run:
        return _headline("ok", "Dry run: no app passwords were deleted"), "Dry run"
    deleted = _total(boxes, "deleted")
    return (_headline("ok", "Clean up: "
                            + _plural(deleted, "app password", "app passwords") + " deleted"),
            f"{_n(deleted)} deleted")


# -- scope (partial runs) ------------------------------------------------------------------

@dataclass(frozen=True)
class _Scope:
    partial: bool = False
    text: str = ""
    unreadable: bool = False  # recorded, but not in a form this version understands


_SCOPE_KEYS = frozenset({"only", "mailbox", "mail_since"})


def _scope(report: JsonDict, command: str, columns: list[str]) -> _Scope:
    """What the run covered, from the report's ``scope``; for reports written before it
    existed, inferred from the kinds present (plan/migrate/verify only)."""
    if "scope" not in report:
        if command in _KIND_SCOPED and 0 < len(columns) < len(KINDS):
            return _Scope(True, " and ".join(columns) + " only")
        return _Scope()
    scope = report["scope"]
    unreadable = _Scope(True, "the scope could not be read", unreadable=True)
    if not isinstance(scope, dict):
        return unreadable
    if any(value for key, value in scope.items() if key not in _SCOPE_KEYS):
        return unreadable  # a narrowing this version does not know
    only, mailbox, since = scope.get("only"), scope.get("mailbox"), scope.get("mail_since")
    # the CLI writes None or a string for each; anything else, falsy or not, is not its work
    if ((only is not None and only not in KINDS)
            or (mailbox is not None and not isinstance(mailbox, str))
            or (since is not None and not isinstance(since, str))):
        return unreadable
    parts = []
    if only:
        parts.append(f"{only} only")
    if mailbox:
        parts.append("one mailbox address")  # --mailbox can select more than one mapping
    if since:
        parts.append(f"mail since {_text(since)}")
    return _Scope(bool(parts), " · ".join(parts))


def _narrow(command: str, headline: JsonDict, step_text: str,
            scope: str) -> tuple[JsonDict, str, str]:
    """A clean partial run: say what was covered; its step stays grey (``unknown``)."""
    if command == "verify":
        detail = f"{headline['detail']} Only part was checked: {scope}.".strip()
        return (_headline("ok", "Everything that was checked has arrived", detail),
                "Partial check passed", "unknown")
    detail = f"{headline['detail']} Partial run: {scope}.".strip()
    return _headline("ok", headline["text"], detail), "Partial run", "unknown"


# -- assembly --------------------------------------------------------------------------------

def _headline(level: str, text: str, detail: str = "") -> JsonDict:
    return {"level": level, "text": text, "detail": detail}


def _ended(name: str, exit_code: int | None) -> tuple[JsonDict, str]:
    """The fallback when the data does not say what happened: grey, with the exit code."""
    if exit_code is None:
        return _headline("unknown", f"{name} ended"), "Unknown"
    return (_headline("unknown", f"{name} ended with exit code {exit_code}"),
            f"Exit code {exit_code}")


def _base(report: JsonDict) -> JsonDict:
    command = report.get("command")
    return {"command": _text(command) if isinstance(command, str) else "",
            "dry_run": report.get("dry_run") is True,
            "started": _opt_text(report.get("started")),
            "finished": _opt_text(report.get("finished")),
            "duration_s": _number(report.get("duration_s")),
            "exit_code": _exit_code(report.get("exit_code")),
            "partial": False, "scope": ""}


def _empty(base: JsonDict, headline: JsonDict) -> JsonDict:
    return {**base, "headline": headline, "columns": [], "mailboxes": [], "more_mailboxes": 0,
            "notes": []}


def _unknown(base: JsonDict) -> tuple[JsonDict, str, str]:
    headline, step = _ended(base["command"] or "Run", base["exit_code"])
    return _empty(base, headline), step, "unknown"


def unreadable_summary() -> JsonDict:
    """What the page shows when the newest report file cannot be read at all."""
    return _empty(_base({}), _headline("unknown", "The newest report could not be read"))


def _build(command: str, entries: list[tuple[str, object]],
           dry_run: bool) -> list[_Mailbox]:
    """One ``_Mailbox`` per entry. Rows beyond ``MAX_MAILBOXES`` keep their numbers for the
    headline but drop their detail at once, so a huge report never holds every detail."""
    boxes: list[_Mailbox] = []
    for key, entry in entries:
        if command == "verify":
            box = _verify_box(key, entry)
        elif command == "migrate":
            box = _migrate_box(key, entry, dry_run)
        elif command == "plan":
            box = _plan_box(key, entry)
        elif command == "provision":
            box = _provision_box(key, entry)
        else:
            box = _cleanup_box(key, entry, dry_run)
        if len(boxes) >= MAX_MAILBOXES:
            box.detail = _Detail()
        boxes.append(box)
    return boxes


def _summarize(report: object) -> tuple[JsonDict, str, str]:
    """``(summary, step text, step level)``."""
    data = report if isinstance(report, dict) else {}
    base = _base(data)
    mailboxes = data.get("mailboxes")
    command, exit_code, dry_run = base["command"], base["exit_code"], base["dry_run"]
    if command in _DISPLAY and exit_code == 2 and not mailboxes:
        # nothing ran: sign-in or configuration failed, or the command refused its options
        return (_empty(base, _headline("bad", f"{_DISPLAY[command]} could not start",
                                       "The job log has the reason.")),
                "Could not start", "bad")
    if command not in _DISPLAY or not isinstance(mailboxes, dict):
        return _unknown(base)
    entries = _entries(mailboxes)
    notes: list[str] = []
    aliases = None
    if command == "provision":
        aliases = _aliases_note(mailboxes["aliases"]) if "aliases" in mailboxes else None
        if aliases is not None:
            notes.append(aliases.note)
        entries = [(key, entry) for key, entry in entries if key != "aliases"]
    boxes = _build(command, entries, dry_run)
    if command == "verify":
        headline, step = _verify(boxes, exit_code)
    elif command == "migrate":
        headline, step = _migrate(boxes, exit_code, dry_run)
    elif command == "plan":
        headline, step = _plan(boxes, exit_code)
    elif command == "provision":
        headline, step = _provision(boxes, exit_code, dry_run, aliases)
    else:
        headline, step = _cleanup(boxes, exit_code, dry_run)
    columns = [kind for kind in KINDS if any(box.cells[kind] is not None for box in boxes)]
    scope = _scope(data, command, columns)
    step_level = headline["level"]
    if scope.unreadable and headline["level"] == "ok":
        headline = _headline("unknown", f"{_DISPLAY[command]} finished, but what it covered "
                                        "could not be read", headline["detail"])
        step, step_level = "Scope unknown", "unknown"
    elif scope.partial and headline["level"] == "ok":
        headline, step, step_level = _narrow(command, headline, step, scope.text)
    summary = {**base, "partial": scope.partial, "scope": scope.text, "headline": headline,
               "columns": columns,
               "mailboxes": [box.view() for box in boxes[:MAX_MAILBOXES]],
               "more_mailboxes": max(len(boxes) - MAX_MAILBOXES, 0), "notes": notes}
    return summary, step, step_level


_LAST_RESORT_STEP = "Unknown"


def _last_resort() -> JsonDict:
    """Built from literals only, so the error path itself cannot fail."""
    return {"command": "", "dry_run": False, "started": None, "finished": None,
            "duration_s": None, "exit_code": None, "partial": False, "scope": "",
            "headline": {"level": "unknown", "text": "Run ended", "detail": ""},
            "columns": [], "mailboxes": [], "more_mailboxes": 0, "notes": []}


def summarize_with_step(report: object) -> tuple[JsonDict, JsonDict]:
    """``(summary, step)``: the summary of contract section 2 and the overview's step entry
    ``{started, finished, exit_code, level, text}``. Never raises."""
    try:
        summary, text, level = _summarize(report)
    except Exception:  # a shape no reader anticipated: grey, never a 500 or a green
        try:
            summary, text, level = _unknown(_base(report if isinstance(report, dict) else {}))
        except Exception:
            summary, text, level = _last_resort(), _LAST_RESORT_STEP, "unknown"
    step = {"started": summary["started"], "finished": summary["finished"],
            "exit_code": summary["exit_code"], "level": level, "text": text}
    return summary, step


def summarize_report(report: object) -> JsonDict:
    """The page's summary of one run report (contract section 2). Never raises."""
    return summarize_with_step(report)[0]
