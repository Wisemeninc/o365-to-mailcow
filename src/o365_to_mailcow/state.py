"""Persistent run state (sqlite) that makes every command idempotent.

Every write is its own transaction so an interrupted run can resume at the item after the
last one recorded. The file never contains credentials: identifiers, statuses and error
summaries, plus the title, sender and date of items that failed (so a report can say which
item it was).
"""

from __future__ import annotations

import os
import sqlite3
import threading
import time
from pathlib import Path

STATUS_DONE = "done"
STATUS_FAILED = "failed"
STATUS_SKIPPED = "skipped"

_SCHEMA = """
CREATE TABLE IF NOT EXISTS messages (
    mailbox TEXT NOT NULL, folder_id TEXT NOT NULL, graph_id TEXT NOT NULL,
    folder TEXT NOT NULL, message_id TEXT, status TEXT NOT NULL, error TEXT,
    dest_uid INTEGER, uidvalidity INTEGER, updated_at REAL NOT NULL,
    PRIMARY KEY (mailbox, folder_id, graph_id));
CREATE TABLE IF NOT EXISTS app_passwords (
    mailbox TEXT NOT NULL, mailcow_id TEXT NOT NULL, created_at REAL NOT NULL,
    PRIMARY KEY (mailbox, mailcow_id));
CREATE TABLE IF NOT EXISTS folder_delta (
    mailbox TEXT NOT NULL, folder_id TEXT NOT NULL, delta_link TEXT NOT NULL,
    updated_at REAL NOT NULL, PRIMARY KEY (mailbox, folder_id));
CREATE TABLE IF NOT EXISTS folder_meta (
    mailbox TEXT NOT NULL, folder_id TEXT NOT NULL, dest_name TEXT NOT NULL,
    uidvalidity INTEGER, updated_at REAL NOT NULL, PRIMARY KEY (mailbox, folder_id));
CREATE TABLE IF NOT EXISTS kv (
    mailbox TEXT NOT NULL, key TEXT NOT NULL, value TEXT, PRIMARY KEY (mailbox, key));
CREATE TABLE IF NOT EXISTS collections (
    mailbox TEXT NOT NULL, kind TEXT NOT NULL, source_id TEXT NOT NULL, slug TEXT NOT NULL,
    updated_at REAL NOT NULL, PRIMARY KEY (mailbox, kind, source_id));
CREATE TABLE IF NOT EXISTS events (
    mailbox TEXT NOT NULL, calendar TEXT NOT NULL, uid TEXT NOT NULL,
    last_modified TEXT, status TEXT NOT NULL, error TEXT, updated_at REAL NOT NULL,
    PRIMARY KEY (mailbox, calendar, uid));
CREATE TABLE IF NOT EXISTS contacts (
    mailbox TEXT NOT NULL, graph_id TEXT NOT NULL, book TEXT NOT NULL,
    last_modified TEXT, status TEXT NOT NULL, error TEXT, updated_at REAL NOT NULL,
    PRIMARY KEY (mailbox, graph_id));
CREATE TABLE IF NOT EXISTS item_labels (
    mailbox TEXT NOT NULL, kind TEXT NOT NULL, collection TEXT NOT NULL, item TEXT NOT NULL,
    place TEXT, title TEXT, hint TEXT, updated_at REAL NOT NULL,
    PRIMARY KEY (mailbox, kind, collection, item));
"""

# Per kind: (delete labels of items no longer failed/skipped, count them, list them with
# their label). Contacts are keyed by graph_id alone, so their labels use an empty
# collection. 'failed'/'skipped' are STATUS_FAILED/STATUS_SKIPPED, spelled out because the
# statements are literals: no SQL is ever assembled from strings at run time.
_LABEL_SQL = {
    "mail": (
        "DELETE FROM item_labels WHERE mailbox=? AND kind='mail' AND NOT EXISTS ("
        "SELECT 1 FROM messages t WHERE t.mailbox=item_labels.mailbox "
        "AND t.folder_id=item_labels.collection AND t.graph_id=item_labels.item "
        "AND t.status IN ('failed','skipped'))",
        "SELECT COUNT(*) FROM messages WHERE mailbox=? "
        "AND status IN ('failed','skipped')",
        "SELECT COALESCE(l.place, t.folder), COALESCE(l.title, ''), COALESCE(l.hint, ''), "
        "t.status, COALESCE(t.error, '') FROM messages t LEFT JOIN item_labels l "
        "ON l.mailbox=t.mailbox AND l.kind='mail' AND l.collection=t.folder_id "
        "AND l.item=t.graph_id WHERE t.mailbox=? AND t.status IN ('failed','skipped') "
        "ORDER BY t.updated_at, t.folder_id, t.graph_id LIMIT ?"),
    "calendar": (
        "DELETE FROM item_labels WHERE mailbox=? AND kind='calendar' AND NOT EXISTS ("
        "SELECT 1 FROM events t WHERE t.mailbox=item_labels.mailbox "
        "AND t.calendar=item_labels.collection AND t.uid=item_labels.item "
        "AND t.status IN ('failed','skipped'))",
        "SELECT COUNT(*) FROM events WHERE mailbox=? "
        "AND status IN ('failed','skipped')",
        "SELECT COALESCE(l.place, t.calendar), COALESCE(l.title, ''), COALESCE(l.hint, ''), "
        "t.status, COALESCE(t.error, '') FROM events t LEFT JOIN item_labels l "
        "ON l.mailbox=t.mailbox AND l.kind='calendar' AND l.collection=t.calendar "
        "AND l.item=t.uid WHERE t.mailbox=? AND t.status IN ('failed','skipped') "
        "ORDER BY t.updated_at, t.calendar, t.uid LIMIT ?"),
    "contacts": (
        "DELETE FROM item_labels WHERE mailbox=? AND kind='contacts' AND NOT EXISTS ("
        "SELECT 1 FROM contacts t WHERE t.mailbox=item_labels.mailbox "
        "AND ''=item_labels.collection AND t.graph_id=item_labels.item "
        "AND t.status IN ('failed','skipped'))",
        "SELECT COUNT(*) FROM contacts WHERE mailbox=? "
        "AND status IN ('failed','skipped')",
        "SELECT COALESCE(l.place, t.book), COALESCE(l.title, ''), COALESCE(l.hint, ''), "
        "t.status, COALESCE(t.error, '') FROM contacts t LEFT JOIN item_labels l "
        "ON l.mailbox=t.mailbox AND l.kind='contacts' AND l.collection='' "
        "AND l.item=t.graph_id WHERE t.mailbox=? AND t.status IN ('failed','skipped') "
        "ORDER BY t.updated_at, t.graph_id LIMIT ?"),
}


class State:
    def __init__(self, path: Path | str) -> None:
        self._path = Path(path)
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        # SQLite creates the -wal/-shm side files with the process umask; pre-creating
        # them keeps every file of the state database at 0600 (ISC-130).
        for side in ("", "-wal", "-shm"):
            side_path = self._path.with_name(self._path.name + side)
            if not side_path.exists():
                side_path.touch(mode=0o600)
            os.chmod(side_path, 0o600)
        self._conn = sqlite3.connect(self._path, check_same_thread=False, isolation_level=None)
        self._conn.execute("PRAGMA journal_mode=WAL")
        self._conn.executescript(_SCHEMA)

    def close(self) -> None:
        self._conn.close()

    # One connection is shared by every mailbox thread. The sqlite3 module is not
    # thread-safe, so a statement is executed *and* its rows are read under the lock:
    # fetching from a cursor while another thread runs a statement on the same connection
    # returned garbage rows (a live run saw a COUNT(*) come back as None).

    def _exec(self, sql: str, params: tuple = ()) -> None:
        with self._lock:
            self._conn.execute(sql, params)

    def _rows(self, sql: str, params: tuple = ()) -> list[tuple]:
        with self._lock:
            return self._conn.execute(sql, params).fetchall()

    def _row(self, sql: str, params: tuple = ()) -> tuple | None:
        with self._lock:
            return self._conn.execute(sql, params).fetchone()

    # -- messages ----------------------------------------------------------------------

    def message_status(self, mailbox: str, folder_id: str, graph_id: str) -> str | None:
        row = self._row(
            "SELECT status FROM messages WHERE mailbox=? AND folder_id=? AND graph_id=?",
            (mailbox, folder_id, graph_id),
        )
        return row[0] if row else None

    def mark_message(self, mailbox: str, folder_id: str, graph_id: str, folder: str,
                     message_id: str | None, status: str, error: str | None = None,
                     dest_uid: int | None = None, uidvalidity: int | None = None) -> None:
        self._exec(
            "INSERT OR REPLACE INTO messages VALUES (?,?,?,?,?,?,?,?,?,?)",
            (mailbox, folder_id, graph_id, folder, message_id, status, _trim(error),
             dest_uid, uidvalidity, time.time()),
        )

    def folder_uidvalidity(self, mailbox: str, folder_id: str) -> int | None:
        """UIDVALIDITY last seen for a folder: folder_meta first, else the newest message."""
        row = self._row(
            "SELECT uidvalidity FROM folder_meta WHERE mailbox=? AND folder_id=? "
            "AND uidvalidity IS NOT NULL",
            (mailbox, folder_id),
        )
        if row:
            return row[0]
        row = self._row(
            "SELECT uidvalidity FROM messages WHERE mailbox=? AND folder_id=? "
            "AND uidvalidity IS NOT NULL ORDER BY updated_at DESC LIMIT 1",
            (mailbox, folder_id),
        )
        return row[0] if row else None

    def message_uidvalidity(self, mailbox: str, folder_id: str, graph_id: str) -> int | None:
        row = self._row(
            "SELECT uidvalidity FROM messages WHERE mailbox=? AND folder_id=? AND graph_id=?",
            (mailbox, folder_id, graph_id),
        )
        return row[0] if row else None

    def set_folder_meta(self, mailbox: str, folder_id: str, dest_name: str,
                        uidvalidity: int | None) -> None:
        self._exec(
            "INSERT OR REPLACE INTO folder_meta VALUES (?,?,?,?,?)",
            (mailbox, folder_id, dest_name, uidvalidity, time.time()),
        )

    def done_count_for_message_id(self, mailbox: str, folder_id: str, message_id: str) -> int:
        """How many source items with this Message-ID are already recorded done here."""
        row = self._row(
            "SELECT COUNT(*) FROM messages WHERE mailbox=? AND folder_id=? "
            "AND message_id=? AND status=?",
            (mailbox, folder_id, message_id, STATUS_DONE),
        )
        return int(row[0]) if row else 0

    def done_message_id_counts(self, mailbox: str, folder_id: str) -> dict[str, int]:
        """{Message-ID (lower-cased): number of source items recorded done} for a folder,
        with the destination UID of the newest row per Message-ID in ``_done_uid``."""
        out: dict[str, int] = {}
        for mid, n in self._rows(
            "SELECT LOWER(message_id), COUNT(*) FROM messages WHERE mailbox=? AND folder_id=? "
            "AND status=? AND message_id IS NOT NULL GROUP BY LOWER(message_id)",
            (mailbox, folder_id, STATUS_DONE),
        ):
            out[mid] = n
        return out

    def done_dest_uids(self, mailbox: str, folder_id: str) -> set[int]:
        """Destination UIDs this tool recorded for its own appends into a folder."""
        return {int(r[0]) for r in self._rows(
            "SELECT dest_uid FROM messages WHERE mailbox=? AND folder_id=? AND status=? "
            "AND dest_uid IS NOT NULL", (mailbox, folder_id, STATUS_DONE))}

    def message_counts_by_folder_id(self, mailbox: str) -> dict[str, dict[str, int]]:
        """{folder_id: {status: count}} for one mailbox (verify keys on the source folder)."""
        out: dict[str, dict[str, int]] = {}
        for folder_id, status, n in self._rows(
            "SELECT folder_id, status, COUNT(*) FROM messages WHERE mailbox=? "
            "GROUP BY folder_id, status",
            (mailbox,),
        ):
            out.setdefault(folder_id, {})[status] = n
        return out

    def done_message_ids(self, mailbox: str, folder_id: str) -> set[str]:
        return {
            r[0] for r in self._rows(
                "SELECT graph_id FROM messages WHERE mailbox=? AND folder_id=? AND status=?",
                (mailbox, folder_id, STATUS_DONE),
            )
        }

    # -- app passwords -----------------------------------------------------------------

    def record_app_password(self, mailbox: str, mailcow_id: str) -> None:
        self._exec(
            "INSERT OR REPLACE INTO app_passwords VALUES (?,?,?)",
            (mailbox, mailcow_id, time.time()),
        )

    def app_passwords(self) -> list[tuple[str, str]]:
        return [(r[0], r[1]) for r in self._rows("SELECT mailbox, mailcow_id FROM app_passwords")]

    def forget_app_password(self, mailbox: str, mailcow_id: str) -> None:
        self._exec(
            "DELETE FROM app_passwords WHERE mailbox=? AND mailcow_id=?", (mailbox, mailcow_id)
        )

    def message_counts(self, mailbox: str) -> dict[str, dict[str, int]]:
        """{folder: {status: count}} for one mailbox."""
        out: dict[str, dict[str, int]] = {}
        for folder, status, n in self._rows(
            "SELECT folder, status, COUNT(*) FROM messages WHERE mailbox=? GROUP BY folder, status",
            (mailbox,),
        ):
            out.setdefault(folder, {})[status] = n
        return out

    def get_delta(self, mailbox: str, folder_id: str) -> str | None:
        row = self._row(
            "SELECT delta_link FROM folder_delta WHERE mailbox=? AND folder_id=?",
            (mailbox, folder_id),
        )
        return row[0] if row else None

    def set_delta(self, mailbox: str, folder_id: str, delta_link: str) -> None:
        self._exec(
            "INSERT OR REPLACE INTO folder_delta VALUES (?,?,?,?)",
            (mailbox, folder_id, delta_link, time.time()),
        )

    def clear_all_deltas(self, mailbox: str) -> int:
        """Forget every delta link of a mailbox (the next pass lists all folders fully)."""
        with self._lock:
            cur = self._conn.execute("DELETE FROM folder_delta WHERE mailbox=?", (mailbox,))
            return cur.rowcount

    def get_kv(self, mailbox: str, key: str) -> str | None:
        row = self._row("SELECT value FROM kv WHERE mailbox=? AND key=?",
                         (mailbox, key))
        return row[0] if row else None

    def set_kv(self, mailbox: str, key: str, value: str | None) -> None:
        self._exec("INSERT OR REPLACE INTO kv VALUES (?,?,?)", (mailbox, key, value))

    def clear_delta(self, mailbox: str, folder_id: str) -> None:
        """Forget an expired delta link so the next pass lists the folder fully."""
        self._exec(
            "DELETE FROM folder_delta WHERE mailbox=? AND folder_id=?", (mailbox, folder_id)
        )

    # -- collection slugs (calendars, address books) -----------------------------------

    def collection_slugs(self, mailbox: str, kind: str) -> dict[str, str]:
        return {
            r[0]: r[1] for r in self._rows(
                "SELECT source_id, slug FROM collections WHERE mailbox=? AND kind=?",
                (mailbox, kind),
            )
        }

    def set_collection_slug(self, mailbox: str, kind: str, source_id: str, slug: str) -> None:
        self._exec(
            "INSERT OR REPLACE INTO collections VALUES (?,?,?,?,?)",
            (mailbox, kind, source_id, slug, time.time()),
        )

    # -- events ------------------------------------------------------------------------

    def event_last_modified(self, mailbox: str, calendar: str, uid: str) -> str | None:
        row = self._row(
            "SELECT last_modified FROM events "
            "WHERE mailbox=? AND calendar=? AND uid=? AND status=?",
            (mailbox, calendar, uid, STATUS_DONE),
        )
        return row[0] if row else None

    def mark_event(self, mailbox: str, calendar: str, uid: str, last_modified: str | None,
                   status: str, error: str | None = None) -> None:
        self._exec(
            "INSERT OR REPLACE INTO events VALUES (?,?,?,?,?,?,?)",
            (mailbox, calendar, uid, last_modified, status, _trim(error), time.time()),
        )

    def event_counts(self, mailbox: str) -> dict[str, dict[str, int]]:
        out: dict[str, dict[str, int]] = {}
        for cal, status, n in self._rows(
            "SELECT calendar, status, COUNT(*) FROM events "
            "WHERE mailbox=? GROUP BY calendar, status",
            (mailbox,),
        ):
            out.setdefault(cal, {})[status] = n
        return out

    # -- contacts ----------------------------------------------------------------------

    def contact_last_modified(self, mailbox: str, graph_id: str) -> str | None:
        row = self._row(
            "SELECT last_modified FROM contacts WHERE mailbox=? AND graph_id=? AND status=?",
            (mailbox, graph_id, STATUS_DONE),
        )
        return row[0] if row else None

    def mark_contact(self, mailbox: str, graph_id: str, book: str, last_modified: str | None,
                     status: str, error: str | None = None) -> None:
        self._exec(
            "INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?,?,?)",
            (mailbox, graph_id, book, last_modified, status, _trim(error), time.time()),
        )

    def contact_counts(self, mailbox: str) -> dict[str, dict[str, int]]:
        out: dict[str, dict[str, int]] = {}
        for book, status, n in self._rows(
            "SELECT book, status, COUNT(*) FROM contacts WHERE mailbox=? GROUP BY book, status",
            (mailbox,),
        ):
            out.setdefault(book, {})[status] = n
        return out

    # -- labels of failed items --------------------------------------------------------

    def set_label(self, mailbox: str, kind: str, collection: str, item: str,
                  place: str | None, title: str | None, hint: str | None) -> None:
        """Remember what a failed or skipped item was (where, title, sender/date)."""
        self._exec(
            "INSERT OR REPLACE INTO item_labels VALUES (?,?,?,?,?,?,?,?)",
            (mailbox, kind, collection, item, _trim_label(place), _trim_label(title),
             _trim_label(hint), time.time()),
        )

    def label_title(self, mailbox: str, kind: str, collection: str, item: str) -> str:
        """The title recorded for an item, "" when there is none (yet)."""
        row = self._row(
            "SELECT title FROM item_labels WHERE mailbox=? AND kind=? AND collection=? "
            "AND item=?", (mailbox, kind, collection, item))
        return row[0] if row and isinstance(row[0], str) else ""

    def failed_items(self, mailbox: str, kind: str,
                     limit: int = 100) -> tuple[list[dict[str, str]], int]:
        """The items of one kind whose current status is failed or skipped, oldest first,
        at most ``limit`` of them, plus how many there are in total. Labels of items that
        have since been copied are deleted first, so a title does not outlive its failure."""
        if kind not in _LABEL_SQL:
            raise ValueError(f"unknown item kind {kind!r}")
        prune, count, listing = _LABEL_SQL[kind]
        self._exec(prune, (mailbox,))
        row = self._row(count, (mailbox,))
        total = int(row[0]) if row else 0
        if limit <= 0:
            return [], total
        rows = self._rows(listing, (mailbox, limit))
        keys = ("place", "title", "hint", "status", "error")
        return [{k: str(v) for k, v in zip(keys, r, strict=True)} for r in rows], total


def _trim(error: str | None) -> str | None:
    return None if error is None else error[:300]


def _trim_label(text: str | None) -> str | None:
    """At most 300 characters, cut at a space: a label is shown after whole-secret
    redaction, so a cut must never leave the first half of a (space-free) secret behind."""
    if text is None or len(text) <= 300:
        return text
    head = text[:299]
    return head[:head.rfind(" ") + 1] + "…"
