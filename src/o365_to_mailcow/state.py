"""Persistent run state (sqlite) that makes every command idempotent.

Every write is its own transaction so an interrupted run can resume at the item after the
last one recorded. The file never contains credentials: identifiers, statuses and error
summaries only.
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
"""


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

    def _exec(self, sql: str, params: tuple = ()) -> sqlite3.Cursor:
        with self._lock:
            return self._conn.execute(sql, params)

    # -- messages ----------------------------------------------------------------------

    def message_status(self, mailbox: str, folder_id: str, graph_id: str) -> str | None:
        row = self._exec(
            "SELECT status FROM messages WHERE mailbox=? AND folder_id=? AND graph_id=?",
            (mailbox, folder_id, graph_id),
        ).fetchone()
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
        row = self._exec(
            "SELECT uidvalidity FROM folder_meta WHERE mailbox=? AND folder_id=? "
            "AND uidvalidity IS NOT NULL",
            (mailbox, folder_id),
        ).fetchone()
        if row:
            return row[0]
        row = self._exec(
            "SELECT uidvalidity FROM messages WHERE mailbox=? AND folder_id=? "
            "AND uidvalidity IS NOT NULL ORDER BY updated_at DESC LIMIT 1",
            (mailbox, folder_id),
        ).fetchone()
        return row[0] if row else None

    def message_uidvalidity(self, mailbox: str, folder_id: str, graph_id: str) -> int | None:
        row = self._exec(
            "SELECT uidvalidity FROM messages WHERE mailbox=? AND folder_id=? AND graph_id=?",
            (mailbox, folder_id, graph_id),
        ).fetchone()
        return row[0] if row else None

    def set_folder_meta(self, mailbox: str, folder_id: str, dest_name: str,
                        uidvalidity: int | None) -> None:
        self._exec(
            "INSERT OR REPLACE INTO folder_meta VALUES (?,?,?,?,?)",
            (mailbox, folder_id, dest_name, uidvalidity, time.time()),
        )

    def done_count_for_message_id(self, mailbox: str, folder_id: str, message_id: str) -> int:
        """How many source items with this Message-ID are already recorded done here."""
        row = self._exec(
            "SELECT COUNT(*) FROM messages WHERE mailbox=? AND folder_id=? "
            "AND message_id=? AND status=?",
            (mailbox, folder_id, message_id, STATUS_DONE),
        ).fetchone()
        return int(row[0]) if row else 0

    def message_counts_by_folder_id(self, mailbox: str) -> dict[str, dict[str, int]]:
        """{folder_id: {status: count}} for one mailbox (verify keys on the source folder)."""
        out: dict[str, dict[str, int]] = {}
        for folder_id, status, n in self._exec(
            "SELECT folder_id, status, COUNT(*) FROM messages WHERE mailbox=? "
            "GROUP BY folder_id, status",
            (mailbox,),
        ):
            out.setdefault(folder_id, {})[status] = n
        return out

    def done_message_ids(self, mailbox: str, folder_id: str) -> set[str]:
        return {
            r[0] for r in self._exec(
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
        return [(r[0], r[1]) for r in self._exec("SELECT mailbox, mailcow_id FROM app_passwords")]

    def forget_app_password(self, mailbox: str, mailcow_id: str) -> None:
        self._exec(
            "DELETE FROM app_passwords WHERE mailbox=? AND mailcow_id=?", (mailbox, mailcow_id)
        )

    def message_counts(self, mailbox: str) -> dict[str, dict[str, int]]:
        """{folder: {status: count}} for one mailbox."""
        out: dict[str, dict[str, int]] = {}
        for folder, status, n in self._exec(
            "SELECT folder, status, COUNT(*) FROM messages WHERE mailbox=? GROUP BY folder, status",
            (mailbox,),
        ):
            out.setdefault(folder, {})[status] = n
        return out

    def get_delta(self, mailbox: str, folder_id: str) -> str | None:
        row = self._exec(
            "SELECT delta_link FROM folder_delta WHERE mailbox=? AND folder_id=?",
            (mailbox, folder_id),
        ).fetchone()
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
        row = self._exec("SELECT value FROM kv WHERE mailbox=? AND key=?",
                         (mailbox, key)).fetchone()
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
            r[0]: r[1] for r in self._exec(
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
        row = self._exec(
            "SELECT last_modified FROM events "
            "WHERE mailbox=? AND calendar=? AND uid=? AND status=?",
            (mailbox, calendar, uid, STATUS_DONE),
        ).fetchone()
        return row[0] if row else None

    def mark_event(self, mailbox: str, calendar: str, uid: str, last_modified: str | None,
                   status: str, error: str | None = None) -> None:
        self._exec(
            "INSERT OR REPLACE INTO events VALUES (?,?,?,?,?,?,?)",
            (mailbox, calendar, uid, last_modified, status, _trim(error), time.time()),
        )

    def event_counts(self, mailbox: str) -> dict[str, dict[str, int]]:
        out: dict[str, dict[str, int]] = {}
        for cal, status, n in self._exec(
            "SELECT calendar, status, COUNT(*) FROM events "
            "WHERE mailbox=? GROUP BY calendar, status",
            (mailbox,),
        ):
            out.setdefault(cal, {})[status] = n
        return out

    # -- contacts ----------------------------------------------------------------------

    def contact_last_modified(self, mailbox: str, graph_id: str) -> str | None:
        row = self._exec(
            "SELECT last_modified FROM contacts WHERE mailbox=? AND graph_id=? AND status=?",
            (mailbox, graph_id, STATUS_DONE),
        ).fetchone()
        return row[0] if row else None

    def mark_contact(self, mailbox: str, graph_id: str, book: str, last_modified: str | None,
                     status: str, error: str | None = None) -> None:
        self._exec(
            "INSERT OR REPLACE INTO contacts VALUES (?,?,?,?,?,?,?)",
            (mailbox, graph_id, book, last_modified, status, _trim(error), time.time()),
        )

    def contact_counts(self, mailbox: str) -> dict[str, dict[str, int]]:
        out: dict[str, dict[str, int]] = {}
        for book, status, n in self._exec(
            "SELECT book, status, COUNT(*) FROM contacts WHERE mailbox=? GROUP BY book, status",
            (mailbox,),
        ):
            out.setdefault(book, {})[status] = n
        return out


def _trim(error: str | None) -> str | None:
    return None if error is None else error[:300]
