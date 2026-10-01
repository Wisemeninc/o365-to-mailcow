"""Destination IMAP connection (Dovecot in mailcow), write-only in the append sense.

Only these IMAP commands are ever issued (ISC-54): LOGIN, LIST, CREATE, SUBSCRIBE,
STATUS, SELECT/EXAMINE, SEARCH, APPEND (one message, or several per command with
MULTIAPPEND), FETCH (Message-ID index and verify sampling, always BODY.PEEK), LOGOUT.
The class deliberately has no way to STORE flags, EXPUNGE or DELETE (ISC-55).

Folder names are passed to imapclient as ``str``; imapclient encodes them to modified
UTF-7 and decodes LIST results back (ISC-40).
"""

from __future__ import annotations

import logging
import re
import ssl
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
from typing import TypeVar

from imapclient import IMAPClient
from imapclient.exceptions import IMAPClientAbortError, IMAPClientError, LoginError
from imapclient.imapclient import SocketTimeout

log = logging.getLogger(__name__)

T = TypeVar("T")
_APPENDUID = re.compile(rb"\[APPENDUID (\d+) (\d+)\]", re.IGNORECASE)
_APPENDUID_SET = re.compile(rb"\[APPENDUID (\d+) ([0-9:,]+)\]", re.IGNORECASE)


@dataclass(frozen=True)
class AppendItem:
    """One message for ``append_many``."""

    mime: bytes
    flags: tuple[str, ...]
    internal_date: datetime
    message_id: str | None = None


def expand_uid_set(text: str) -> list[int]:
    """``"13:22,30"`` -> ``[13, 14, …, 22, 30]`` (RFC 4315 uid-set, ascending ranges)."""
    out: list[int] = []
    for part in text.split(","):
        if not part:
            continue
        lo, _, hi = part.partition(":")
        a = int(lo)
        b = int(hi) if hi else a
        out.extend(range(min(a, b), max(a, b) + 1))
    return out
# "Message-ID: <id>", the obsolete "Message-ID : <id>" and a comment before the id
_HEADER_MID = re.compile(rb"^message-id[ \t]*:[ \t]*(?:\([^)]*\)[ \t]*)*(<[^>]*>)",
                         re.IGNORECASE | re.MULTILINE)
INDEX_CHUNK = 2000  # messages per FETCH when indexing a folder's Message-IDs
CONNECTION_ERRORS = (IMAPClientAbortError, OSError)


class ImapError(Exception):
    """The server answered NO or BAD (for example OVERQUOTA). Not retried."""


def is_quota_error(exc: BaseException) -> bool:
    """True for the one NO that should stop a mailbox: the destination is out of space."""
    text = str(exc).upper()
    return "OVERQUOTA" in text or "QUOTA" in text


class ImapConnectionError(Exception):
    """The connection failed and one reconnect-and-retry did not help."""


AppendOutcome = int | None | ImapError  # per item of append_many


class ImapDestination:
    """One authenticated IMAP connection to the mailcow host for one mailbox."""

    def __init__(self, host: str, port: int, user: str, password: str,
                 verify: bool | str = True, client_factory: Callable[..., IMAPClient] = IMAPClient,
                 timeout: SocketTimeout | None = None) -> None:
        """``verify`` is True (system CAs) or the path of a private CA bundle."""
        self._host = host
        self._port = port
        self._user = user
        self._password = password
        self._verify = verify
        self._factory = client_factory
        self._timeout = timeout or SocketTimeout(connect=15.0, read=300.0)
        self._client: IMAPClient | None = None
        self._delimiter: str | None = None

    # -- connection --------------------------------------------------------------------

    def _ssl_context(self) -> ssl.SSLContext:
        cafile = self._verify if isinstance(self._verify, str) else None
        ctx = ssl.create_default_context(cafile=cafile)
        if self._verify is False:  # test hook only; production callers never pass False
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
        return ctx

    def connect(self) -> None:
        """Open the TLS connection and LOGIN. Raises ImapError on bad credentials."""
        try:
            client = self._factory(self._host, port=self._port, ssl=True,
                                   ssl_context=self._ssl_context(), timeout=self._timeout)
        except CONNECTION_ERRORS as exc:
            raise ImapConnectionError(
                f"cannot connect to {self._host}:{self._port}: {exc.__class__.__name__}"
            ) from exc
        try:
            client.login(self._user, self._password)
        except LoginError as exc:
            raise ImapError(f"IMAP login failed for {self._user}") from exc
        self._client = client

    def close(self) -> None:
        if self._client is None:
            return
        try:
            self._client.logout()
        except (IMAPClientError, OSError):  # already gone; nothing to clean up
            log.debug("IMAP logout failed for %s", self._user)
        self._client = None

    def __enter__(self) -> ImapDestination:
        self.connect()
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    def _call(self, fn: Callable[[IMAPClient], T], what: str) -> T:
        """Run ``fn`` with the client; on a dropped connection reconnect once and retry."""
        if self._client is None:
            self.connect()
        try:
            return fn(self._client)  # type: ignore[arg-type]
        except CONNECTION_ERRORS as exc:
            log.warning("IMAP connection lost during %s (%s); reconnecting once",
                        what, exc.__class__.__name__)
            self._client = None
            try:
                self.connect()
                return fn(self._client)  # type: ignore[arg-type]
            except CONNECTION_ERRORS as exc2:
                self._client = None
                raise ImapConnectionError(
                    f"IMAP {what} failed after reconnect: {exc2.__class__.__name__}"
                ) from exc2
            except IMAPClientError as exc2:
                raise ImapError(f"IMAP {what} refused: {_short(exc2)}") from exc2
        except IMAPClientError as exc:
            raise ImapError(f"IMAP {what} refused: {_short(exc)}") from exc

    # -- folders -----------------------------------------------------------------------

    @property
    def delimiter(self) -> str:
        """Hierarchy delimiter from ``LIST "" ""`` (RFC 3501 6.3.8); ``/`` if flat."""
        if self._delimiter is None:
            rows = self._call(lambda c: c.list_folders("", ""), "LIST")
            delim: object = rows[0][1] if rows else None
            if isinstance(delim, bytes):
                delim = delim.decode("ascii", "replace")
            self._delimiter = str(delim) if delim else "/"
        return self._delimiter

    def _exists(self, client: IMAPClient, name: str) -> bool:
        return any(row[2] == name for row in client.list_folders("", name))

    def ensure_folder(self, name: str) -> None:
        """CREATE the folder if LIST does not show it, then SUBSCRIBE. Idempotent."""
        def op(c: IMAPClient) -> None:
            if name.upper() != "INBOX" and not self._exists(c, name):
                c.create_folder(name)
            c.subscribe_folder(name)
        self._call(op, f"CREATE/SUBSCRIBE {name!r}")

    def folder_status(self, name: str) -> tuple[int, int]:
        """Return ``(MESSAGES, UIDVALIDITY)`` for a folder."""
        st = self._call(lambda c: c.folder_status(name, ["MESSAGES", "UIDVALIDITY"]),
                        f"STATUS {name!r}")
        return int(st.get(b"MESSAGES", 0)), int(st.get(b"UIDVALIDITY", 0))

    # -- messages ----------------------------------------------------------------------

    @staticmethod
    def _search(c: IMAPClient, folder: str, message_id: str) -> list[int]:
        c.select_folder(folder, readonly=True)
        return [int(u) for u in c.search(["HEADER", "Message-ID", message_id])]

    def search_message_id(self, folder: str, message_id: str) -> list[int]:
        """UIDs in ``folder`` whose Message-ID header matches (EXAMINE + SEARCH)."""
        return self._call(lambda c: self._search(c, folder, message_id), f"SEARCH {folder!r}")

    def has_message_id(self, folder: str, message_id: str) -> bool:
        return bool(self.search_message_id(folder, message_id))

    def message_id_index(self, folder: str,
                         on_progress: Callable[[int, int], None] | None = None,
                         ) -> dict[str, list[int]]:
        """Lower-cased Message-ID -> UIDs of every message in ``folder``: one EXAMINE and
        chunked ``FETCH <seq range> (UID BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)])`` calls by
        message sequence number. No ``SEARCH ALL``: its single-line reply exceeds
        imaplib's 1 MB line limit at roughly 150 000 UIDs. A resumed run checks each
        source message locally instead of running one SEARCH per message (ISC-189)."""
        def op(c: IMAPClient) -> dict[str, list[int]]:
            info = c.select_folder(folder, readonly=True)
            exists = int(info.get(b"EXISTS", 0) or 0)
            index: dict[str, list[int]] = {}
            c.use_uid = False  # sequence ranges: every reply is one line per message
            try:
                for start in range(1, exists + 1, INDEX_CHUNK):
                    end = min(start + INDEX_CHUNK - 1, exists)
                    data = c.fetch(f"{start}:{end}",
                                   ["UID", "BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)]"])
                    for item in data.values():
                        uid = item.get(b"UID")
                        for mid in header_message_ids(item):
                            if isinstance(uid, int):
                                index.setdefault(mid, []).append(uid)
                    if on_progress:
                        on_progress(end, exists)
            finally:
                c.use_uid = True
            for uids in index.values():
                uids.sort()
            return index
        return self._call(op, f"FETCH headers {folder!r}")

    def append(self, folder: str, mime: bytes, flags: list[str],
               internal_date: datetime, message_id: str | None = None) -> int | None:
        """APPEND the message byte-for-byte; return the UID from APPENDUID if present.

        If the connection drops during APPEND the server may already have committed the
        message. After reconnecting, a message with a Message-ID is looked up first and,
        if present, its UID is returned instead of appending a duplicate.
        """
        def do_append(c: IMAPClient) -> object:
            return c.append(folder, mime, flags=tuple(flags), msg_time=internal_date)

        if self._client is None:
            self.connect()
        try:
            resp = do_append(self._client)  # type: ignore[arg-type]
        except CONNECTION_ERRORS as exc:
            log.warning("IMAP connection lost during APPEND %r (%s); reconnecting once",
                        folder, exc.__class__.__name__)
            self._client = None
            try:
                self.connect()
                if message_id:
                    client = self._client
                    for uid in self._search(client, folder, message_id)[-5:]:  # type: ignore[arg-type]
                        data = client.fetch([uid], ["BODY.PEEK[]"]).get(uid) or {}  # type: ignore[union-attr]
                        if _same_bytes(bytes(data.get(b"BODY[]", b"")), mime):
                            return uid  # committed before the connection dropped
                resp = do_append(self._client)  # type: ignore[arg-type]
            except CONNECTION_ERRORS as exc2:
                self._client = None
                raise ImapConnectionError(
                    f"IMAP APPEND failed after reconnect: {exc2.__class__.__name__}"
                ) from exc2
            except IMAPClientError as exc2:
                raise ImapError(f"IMAP APPEND refused: {_short(exc2)}") from exc2
        except IMAPClientError as exc:
            raise ImapError(f"IMAP APPEND refused: {_short(exc)}") from exc
        raw = resp if isinstance(resp, bytes) else str(resp).encode()
        m = _APPENDUID.search(raw)
        return int(m.group(2)) if m else None

    def supports_multiappend(self) -> bool:
        if self._client is None:
            self.connect()
        try:
            return self._client.has_capability("MULTIAPPEND")  # type: ignore[union-attr]
        except (IMAPClientError, *CONNECTION_ERRORS):
            return False

    def _append_singly(self, folder: str, items: list[AppendItem]) -> list[AppendOutcome]:
        """One APPEND per item; a refused item becomes its ImapError. An out-of-quota
        NO ends the list there (the rest is not attempted): the caller stops."""
        out: list[AppendOutcome] = []
        for it in items:
            try:
                out.append(self.append(folder, it.mime, list(it.flags), it.internal_date,
                                       message_id=it.message_id))
            except ImapError as exc:
                out.append(exc)
                if is_quota_error(exc):
                    break
        return out

    def append_many(self, folder: str, items: list[AppendItem]) -> list[AppendOutcome]:
        """APPEND several messages in one MULTIAPPEND command (RFC 3502; LITERAL+ makes
        it one round trip). Returns, per item and in order, its UID from APPENDUID,
        None when unknown, or the ImapError that refused that one message. Dovecot
        spends ~80 ms per message this way against ~400 ms for one APPEND per command
        (ISC-193). The command is atomic: either every message is stored or none. On a
        server NO the batch falls back to single appends so one refused message does
        not hold back the others and the messages that still fit under a quota are
        stored; an out-of-quota NO is the last outcome returned (shorter list) and the
        caller stops the mailbox. A dropped connection is handled like ``append``:
        after reconnecting, a batch whose first identifiable message is already present
        is treated as committed."""
        if not items:
            return []
        if not self.supports_multiappend():
            return self._append_singly(folder, items)
        batch = [{"msg": it.mime, "flags": tuple(it.flags), "date": it.internal_date}
                 for it in items]

        def do_multi(c: IMAPClient) -> object:
            return c.multiappend(folder, batch)

        if self._client is None:
            self.connect()
        try:
            resp = do_multi(self._client)  # type: ignore[arg-type]
        except CONNECTION_ERRORS as exc:
            log.warning("IMAP connection lost during MULTIAPPEND %r (%s); reconnecting once",
                        folder, exc.__class__.__name__)
            self._client = None
            try:
                self.connect()
                probe = next((it for it in items if it.message_id), None)
                if probe is not None and self._search(self._client, folder,  # type: ignore[arg-type]
                                                      probe.message_id):  # type: ignore[arg-type]
                    # atomic batch, and its first message is there: all committed
                    return [self._search(self._client, folder, it.message_id)[-1]  # type: ignore[arg-type]
                            if it.message_id and self._search(self._client, folder, it.message_id)  # type: ignore[arg-type]
                            else None for it in items]
                resp = do_multi(self._client)  # type: ignore[arg-type]
            except CONNECTION_ERRORS as exc2:
                self._client = None
                raise ImapConnectionError(
                    f"IMAP MULTIAPPEND failed after reconnect: {exc2.__class__.__name__}"
                ) from exc2
            except IMAPClientError:
                return self._append_singly(folder, items)
        except IMAPClientError as exc:
            log.info("MULTIAPPEND of %d messages into %r refused (%s); appending one by one",
                     len(items), folder, _short(exc))
            return self._append_singly(folder, items)
        return list(self._uids_from_multiappend(resp, len(items)))

    @staticmethod
    def _uids_from_multiappend(resp: object, n: int) -> list[int | None]:
        raw = b" ".join(x for x in (resp[1] if isinstance(resp, tuple) else [resp])
                        if isinstance(x, bytes)) if isinstance(resp, (tuple, list)) \
            else (resp if isinstance(resp, bytes) else str(resp).encode())
        m = _APPENDUID_SET.search(raw)
        if not m:
            return [None] * n
        uids = expand_uid_set(m.group(2).decode("ascii"))
        return [int(u) for u in uids] if len(uids) == n else [None] * n

    def fetch_message(self, folder: str, uid: int) -> bytes:
        """FETCH BODY.PEEK[] of one message (dedupe comparison and verify sampling; does
        not set \\Seen). A missing message or a NIL body is an ``ImapError``, never an
        empty byte string: the callers compare content, and "" would read as "different"
        and produce a duplicate copy."""
        def op(c: IMAPClient) -> bytes:
            c.select_folder(folder, readonly=True)
            data = c.fetch([uid], ["BODY.PEEK[]"])
            item = data.get(uid)
            body = item.get(b"BODY[]") if item else None
            if not isinstance(body, bytes):
                raise ImapError(f"FETCH {folder!r} uid {uid}: no message body returned "
                                f"({'absent' if item is None else 'NIL'})")
            return body
        return self._call(op, f"FETCH {folder!r}")


def header_message_ids(item: dict) -> list[str]:
    """Every ``<msg-id>`` of a fetched ``HEADER.FIELDS (MESSAGE-ID)`` item, lower-cased
    (IMAP SEARCH compares headers case-insensitively, so does this index). A message
    carrying two Message-ID headers is indexed under both."""
    for key, value in item.items():
        if isinstance(key, bytes) and key.upper().startswith(b"BODY[HEADER.FIELDS") \
                and isinstance(value, bytes):
            unfolded = re.sub(rb"\r?\n[ \t]+", b" ", value)
            return [m.decode("latin-1").strip().lower() for m in _HEADER_MID.findall(unfolded)]
    return []


def header_message_id(item: dict) -> str | None:
    """First Message-ID of a fetched header item, or None."""
    ids = header_message_ids(item)
    return ids[0] if ids else None


def _short(exc: BaseException) -> str:
    return str(exc)[:200]


def _same_bytes(a: bytes, b: bytes) -> bool:
    norm = (lambda d: d.replace(b"\r\n", b"\n").replace(b"\r", b"\n"))  # noqa: E731
    return norm(a) == norm(b)
