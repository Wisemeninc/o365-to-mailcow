"""Destination IMAP connection (Dovecot in mailcow), write-only in the append sense.

Only these IMAP commands are ever issued (ISC-54): LOGIN, LIST, CREATE, SUBSCRIBE,
STATUS, SELECT/EXAMINE, SEARCH, APPEND, FETCH (verify sampling only), LOGOUT. The
class deliberately has no way to STORE flags, EXPUNGE or DELETE (ISC-55).

Folder names are passed to imapclient as ``str``; imapclient encodes them to modified
UTF-7 and decodes LIST results back (ISC-40).
"""

from __future__ import annotations

import logging
import re
import ssl
from collections.abc import Callable
from datetime import datetime
from typing import TypeVar

from imapclient import IMAPClient
from imapclient.exceptions import IMAPClientAbortError, IMAPClientError, LoginError
from imapclient.imapclient import SocketTimeout

log = logging.getLogger(__name__)

T = TypeVar("T")
_APPENDUID = re.compile(rb"\[APPENDUID (\d+) (\d+)\]", re.IGNORECASE)
CONNECTION_ERRORS = (IMAPClientAbortError, OSError)


class ImapError(Exception):
    """The server answered NO or BAD (for example OVERQUOTA). Not retried."""


def is_quota_error(exc: BaseException) -> bool:
    """True for the one NO that should stop a mailbox: the destination is out of space."""
    text = str(exc).upper()
    return "OVERQUOTA" in text or "QUOTA" in text


class ImapConnectionError(Exception):
    """The connection failed and one reconnect-and-retry did not help."""


class ImapDestination:
    """One authenticated IMAP connection to the mailcow host for one mailbox."""

    def __init__(self, host: str, port: int, user: str, password: str,
                 verify: bool = True, client_factory: Callable[..., IMAPClient] = IMAPClient,
                 timeout: SocketTimeout | None = None) -> None:
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
        ctx = ssl.create_default_context()
        if not self._verify:  # test hook only; production callers never pass False
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
                    found = self._search(self._client, folder, message_id)  # type: ignore[arg-type]
                    if found:
                        return found[-1]
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

    def fetch_message(self, folder: str, uid: int) -> bytes:
        """FETCH BODY.PEEK[] of one message (verify sampling only; does not set \\Seen)."""
        def op(c: IMAPClient) -> bytes:
            c.select_folder(folder, readonly=True)
            data = c.fetch([uid], ["BODY.PEEK[]"])
            item = data.get(uid) or {}
            return bytes(item.get(b"BODY[]", b""))
        return self._call(op, f"FETCH {folder!r}")


def _short(exc: BaseException) -> str:
    return str(exc)[:200]
