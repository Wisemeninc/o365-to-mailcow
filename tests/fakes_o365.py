"""Shared fakes for the migration tests: Graph, IMAP destination, DAV, config, converters.

Importing this module also makes sure ``o365_to_mailcow.calendar_conv`` and
``o365_to_mailcow.contacts_conv`` are importable: when the real modules are missing (they
are owned by a parallel work stream) minimal stand-ins with the agreed signatures are
registered. Tests monkeypatch the converter functions anyway.
"""

from __future__ import annotations

import importlib
import sys
import types
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import o365_to_mailcow
from o365_to_mailcow.config import Config, MailboxMapping
from o365_to_mailcow.graph import GraphError, GraphTooLarge
from o365_to_mailcow.imap_dest import ImapError


def _ensure_module(name: str, build: Any) -> None:
    full = f"o365_to_mailcow.{name}"
    try:
        importlib.import_module(full)
    except ModuleNotFoundError as exc:
        if exc.name != full:
            raise
        mod = types.ModuleType(full)
        build(mod)
        sys.modules[full] = mod
        setattr(o365_to_mailcow, name, mod)


def _calendar_stub(mod: types.ModuleType) -> None:
    @dataclass(frozen=True)
    class ConvertedEvent:
        uid: str
        ics: bytes
        last_modified: str | None
        warnings: tuple[str, ...]

    def convert_event(master, instances, *, window, attendees="keep"):
        uid = master.get("iCalUId") or master["id"]
        return ConvertedEvent(uid, b"BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n",
                              master.get("lastModifiedDateTime"), ())

    mod.ConvertedEvent = ConvertedEvent
    mod.convert_event = convert_event


def _contacts_stub(mod: types.ModuleType) -> None:
    import uuid

    @dataclass(frozen=True)
    class ConvertedContact:
        uid: str
        vcf: bytes
        last_modified: str | None

    def contact_uid(graph_id: str) -> str:
        return str(uuid.uuid5(uuid.NAMESPACE_URL, f"o365-contact:{graph_id}"))

    def convert_contact(contact, photo=None):
        return ConvertedContact(contact_uid(contact["id"]), b"BEGIN:VCARD\r\nEND:VCARD\r\n",
                                contact.get("lastModifiedDateTime"))

    mod.ConvertedContact = ConvertedContact
    mod.contact_uid = contact_uid
    mod.convert_contact = convert_contact


_ensure_module("calendar_conv", _calendar_stub)
_ensure_module("contacts_conv", _contacts_stub)


# -- config ----------------------------------------------------------------------------

def make_config(tmp_path: Path, **overrides: Any) -> Config:
    values: dict[str, Any] = dict(
        tenant_id="tenant", client_id="client", auth_mode="app",
        client_secret="client-secret-value-123", mailcow_host="mail.example.net",
        mailcow_api_key="api-key-value-456", state_dir=tmp_path / "state",
        mailboxes=(MailboxMapping("alice@contoso.com", "alice@example.net"),),
    )
    values.update(overrides)
    return Config(**values)


MAPPING = MailboxMapping("alice@contoso.com", "alice@example.net")


# -- Graph -----------------------------------------------------------------------------

class FakeGraph:
    """Implements the GraphClient surface over a dict of routes.

    Route values: ``list`` (collection for iter_pages), ``dict`` (get), ``bytes``
    (get_bytes), ``tuple(items, link)`` (get_delta), an ``Exception`` instance (raised),
    or a callable ``(params) -> value``. Every call is recorded with its headers.
    """

    def __init__(self, routes: dict[str, Any] | None = None) -> None:
        self.routes: dict[str, Any] = dict(routes or {})
        self.calls: list[tuple[str, str, dict | None, dict | None]] = []
        self.retries: list[tuple[str, int | None]] = []  # get(): path and retry override

    def _resolve(self, method: str, path: str, params: dict | None,
                 headers: dict | None) -> Any:
        self.calls.append((method, path, params, headers))
        if path not in self.routes:
            raise GraphError(404, "ErrorItemNotFound: no route", path)
        value = self.routes[path]
        if callable(value) and not isinstance(value, type):
            value = value(params)
        if isinstance(value, Exception):
            raise value
        return value

    def get(self, path: str, params: dict | None = None, headers: dict | None = None,
            retries: int | None = None) -> dict:
        self.retries.append((path, retries))
        return self._resolve("get", path, params, headers)

    def iter_pages(self, path: str, params: dict | None = None, headers: dict | None = None):
        yield from self._resolve("iter_pages", path, params, headers)

    def get_delta(self, path: str, params: dict | None = None, headers: dict | None = None):
        return self._resolve("get_delta", path, params, headers)

    def get_bytes(self, path: str, headers: dict | None = None,
                  max_bytes: int | None = None) -> bytes:
        data = self._resolve("get_bytes", path, None, headers)
        if max_bytes is not None and len(data) > max_bytes:
            raise GraphTooLarge(path, max_bytes)
        return data

    def paths(self, method: str | None = None) -> list[str]:
        return [p for m, p, _, _ in self.calls if method is None or m == method]


# -- IMAP ------------------------------------------------------------------------------

@dataclass
class Stored:
    uid: int
    mime: bytes
    flags: list[str]
    date: datetime
    message_id: str | None


@dataclass
class ImapWorld:
    """Server-side state shared by every FakeImap connection."""

    folders: dict[str, list[Stored]] = field(default_factory=dict)
    uidvalidity: dict[str, int] = field(default_factory=dict)
    connections: int = 0
    quota_after: int | None = None  # APPEND NO [OVERQUOTA] after this many appends
    appends: int = 0
    append_threads: set[str] = field(default_factory=set)
    indexed: list[str] = field(default_factory=list)  # folders indexed by Message-ID
    batches: list[int] = field(default_factory=list)  # sizes of append_many calls


class FakeImap:
    """Implements the ImapDestination surface against an ImapWorld."""

    def __init__(self, world: ImapWorld, delimiter: str = "/") -> None:
        self.world = world
        self._delimiter = delimiter
        self.connected = False

    def connect(self) -> None:
        self.world.connections += 1
        self.connected = True

    def close(self) -> None:
        self.connected = False

    @property
    def delimiter(self) -> str:
        return self._delimiter

    def ensure_folder(self, name: str) -> None:
        self.world.folders.setdefault(name, [])
        self.world.uidvalidity.setdefault(name, 1000 + len(self.world.uidvalidity))

    def folder_status(self, name: str) -> tuple[int, int]:
        if name not in self.world.folders:
            raise ImapError(f"STATUS {name!r} refused: NO mailbox does not exist")
        return len(self.world.folders[name]), self.world.uidvalidity[name]

    def search_message_id(self, folder: str, message_id: str) -> list[int]:
        return [s.uid for s in self.world.folders.get(folder, []) if s.message_id == message_id]

    def has_message_id(self, folder: str, message_id: str) -> bool:
        return bool(self.search_message_id(folder, message_id))

    def message_id_index(self, folder: str, on_progress=None) -> dict[str, list[int]]:
        self.world.indexed.append(folder)
        index: dict[str, list[int]] = {}
        for s in self.world.folders.get(folder, []):
            if s.message_id:
                index.setdefault(s.message_id.lower(), []).append(s.uid)
        if on_progress:
            on_progress(len(self.world.folders.get(folder, [])),
                        len(self.world.folders.get(folder, [])))
        return index

    def append_many(self, folder: str, items: list) -> list:
        """Like the real one: one atomic batch; a refused item is its ImapError
        (quota raised), recorded by the world as a batch of this size."""
        from o365_to_mailcow.imap_dest import ImapError, is_quota_error

        self.world.batches.append(len(items))
        out: list = []
        before = len(self.world.folders.get(folder, []))
        try:
            for it in items:
                try:
                    out.append(self.append(folder, it.mime, list(it.flags), it.internal_date,
                                           message_id=it.message_id))
                except ImapError as exc:
                    out.append(exc)
                    if is_quota_error(exc):
                        break
        except BaseException:  # MULTIAPPEND is atomic: a failed batch stores nothing
            del self.world.folders[folder][before:]
            raise
        return out

    def append(self, folder: str, mime: bytes, flags: list[str],
               internal_date: datetime, message_id: str | None = None) -> int | None:
        import threading

        self.world.append_threads.add(threading.current_thread().name)
        if self.world.quota_after is not None and self.world.appends >= self.world.quota_after:
            raise ImapError("IMAP APPEND refused: [OVERQUOTA] Quota exceeded")
        self.world.appends += 1
        box = self.world.folders[folder]
        uid = len(box) + 1
        mid = None
        for line in mime.splitlines():
            if line.lower().startswith(b"message-id:"):
                mid = line.split(b":", 1)[1].strip().decode()
        box.append(Stored(uid, mime, list(flags), internal_date, mid))
        return uid

    def fetch_message(self, folder: str, uid: int) -> bytes:
        return next(s.mime for s in self.world.folders[folder] if s.uid == uid)


# -- DAV -------------------------------------------------------------------------------

class FakeDav:
    """Implements the SogoDav surface in memory, recording every write."""

    def __init__(self, calendar_home: bool = True, addressbook_home: bool = True,
                 refuse_mkcol: bool = False) -> None:
        self.calendar_home = calendar_home
        self.addressbook_home = addressbook_home
        self.refuse_mkcol = refuse_mkcol
        self.calendars: dict[str, dict[str, bytes]] = {"personal": {}}
        self.books: dict[str, dict[str, bytes]] = {"personal": {}, "collected": {}}
        self.writes: list[tuple[str, str]] = []
        self.put_errors: dict[str, Exception] = {}

    def calendar_home_exists(self) -> bool:
        return self.calendar_home

    def addressbook_home_exists(self) -> bool:
        return self.addressbook_home

    def list_calendars(self) -> list[tuple[str, str]]:
        return [(s, s) for s in self.calendars]

    def list_addressbooks(self) -> list[tuple[str, str]]:
        return [(s, s) for s in self.books]

    def ensure_calendar(self, name: str, slug: str | None = None) -> str:
        assert slug
        if slug not in self.calendars:
            self.writes.append(("MKCALENDAR", slug))
            self.calendars[slug] = {}
        return slug

    def ensure_addressbook(self, name: str, slug: str | None = None) -> str:
        from o365_to_mailcow.dav import DavError

        assert slug
        if self.refuse_mkcol:
            raise DavError(405, "Method Not Allowed")
        if slug not in self.books:
            self.writes.append(("MKCOL", slug))
            self.books[slug] = {}
        return slug

    def put_event(self, slug: str, uid: str, ics: bytes) -> None:
        if uid in self.put_errors:
            raise self.put_errors[uid]
        self.writes.append(("PUT", f"Calendar/{slug}/{uid}.ics"))
        self.calendars[slug][uid] = ics

    def put_contact(self, slug: str, uid: str, vcf: bytes) -> None:
        if uid in self.put_errors:
            raise self.put_errors[uid]
        self.writes.append(("PUT", f"Contacts/{slug}/{uid}.vcf"))
        self.books[slug][uid] = vcf

    def count_resources(self, kind: str, slug: str) -> int:
        from o365_to_mailcow.dav import DavError

        store = self.calendars if kind.lower() == "calendar" else self.books
        if slug not in store:
            raise DavError(404, "Not Found")
        return len(store[slug])
