"""Mail migration: Graph mail folders and MIME -> IMAP APPEND on mailcow.

Flow per mailbox: build the folder plan (well-known mapping, skips, sanitised hierarchy),
then per folder: ensure the destination folder, check UIDVALIDITY against state, list
the messages (full listing, or the stored delta link on later runs), filter what is
already done, download ``$value`` on four threads and APPEND through the one IMAP
connection owned by the calling thread. State is committed after every APPEND.

Nothing is ever deleted, moved or flagged at either end; ``@removed`` delta entries are
only counted (ISC-124).
"""

from __future__ import annotations

import hashlib
import logging
import re
import secrets
import time
import unicodedata
from collections import deque
from collections.abc import Callable, Iterator
from concurrent.futures import Future, ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import UTC, datetime
from email import message_from_bytes, policy
from email.parser import BytesHeaderParser
from urllib.parse import quote

from dateutil import parser as dtparser

from .config import Config, MailboxMapping
from .graph import GraphClient, GraphError, GraphTooLarge, delta_expired
from .imap_dest import ImapConnectionError, ImapDestination, ImapError, is_quota_error
from .report import NullProgress, Progress
from .state import STATUS_DONE, STATUS_FAILED, STATUS_SKIPPED, State

log = logging.getLogger(__name__)

PREFER_IMMUTABLE = {"Prefer": 'IdType="ImmutableId"'}
# Delta pages default to 10 items; ask for bigger pages. RFC 7240 allows a list.
PREFER_DELTA = {"Prefer": 'IdType="ImmutableId", odata.maxpagesize=200'}
# Graph v1.0 messages have no "size" property; PR_MESSAGE_SIZE is the documented way.
SIZE_PROPERTY = "Integer 0x0E08"
MESSAGE_SELECT = (
    "id,internetMessageId,isRead,isDraft,flag,categories,"
    "receivedDateTime,lastModifiedDateTime"
)
MESSAGE_EXPAND = f"singleValueExtendedProperties($filter=id eq '{SIZE_PROPERTY}')"
# v1.0 mail folders have no sizeInBytes either; PR_MESSAGE_SIZE_EXTENDED gives it.
FOLDER_SIZE_PROPERTY = "Long 0x0E08"
FOLDER_EXPAND = f"singleValueExtendedProperties($filter=id eq '{FOLDER_SIZE_PROPERTY}')"
PAGE_SIZE = 100
DOWNLOAD_WORKERS = 4
PREFETCH_WINDOW = 4  # bounded by count; each item is at most max_message_bytes
MAX_KEYWORDS = 20
MAX_KEYWORD_LEN = 50  # Dovecot's default mail_max_keyword_length

WELL_KNOWN_MAP = {
    "inbox": "INBOX",
    "sentitems": "Sent",
    "drafts": "Drafts",
    "deleteditems": "Trash",
    "junkemail": "Junk",
    "archive": "Archive",
}
WELL_KNOWN_SKIP = (
    "conversationhistory",
    "outbox",
    "syncissues",
    "recoverableitemsdeletions",
    "serverfailures",
    "localfailures",
)
_KEYWORD_BAD = re.compile(r"[^A-Za-z0-9_\-.+:@#&!]")
# RFC 5322 msg-id restricted to atext (no IMAP atom-specials, no domain literals); anything
# else is treated as "no Message-ID" so tenant data never reaches an IMAP SEARCH unquoted.
_ATEXT = r"[A-Za-z0-9!#$&'+\-/=?^_`|~.]"
_MESSAGE_ID = re.compile(rf"^<{_ATEXT}{{1,500}}@{_ATEXT}{{1,494}}>$")
_WS = re.compile(r"\s+")


# -- dataclasses -----------------------------------------------------------------------

@dataclass
class FolderPlan:
    folder_id: str
    source_path: str
    dest_name: str
    total: int
    size_bytes: int | None
    well_known: str | None = None
    skip: bool = False
    skip_reason: str | None = None


@dataclass
class MailPlan:
    mailbox: str
    delimiter: str
    folders: list[FolderPlan]
    total_messages: int = 0
    total_bytes: int = 0
    bytes_known: bool = True
    skipped_folders: int = 0


@dataclass
class FolderResult:
    folder_id: str
    dest_name: str
    listed: int = 0
    appended: int = 0
    already_done: int = 0
    dedup_hits: int = 0
    failed: int = 0
    skipped_too_large: int = 0
    removed_in_source: int = 0
    vanished_in_source: int = 0
    reappended_after_uidvalidity: int = 0
    would_append: int = 0
    uidvalidity_changed: bool = False
    delta_reset: bool = False
    error: str | None = None


@dataclass
class MailResult:
    mailbox: str
    dry_run: bool
    folders: list[FolderResult] = field(default_factory=list)
    skipped_folders: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    stopped: bool = False
    duration_s: float = 0.0

    def total(self, attr: str) -> int:
        return sum(getattr(f, attr) for f in self.folders)

    @property
    def failed(self) -> int:
        return self.total("failed") + sum(1 for f in self.folders if f.error)


@dataclass
class FolderVerify:
    dest_name: str
    graph_total: int
    done: int
    failed: int
    skipped: int
    imap_count: int
    expected: int
    mismatch: bool
    note: str | None = None


@dataclass
class MailVerify:
    mailbox: str
    folders: list[FolderVerify] = field(default_factory=list)
    skipped_folders: list[dict] = field(default_factory=list)  # {path, reason, total}
    sample_requested: int = 0
    sample_checked: int = 0
    sample_mismatches: list[str] = field(default_factory=list)
    sample_regenerated: int = 0  # same headers and size class, different MIME rendering
    sample_unverifiable: int = 0
    errors: list[str] = field(default_factory=list)


# -- pure helpers ----------------------------------------------------------------------

def sanitize_folder_name(name: str, delimiter: str) -> str:
    """Remove control characters, strip whitespace, replace the hierarchy delimiter."""
    cleaned = "".join(ch for ch in name if unicodedata.category(ch) != "Cc").strip()
    if delimiter:
        cleaned = cleaned.replace(delimiter, "_")
    return cleaned or "Unnamed"


def category_keyword(category: str) -> str:
    """Graph category -> IMAP keyword (atom).

    Spaces and disallowed characters become ``_``; a leading ``$`` or ``\\`` is replaced
    so a category can never masquerade as a system flag or a client-defined ``$Keyword``;
    the result is capped at Dovecot's default keyword length.
    """
    kw = _KEYWORD_BAD.sub("_", category.strip())[:MAX_KEYWORD_LEN]
    if kw[:1] in ("$", "\\"):
        kw = "_" + kw[1:]
    return kw


def valid_message_id(value: object) -> str | None:
    """Return the Message-ID if it is a plain RFC 5322 msg-id, else None (ISC-98)."""
    if isinstance(value, str) and _MESSAGE_ID.match(value):
        return value
    return None


def imap_flags(msg: dict) -> list[str]:
    """ISC-42..45: \\Seen, \\Flagged, \\Draft and category keywords."""
    flags: list[str] = []
    if msg.get("isRead"):
        flags.append("\\Seen")
    if (msg.get("flag") or {}).get("flagStatus") == "flagged":
        flags.append("\\Flagged")
    if msg.get("isDraft"):
        flags.append("\\Draft")
    keywords = 0
    for cat in msg.get("categories") or []:
        kw = category_keyword(str(cat))
        if kw and kw not in flags:
            flags.append(kw)
            keywords += 1
            if keywords >= MAX_KEYWORDS:
                break
    return flags


def internal_date(msg: dict, now: Callable[[], datetime] | None = None) -> datetime:
    """ISC-46/103: receivedDateTime, else lastModifiedDateTime, else now (UTC)."""
    for key in ("receivedDateTime", "lastModifiedDateTime"):
        raw = msg.get(key)
        if raw:
            try:
                value = dtparser.isoparse(raw)
            except (ValueError, OverflowError):
                continue
            return value if value.tzinfo else value.replace(tzinfo=UTC)
    return (now or (lambda: datetime.now(UTC)))()


def _prop_key(prop_id: str) -> tuple[str, int] | None:
    """``"Integer 0x0E08"`` -> ``("integer", 0xE08)``; Graph echoes ids as ``0xe08``."""
    kind, _, tag = prop_id.strip().partition(" ")
    try:
        return kind.lower(), int(tag, 16)
    except ValueError:
        return None


def _extended_int(obj: dict, prop_id: str) -> int | None:
    wanted = _prop_key(prop_id)
    for prop in obj.get("singleValueExtendedProperties") or []:
        if _prop_key(str(prop.get("id", ""))) == wanted:
            try:
                return int(prop.get("value"))
            except (TypeError, ValueError):
                return None
    return None


def message_size(msg: dict) -> int | None:
    """Size in bytes from ``size`` or the PR_MESSAGE_SIZE extended property, if present."""
    if isinstance(msg.get("size"), int):
        return msg["size"]
    return _extended_int(msg, SIZE_PROPERTY)


def folder_size(folder: dict) -> int | None:
    """Folder size from ``sizeInBytes`` (beta) or PR_MESSAGE_SIZE_EXTENDED, if present."""
    if isinstance(folder.get("sizeInBytes"), int):
        return folder["sizeInBytes"]
    return _extended_int(folder, FOLDER_SIZE_PROPERTY)


def _normalised_sha256(data: bytes) -> str:
    # IMAP APPEND (imaplib) rewrites bare CR/LF as CRLF; compare line-ending-agnostic
    return hashlib.sha256(data.replace(b"\r\n", b"\n").replace(b"\r", b"\n")).hexdigest()


def _leaf_digests(raw: bytes) -> list[str] | None:
    """SHA-256 of every decoded leaf part (text and attachments), boundaries excluded.
    Stable across Exchange's MIME re-rendering, which changes boundaries, header order
    and encodings but not the payload bytes. None if the message cannot be parsed."""
    try:
        msg = message_from_bytes(raw, policy=policy.default)
        out: list[str] = []
        for part in msg.walk():
            if part.is_multipart():
                continue
            payload = part.get_payload(decode=True) or b""
            if part.get_content_maintype() == "text":
                payload = payload.replace(b"\r\n", b"\n").replace(b"\r", b"\n").strip()
            out.append(hashlib.sha256(payload).hexdigest())
        return sorted(out)
    except Exception:  # noqa: BLE001 - malformed MIME: no verdict possible
        return None


def _header(msg, name: str) -> str:
    value = msg.get(name)
    return _WS.sub(" ", str(value)).strip().lower() if value else ""


def same_message(source: bytes, dest: bytes) -> str:
    """Decide whether two MIME renderings are the same message.

    Returns ``"identical"`` (byte-equal after line-ending normalisation), ``"regenerated"``
    (Exchange re-renders MIME for items it stores as MAPI: same Message-ID, Date, From and
    Subject, and every decoded leaf part, attachments included, has the same digest) or
    ``"different"``. A copy that lost or changed any part is therefore never "regenerated".
    """
    if _normalised_sha256(source) == _normalised_sha256(dest):
        return "identical"
    try:
        # policy.default parses header values lazily inside get(); a malformed From or
        # Message-ID raises there, so the whole comparison sits inside the try.
        parser = BytesHeaderParser(policy=policy.default)
        a, b = parser.parsebytes(source), parser.parsebytes(dest)
        for name in ("message-id", "date", "from", "subject"):
            if _header(a, name) != _header(b, name):
                return "different"
    except Exception:  # noqa: BLE001 - unparsable headers are simply "different"
        return "different"
    src_parts, dst_parts = _leaf_digests(source), _leaf_digests(dest)
    if src_parts is None or dst_parts is None or src_parts != dst_parts:
        return "different"
    return "regenerated"


# -- migrator --------------------------------------------------------------------------

class _DeltaLinkError(Exception):
    """A GraphError raised by the delta call itself (not by a per-message fetch)."""

    def __init__(self, cause: GraphError) -> None:
        super().__init__(str(cause))
        self.cause = cause


class MailMigrator:
    """Plan, migrate and verify the mail of one mailbox."""

    def __init__(self, cfg: Config, graph: GraphClient, state: State,
                 dest_factory: Callable[[], ImapDestination] | None,
                 mapping: MailboxMapping, dry_run: bool,
                 progress: Progress | NullProgress | None = None) -> None:
        self._cfg = cfg
        self._graph = graph
        self._state = state
        self._dest_factory = dest_factory
        self._mapping = mapping
        self._src = mapping.source
        self._dry_run = dry_run
        self._progress = progress or NullProgress()
        self._key = f"{mapping.source} mail"

    # -- folder plan -------------------------------------------------------------------

    def _user_path(self, rest: str) -> str:
        return f"/users/{quote(self._src, safe='@')}/{rest}"

    def _well_known_ids(self) -> dict[str, str]:
        """Graph folder id -> well-known name, for every well-known folder that exists."""
        out: dict[str, str] = {}
        for name in (*WELL_KNOWN_MAP, *WELL_KNOWN_SKIP):
            try:
                data = self._graph.get(self._user_path(f"mailFolders/{name}"),
                                       params={"$select": "id"}, headers=PREFER_IMMUTABLE)
            except GraphError as exc:
                if exc.status in (400, 404):
                    continue
                raise
            if data.get("id"):
                out[data["id"]] = name
        return out

    def _walk(self) -> Iterator[tuple[dict, list[dict]]]:
        """Yield ``(folder, ancestors)`` depth-first, parents before children."""
        params = {"$top": PAGE_SIZE, "$expand": FOLDER_EXPAND, "includeHiddenFolders": "true"}
        stack: list[tuple[dict, list[dict]]] = [
            (f, []) for f in self._graph.iter_pages(
                self._user_path("mailFolders"), params=params, headers=PREFER_IMMUTABLE)
        ]
        stack.reverse()
        while stack:
            folder, ancestors = stack.pop()
            yield folder, ancestors
            if folder.get("childFolderCount", 1):
                children = list(self._graph.iter_pages(
                    self._user_path(f"mailFolders/{folder['id']}/childFolders"),
                    params=params, headers=PREFER_IMMUTABLE))
                for child in reversed(children):
                    stack.append((child, [*ancestors, folder]))

    def _build_plan(self, delimiter: str) -> MailPlan:
        wk_ids = self._well_known_ids()
        user_skips = {s.strip().lower() for s in self._cfg.source_folder_skip}
        used: set[str] = {v.lower() for v in WELL_KNOWN_MAP.values()}
        dest_by_id: dict[str, str] = {}
        skipped_ids: set[str] = set()
        folders: list[FolderPlan] = []
        for folder, ancestors in self._walk():
            fid = folder["id"]
            display = str(folder.get("displayName") or "")
            path = "/".join([*(str(a.get("displayName", "")) for a in ancestors), display])
            wk = wk_ids.get(fid)
            skip_reason: str | None = None
            if ancestors and ancestors[-1]["id"] in skipped_ids:
                skip_reason = "parent skipped"
            elif wk in WELL_KNOWN_SKIP:
                skip_reason = f"well-known folder {wk}"
            elif folder.get("isHidden") and not folder.get("totalItemCount"):
                skip_reason = "hidden empty system folder"
            elif path.lower() in user_skips or display.lower() in user_skips:
                skip_reason = "source_folder_skip"
            if wk in WELL_KNOWN_MAP:
                dest = WELL_KNOWN_MAP[wk]
            else:
                base = sanitize_folder_name(display, delimiter)
                prefix = (dest_by_id.get(ancestors[-1]["id"], "") + delimiter) if ancestors else ""
                dest = prefix + base
                n = 1
                while dest.lower() in used:
                    n += 1
                    dest = f"{prefix}{base} ({n})"
                used.add(dest.lower())
            dest_by_id[fid] = dest
            if skip_reason:
                skipped_ids.add(fid)
            size = folder_size(folder)
            folders.append(FolderPlan(
                folder_id=fid, source_path=path, dest_name=dest,
                total=int(folder.get("totalItemCount") or 0),
                size_bytes=size,
                well_known=wk, skip=skip_reason is not None, skip_reason=skip_reason,
            ))
        active = [f for f in folders if not f.skip]
        return MailPlan(
            mailbox=self._src, delimiter=delimiter, folders=folders,
            total_messages=sum(f.total for f in active),
            total_bytes=sum(f.size_bytes or 0 for f in active),
            bytes_known=all(f.size_bytes is not None for f in active),
            skipped_folders=len(folders) - len(active),
        )

    def plan(self, delimiter: str = "/") -> MailPlan:
        """Read-only plan. Uses mailcow's default ``/`` delimiter; no IMAP needed."""
        return self._build_plan(delimiter)

    # -- listing -----------------------------------------------------------------------

    def _list_messages(self, folder_id: str) -> Iterator[dict]:
        return self._graph.iter_pages(
            self._user_path(f"mailFolders/{folder_id}/messages"),
            params={"$select": MESSAGE_SELECT, "$top": PAGE_SIZE, "$expand": MESSAGE_EXPAND},
            headers=PREFER_IMMUTABLE,
        )

    def _initial_delta_link(self, folder_id: str) -> str | None:
        """Take a delta snapshot (ids only) *before* the full listing so nothing is missed."""
        try:
            _, link = self._graph.get_delta(
                self._user_path(f"mailFolders/{folder_id}/messages/delta"),
                params={"$select": "id"}, headers=PREFER_DELTA)
        except GraphError as exc:
            log.warning("%s: no delta link for folder (%s); next run lists fully",
                        self._src, exc)
            return None
        return link

    def _delta_messages(self, link: str, fr: FolderResult,
                        done: set[str]) -> tuple[list[dict], str | None]:
        try:
            items, new_link = self._graph.get_delta(link, headers=PREFER_DELTA)
        except GraphError as exc:
            raise _DeltaLinkError(exc) from exc
        out: list[dict] = []
        for item in items:
            if "@removed" in item:
                fr.removed_in_source += 1  # counted, never applied (ISC-124)
                continue
            if item.get("id") in done:
                fr.listed += 1
                fr.already_done += 1
                continue
            try:
                out.append(self._graph.get(
                    self._user_path(f"messages/{item['id']}"),
                    params={"$select": MESSAGE_SELECT, "$expand": MESSAGE_EXPAND},
                    headers=PREFER_IMMUTABLE))
            except GraphError as exc:
                if exc.status != 404:
                    raise
                fr.vanished_in_source += 1  # listed by delta, gone before we fetched it
        return out, new_link

    def _download(self, graph_id: str) -> bytes:
        return self._graph.get_bytes(self._user_path(f"messages/{graph_id}/$value"),
                                     headers=PREFER_IMMUTABLE,
                                     max_bytes=self._cfg.max_message_bytes)

    # -- migrate -----------------------------------------------------------------------

    def migrate(self) -> MailResult:
        started = time.monotonic()
        result = MailResult(mailbox=self._src, dry_run=self._dry_run)
        try:
            if self._dry_run:
                self._dry_run_count(result)
            else:
                self._migrate(result)
        finally:
            result.duration_s = round(time.monotonic() - started, 3)
            self._progress.finish(self._key)
        return result

    def _dry_run_count(self, result: MailResult) -> None:
        plan = self.plan()
        result.skipped_folders = [f"{f.source_path} ({f.skip_reason})"
                                  for f in plan.folders if f.skip]
        self._progress.start(self._key, plan.total_messages)
        for fp in plan.folders:
            if fp.skip:
                continue
            fr = FolderResult(fp.folder_id, fp.dest_name)
            result.folders.append(fr)
            done = self._state.done_message_ids(self._src, fp.folder_id)
            try:
                for msg in self._list_messages(fp.folder_id):
                    fr.listed += 1
                    size = message_size(msg)
                    if msg["id"] in done:
                        fr.already_done += 1
                    elif size is not None and size > self._cfg.max_message_bytes:
                        fr.skipped_too_large += 1
                    else:
                        fr.would_append += 1
                    self._progress.advance(self._key)
            except GraphError as exc:
                fr.error = str(exc)
                result.errors.append(f"{fp.source_path}: {exc}")

    def _migrate(self, result: MailResult) -> None:
        if self._dest_factory is None:
            raise RuntimeError("migrate needs a destination factory")
        dest = self._dest_factory()
        dest.connect()
        try:
            plan = self._build_plan(dest.delimiter)
            result.skipped_folders = [f"{f.source_path} ({f.skip_reason})"
                                      for f in plan.folders if f.skip]
            self._progress.start(self._key, plan.total_messages)
            with ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS,
                                    thread_name_prefix="dl") as pool:
                for fp in plan.folders:
                    if fp.skip:
                        continue
                    fr = FolderResult(fp.folder_id, fp.dest_name)
                    result.folders.append(fr)
                    self._migrate_folder(dest, fp, fr, result, pool)
                    if result.stopped:
                        break
        finally:
            dest.close()

    def _stop(self, result: MailResult, fp: FolderPlan, exc: Exception) -> None:
        result.stopped = True
        line = (f"{self._src}: stopped in folder {fp.dest_name!r}: {exc}; "
                "remaining messages skipped for this run")
        result.errors.append(line)
        log.error("%s", line)

    def _migrate_folder(self, dest: ImapDestination, fp: FolderPlan, fr: FolderResult,
                        result: MailResult, pool: ThreadPoolExecutor) -> None:
        src, fid, name = self._src, fp.folder_id, fp.dest_name
        try:
            dest.ensure_folder(name)
            count_at_start, uidvalidity = dest.folder_status(name)
        except ImapConnectionError as exc:
            fr.error = str(exc)
            self._stop(result, fp, exc)
            return
        except ImapError as exc:
            fr.error = str(exc)
            result.errors.append(f"{fp.source_path}: {exc}")
            return

        stored_uv = self._state.folder_uidvalidity(src, fid)
        forced = stored_uv is not None and stored_uv != uidvalidity
        if forced:
            fr.uidvalidity_changed = True
            log.warning("%s: UIDVALIDITY of %r changed (%s -> %s); re-checking Message-IDs",
                        src, name, stored_uv, uidvalidity)
        done = self._state.done_message_ids(src, fid)
        stored_link = None if forced else self._state.get_delta(src, fid)
        try:
            messages: list[dict] = []
            new_link: str | None = None
            if stored_link:
                try:
                    messages, new_link = self._delta_messages(stored_link, fr, done)
                except _DeltaLinkError as wrapped:
                    exc = wrapped.cause
                    if not delta_expired(exc):
                        raise exc from None
                    # Graph discards old delta tokens; a stale one must never wedge the
                    # folder, so forget it and list fully in the same run.
                    log.warning("%s: delta link for %r expired (%s); listing fully",
                                src, name, exc)
                    self._state.clear_delta(src, fid)
                    fr.delta_reset = True
                    stored_link = None
            if not stored_link:
                new_link = self._initial_delta_link(fid)
                messages = list(self._list_messages(fid))
        except GraphError as exc:
            fr.error = str(exc)
            result.errors.append(f"{fp.source_path}: listing failed: {exc}")
            return

        try:
            todo = self._select_todo(dest, fp, fr, messages, done, forced,
                                     count_at_start, uidvalidity)
        except (ImapError, ImapConnectionError) as exc:
            fr.error = str(exc)
            self._stop(result, fp, exc)
            return
        todo.sort(key=internal_date)  # oldest first, so destination UIDs follow date order
        self._append_all(dest, fp, fr, result, pool, todo, uidvalidity)

        if not result.stopped:
            self._state.set_folder_meta(src, fid, name, uidvalidity)
        if new_link and not result.stopped and fr.failed == 0:
            self._state.set_delta(src, fid, new_link)

    def _select_todo(self, dest: ImapDestination, fp: FolderPlan, fr: FolderResult,
                     messages: list[dict], done: set[str], forced: bool,
                     count_at_start: int, uidvalidity: int) -> list[dict]:
        src, fid, name = self._src, fp.folder_id, fp.dest_name
        todo: list[dict] = []
        for msg in messages:
            fr.listed += 1
            gid = msg["id"]
            mid = valid_message_id(msg.get("internetMessageId"))
            if gid in done:
                # after a UIDVALIDITY change, confirm the copy still exists
                if not forced:
                    fr.already_done += 1
                    self._progress.advance(self._key)
                    continue
                if not mid:
                    # cannot be searched for; if it was recorded under the old UIDVALIDITY
                    # the folder was recreated since, so copy it again (a duplicate is
                    # visible and fixable, a silent gap is not); a row already carrying
                    # the current UIDVALIDITY was appended after the change and is done
                    if self._state.message_uidvalidity(src, fid, gid) == uidvalidity:
                        fr.already_done += 1
                    else:
                        fr.reappended_after_uidvalidity += 1
                        todo.append(msg)
                    self._progress.advance(self._key)
                    continue
                if dest.has_message_id(name, mid):
                    fr.already_done += 1
                    self._progress.advance(self._key)
                    continue
            size = message_size(msg)
            if size is not None and size > self._cfg.max_message_bytes:
                self._state.mark_message(src, fid, gid, name, mid, STATUS_SKIPPED,
                                         error=f"too large: {size} bytes")
                fr.skipped_too_large += 1
                self._progress.advance(self._key)
                continue
            # ISC-48/98: only messages with a Message-ID, only if the folder had content
            if mid and gid not in done and count_at_start > 0:
                try:
                    verdict = self._dedupe(dest, fp, fr, msg, mid, uidvalidity)
                except ImapConnectionError:
                    raise
                except Exception as exc:  # noqa: BLE001 - a dedupe problem never loses mail
                    log.warning("%s: dedupe check failed for a message in %r (%s); appending",
                                src, name, exc.__class__.__name__)
                    verdict = "append"
                if verdict != "append":
                    continue
            todo.append(msg)
        return todo

    def _dedupe(self, dest: ImapDestination, fp: FolderPlan, fr: FolderResult, msg: dict,
                mid: str, uidvalidity: int) -> str:
        """Message-ID hit at the destination: count it as already migrated only when a
        destination copy really is this message (content compared) and the destination
        holds more copies than source items already recorded for this Message-ID.
        Anything else is appended. Returns "done", "skipped" or "append"."""
        src, fid, name, gid = self._src, fp.folder_id, fp.dest_name, msg["id"]
        uids = dest.search_message_id(name, mid)
        already = self._state.done_count_for_message_id(src, fid, mid)
        if len(uids) <= already:
            return "append"
        try:
            mime = self._download(gid)
        except GraphTooLarge as exc:
            self._state.mark_message(src, fid, gid, name, mid, STATUS_SKIPPED,
                                     error=f"too large: over {exc.limit} bytes")
            fr.skipped_too_large += 1
            self._progress.advance(self._key)
            return "skipped"
        except GraphError:
            return "append"  # the normal path records the failure
        for uid in uids[-5:]:
            if same_message(mime, dest.fetch_message(name, uid)) != "different":
                self._state.mark_message(src, fid, gid, name, mid, STATUS_DONE,
                                         dest_uid=uid, uidvalidity=uidvalidity)
                fr.dedup_hits += 1
                self._progress.advance(self._key)
                return "done"
        return "append"

    def _append_all(self, dest: ImapDestination, fp: FolderPlan, fr: FolderResult,
                    result: MailResult, pool: ThreadPoolExecutor, todo: list[dict],
                    uidvalidity: int) -> None:
        src, fid, name = self._src, fp.folder_id, fp.dest_name
        queue = iter(todo)
        pending: deque[tuple[dict, Future[bytes]]] = deque()

        def refill() -> None:
            while len(pending) < PREFETCH_WINDOW:
                nxt = next(queue, None)
                if nxt is None:
                    return
                pending.append((nxt, pool.submit(self._download, nxt["id"])))

        refill()
        while pending:
            msg, fut = pending.popleft()
            refill()
            gid, mid = msg["id"], valid_message_id(msg.get("internetMessageId"))
            try:
                mime = fut.result()
            except GraphTooLarge as exc:  # ISC-52: abandoned while streaming
                self._state.mark_message(src, fid, gid, name, mid, STATUS_SKIPPED,
                                         error=f"too large: over {exc.limit} bytes")
                fr.skipped_too_large += 1
                self._progress.advance(self._key)
                continue
            except GraphError as exc:  # ISC-49: record and continue
                self._state.mark_message(src, fid, gid, name, mid, STATUS_FAILED,
                                         error=f"graph HTTP {exc.status}: {exc}")
                fr.failed += 1
                self._progress.advance(self._key)
                continue
            try:
                uid = dest.append(name, mime, imap_flags(msg), internal_date(msg),
                                  message_id=mid)
            except ImapError as exc:
                self._state.mark_message(src, fid, gid, name, mid, STATUS_FAILED,
                                         error=str(exc))
                fr.failed += 1
                if not is_quota_error(exc):  # one refused message must not halt a mailbox
                    self._progress.advance(self._key)
                    continue
                for _, other in pending:  # ISC-99: out of space, stop this mailbox
                    other.cancel()
                self._stop(result, fp, exc)
                return
            except ImapConnectionError as exc:  # ISC-100: reconnect already failed once
                self._state.mark_message(src, fid, gid, name, mid, STATUS_FAILED,
                                         error=str(exc))
                fr.failed += 1
                for _, other in pending:
                    other.cancel()
                self._stop(result, fp, exc)
                return
            self._state.mark_message(src, fid, gid, name, mid, STATUS_DONE,
                                     dest_uid=uid, uidvalidity=uidvalidity)
            fr.appended += 1
            self._progress.advance(self._key)

    # -- verify ------------------------------------------------------------------------

    def verify(self, sample: int = 0) -> MailVerify:
        if self._dest_factory is None:
            raise RuntimeError("verify needs a destination factory")
        out = MailVerify(mailbox=self._src, sample_requested=sample)
        dest = self._dest_factory()
        dest.connect()
        try:
            plan = self._build_plan(dest.delimiter)
            counts = self._state.message_counts_by_folder_id(self._src)
            for fp in plan.folders:
                if fp.skip:
                    out.skipped_folders.append(
                        {"path": fp.source_path, "reason": fp.skip_reason, "total": fp.total})
                    continue
                c = counts.get(fp.folder_id, {})
                done, failed = c.get(STATUS_DONE, 0), c.get(STATUS_FAILED, 0)
                skipped = c.get(STATUS_SKIPPED, 0)
                note = None
                try:
                    imap_count, _ = dest.folder_status(fp.dest_name)
                except ImapError as exc:
                    imap_count, note = 0, f"STATUS failed: {exc}"
                expected = fp.total - skipped - failed
                out.folders.append(FolderVerify(
                    dest_name=fp.dest_name, graph_total=fp.total, done=done, failed=failed,
                    skipped=skipped, imap_count=imap_count, expected=expected,
                    mismatch=imap_count != expected, note=note,
                ))
            if sample > 0:
                self._sample(dest, plan, sample, out)
        finally:
            dest.close()
        return out

    def _sample(self, dest: ImapDestination, plan: MailPlan, n: int, out: MailVerify) -> None:
        pool = [(fp.dest_name, gid) for fp in plan.folders if not fp.skip
                for gid in sorted(self._state.done_message_ids(self._src, fp.folder_id))]
        chosen = secrets.SystemRandom().sample(pool, min(n, len(pool)))
        for folder, gid in chosen:
            try:
                meta = self._graph.get(self._user_path(f"messages/{gid}"),
                                       params={"$select": "internetMessageId"},
                                       headers=PREFER_IMMUTABLE)
                mid = meta.get("internetMessageId")
                mid = valid_message_id(mid)
                if not mid:
                    out.sample_unverifiable += 1
                    continue
                source = self._download(gid)
            except GraphError as exc:
                if exc.status == 404:
                    out.sample_unverifiable += 1
                    continue
                out.errors.append(f"sample {gid}: {exc}")
                continue
            out.sample_checked += 1
            verdicts = {same_message(source, dest.fetch_message(folder, uid))
                        for uid in dest.search_message_id(folder, mid)[:5]}
            if "identical" in verdicts:
                continue
            if "regenerated" in verdicts:
                out.sample_regenerated += 1
                continue
            out.sample_mismatches.append(f"{folder}: {mid}")
