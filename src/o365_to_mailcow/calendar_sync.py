"""Calendar migration: Graph calendars/events -> SOGo CalDAV PUT.

Only calendars owned by the mailbox are touched (ISC-56/83). The default calendar goes
to ``Calendar/personal``; others get a MKCALENDAR'd collection named after their slug.
Each Graph singleInstance or seriesMaster becomes one ``.ics`` resource; exceptions are
folded into the master's resource by ``calendar_conv``. State keeps the Graph
``lastModifiedDateTime`` so unchanged events are not PUT again (ISC-81).
"""

from __future__ import annotations

import logging
import time
from datetime import UTC, datetime, timedelta
from urllib.parse import quote

from . import calendar_conv
from .config import Config, MailboxMapping
from .dav import DavError, SogoDav, slugify
from .graph import GraphClient, GraphError
from .report import (
    CollectionPlan,
    CollectionResult,
    CollectionsPlan,
    CollectionsResult,
    CollectionsVerify,
    CollectionVerify,
    NullProgress,
    Progress,
)
from .state import STATUS_DONE, STATUS_FAILED, State

log = logging.getLogger(__name__)

# SOGo answers 403 for a user it cannot resolve; in mailcow that is what happens for a
# domain added after SOGo last started (its user sources are generated at start-up).
SOGO_403_HINT = (" (SOGo does not know this user: if the domain was added to mailcow "
                 "recently, restart SOGo via E-Mail > Restart SOGo, then run again)")

PREFER_UTC = {"Prefer": 'outlook.timezone="UTC"'}
EVENT_SELECT = (
    "id,iCalUId,subject,body,bodyPreview,start,end,isAllDay,originalStartTimeZone,"
    "originalEndTimeZone,recurrence,type,seriesMasterId,attendees,organizer,isReminderOn,"
    "reminderMinutesBeforeStart,sensitivity,showAs,location,isCancelled,categories,"
    "createdDateTime,lastModifiedDateTime,onlineMeeting,webLink"
)
# instances additionally need the original start for RECURRENCE-ID
INSTANCE_SELECT = EVENT_SELECT + ",originalStart"
KEEP_TYPES = ("singleInstance", "seriesMaster")
DEFAULT_SLUG = "personal"
PAGE_SIZE = 100


def iso_utc(value: datetime) -> str:
    return value.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


class CalendarMigrator:
    """Plan, migrate and verify the calendars of one mailbox."""

    kind = "calendar"

    def __init__(self, cfg: Config, graph: GraphClient, state: State, dav: SogoDav | None,
                 mapping: MailboxMapping, dry_run: bool,
                 progress: Progress | NullProgress | None = None,
                 now: datetime | None = None) -> None:
        self._cfg = cfg
        self._graph = graph
        self._state = state
        self._dav = dav
        self._src = mapping.source
        self._dry_run = dry_run
        self._progress = progress or NullProgress()
        self._key = f"{mapping.source} calendar"
        now = now or datetime.now(UTC)
        self._window = (now - timedelta(days=cfg.calendar_exceptions_from_days),
                        now + timedelta(days=cfg.calendar_exceptions_to_days))
        self._seen_warnings: set[str] = set()

    def _require_dav(self) -> SogoDav:
        if self._dav is None:
            raise RuntimeError("this operation needs a DAV client (not available in dry run)")
        return self._dav

    # -- listing -----------------------------------------------------------------------

    @staticmethod
    def _owner(cal: dict) -> str:
        return str((cal.get("owner") or {}).get("address") or "").lower()

    def _owned(self, cal: dict, mine: set[str]) -> bool:
        """A calendar is the mailbox's own when it is the default one, when Graph says the
        user may share it (only the creator can), or when its owner address is one of the
        mailbox's addresses. ``mine`` holds the configured source address plus the owner
        address Graph reports on the default calendar, so a UPN that differs from the
        primary SMTP address does not turn every secondary calendar into "shared"."""
        if cal.get("isDefaultCalendar") or cal.get("canShare") is True:
            return True
        return self._owner(cal) in mine

    def _calendars(self) -> tuple[list[tuple[dict, str]], list[str]]:
        """Owned calendars with their destination slug, plus a list of skipped shared ones."""
        owned: list[tuple[dict, str]] = []
        skipped: list[str] = []
        # Slugs are remembered per Graph calendar id (ISC-57): two calendars with the
        # same name keep their own destinations across runs whatever the listing order.
        known = self._state.collection_slugs(self._src, self.kind)
        used = {DEFAULT_SLUG, *known.values()}
        calendars = list(self._graph.iter_pages(
            f"/users/{quote(self._src, safe='@')}/calendars", params={"$top": PAGE_SIZE}))
        mine = {self._src.lower()}
        mine.update(self._owner(c) for c in calendars if c.get("isDefaultCalendar"))
        mine.discard("")
        for cal in calendars:
            name = str(cal.get("name") or "Calendar")
            if not self._owned(cal, mine):
                owner = (cal.get("owner") or {}).get("address") or "unknown owner"
                skipped.append(f"{name} (shared by {owner})")
                continue
            if cal.get("isDefaultCalendar"):
                owned.append((cal, DEFAULT_SLUG))
                continue
            slug = known.get(cal["id"])
            if slug is None:
                base = slugify(name)
                slug, n = base, 1
                while slug in used:
                    n += 1
                    slug = f"{base}-{n}"
                used.add(slug)
                if not self._dry_run:  # plan creates nothing, not even state rows
                    self._state.set_collection_slug(self._src, self.kind, cal["id"], slug)
            owned.append((cal, slug))
        return owned, skipped

    def _events(self, cal_id: str):
        for ev in self._graph.iter_pages(
                f"/users/{quote(self._src, safe='@')}/calendars/{quote(cal_id, safe='')}/events",
                params={"$select": EVENT_SELECT, "$top": PAGE_SIZE}, headers=PREFER_UTC):
            if ev.get("type", "singleInstance") in KEEP_TYPES:  # ISC-105
                yield ev

    def _instances(self, event_id: str) -> list[dict]:
        start, end = self._window
        return list(self._graph.iter_pages(
            f"/users/{quote(self._src, safe='@')}/events/{quote(event_id, safe='')}/instances",
            params={"startDateTime": iso_utc(start), "endDateTime": iso_utc(end),
                    "$select": INSTANCE_SELECT, "$top": PAGE_SIZE},
            headers=PREFER_UTC))

    @staticmethod
    def _state_key(ev: dict) -> str:
        return str(ev.get("iCalUId") or ev["id"])

    # -- plan --------------------------------------------------------------------------

    def plan(self) -> CollectionsPlan:
        owned, skipped = self._calendars()
        out = CollectionsPlan(mailbox=self._src, kind=self.kind, skipped=skipped)
        for cal, slug in owned:
            count = sum(1 for _ in self._events(cal["id"]))
            out.collections.append(CollectionPlan(cal["id"], str(cal.get("name")), slug, count))
        return out

    # -- migrate -----------------------------------------------------------------------

    def migrate(self) -> CollectionsResult:
        started = time.monotonic()
        result = CollectionsResult(mailbox=self._src, kind=self.kind, dry_run=self._dry_run)
        try:
            self._migrate(result)
        except GraphError as exc:
            result.errors.append(f"listing calendars failed: {exc}")
        finally:
            result.duration_s = round(time.monotonic() - started, 3)
            self._progress.finish(self._key)
        return result

    def _migrate(self, result: CollectionsResult) -> None:
        if not self._dry_run:
            try:
                home = self._require_dav().calendar_home_exists()
            except DavError as exc:
                result.errors.append(f"SOGo calendar home check failed: {exc}"
                                     + (SOGO_403_HINT if exc.status == 403 else ""))
                return
            if not home:  # ISC-107
                result.errors.append(
                    f"SOGo calendar home for {self._src} not found (404); "
                    "skipping calendars for this mailbox")
                return
        owned, result.skipped = self._calendars()
        for cal, slug in owned:
            cr = CollectionResult(cal["id"], str(cal.get("name")), slug)
            result.collections.append(cr)
            self._migrate_calendar(cal, cr, result)

    def _migrate_calendar(self, cal: dict, cr: CollectionResult,
                          result: CollectionsResult) -> None:
        if not self._dry_run and cr.slug != DEFAULT_SLUG:
            try:
                self._require_dav().ensure_calendar(cr.name, slug=cr.slug)
            except DavError as exc:
                cr.error = f"MKCALENDAR refused: {exc}"
                return
        try:
            events = list(self._events(cal["id"]))
        except GraphError as exc:
            cr.error = f"listing events failed: {exc}"
            return
        self._progress.start(self._key, len(events))
        for ev in events:
            cr.listed += 1
            self._migrate_event(ev, cr, result)
            self._progress.advance(self._key)

    def _migrate_event(self, ev: dict, cr: CollectionResult, result: CollectionsResult) -> None:
        key = self._state_key(ev)
        last_modified = ev.get("lastModifiedDateTime")
        if last_modified and self._state.event_last_modified(
                self._src, cr.slug, key) == last_modified:
            cr.unchanged += 1
            return
        if self._dry_run:
            cr.would_put += 1
            return
        try:
            instances = self._instances(ev["id"]) if ev.get("type") == "seriesMaster" else []
        except GraphError as exc:
            self._fail(cr, key, last_modified, f"instances: {exc}")
            return
        try:
            conv = calendar_conv.convert_event(ev, instances, window=self._window,
                                               attendees=self._cfg.calendar_attendees)
        except Exception as exc:  # a converter bug must not end the whole run
            self._fail(cr, key, last_modified,
                       f"conversion failed: {exc.__class__.__name__}: {exc}")
            return
        for warning in conv.warnings:
            if warning not in self._seen_warnings:  # ISC-62: once per run
                self._seen_warnings.add(warning)
                result.warnings.append(warning)
                log.warning("%s: %s", self._src, warning)
        try:
            self._require_dav().put_event(cr.slug, conv.uid, conv.ics)
        except DavError as exc:  # ISC-108
            self._fail(cr, key, last_modified, str(exc))
            return
        self._state.mark_event(self._src, cr.slug, key, last_modified, STATUS_DONE)
        cr.put += 1

    def _fail(self, cr: CollectionResult, key: str, last_modified: str | None,
              error: str) -> None:
        self._state.mark_event(self._src, cr.slug, key, last_modified, STATUS_FAILED, error)
        cr.failed += 1
        log.warning("%s: event in %s failed: %s", self._src, cr.name, error)

    # -- verify ------------------------------------------------------------------------

    def verify(self) -> CollectionsVerify:
        out = CollectionsVerify(mailbox=self._src, kind=self.kind)
        try:
            owned, out.skipped = self._calendars()
        except GraphError as exc:
            out.errors.append(f"listing calendars failed: {exc}")
            return out
        counts = self._state.event_counts(self._src)
        for cal, slug in owned:
            name = str(cal.get("name"))
            try:
                graph_count = sum(1 for _ in self._events(cal["id"]))
            except GraphError as exc:
                out.errors.append(f"{name}: listing events failed: {exc}")
                continue
            try:
                dav_count = self._require_dav().count_resources("Calendar", slug)
            except DavError as exc:
                if exc.status != 404:
                    out.errors.append(f"{name}: {exc}")
                    continue
                dav_count = 0
            c = counts.get(slug, {})
            failed = c.get(STATUS_FAILED, 0)
            expected = graph_count - failed
            out.collections.append(CollectionVerify(
                name=name, slug=slug, graph_count=graph_count, done=c.get(STATUS_DONE, 0),
                failed=failed, dav_count=dav_count, expected=expected,
                mismatch=dav_count != expected))
        return out
