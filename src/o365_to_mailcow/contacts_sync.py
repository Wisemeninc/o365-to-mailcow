"""Contact migration: Graph contact folders -> SOGo CardDAV PUT.

The default contact folder goes to ``Contacts/personal``. Every other folder (found
recursively via ``childFolders``, ISC-127) gets its own address book through extended
MKCOL; if SOGo refuses, its contacts go to ``personal`` tagged with the folder name as a
category and the fallback is reported (ISC-84). Contacts are idempotent on Graph
``lastModifiedDateTime``.
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from urllib.parse import quote

from . import contacts_conv
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

DEFAULT_SLUG = "personal"
RESERVED_SLUGS = {"personal", "collected"}  # SOGo's built-in address books
DEFAULT_ID = "default"
PAGE_SIZE = 100


@dataclass
class ContactFolder:
    folder_id: str
    name: str
    contacts_path: str
    slug: str
    is_default: bool = False


class ContactsMigrator:
    """Plan, migrate and verify the contacts of one mailbox."""

    kind = "contacts"

    def __init__(self, cfg: Config, graph: GraphClient, state: State, dav: SogoDav | None,
                 mapping: MailboxMapping, dry_run: bool,
                 progress: Progress | NullProgress | None = None) -> None:
        self._cfg = cfg
        self._graph = graph
        self._state = state
        self._dav = dav
        self._src = mapping.source
        self._dry_run = dry_run
        self._progress = progress or NullProgress()
        self._key = f"{mapping.source} contacts"
        self._photo_errors = 0

    def _require_dav(self) -> SogoDav:
        if self._dav is None:
            raise RuntimeError("this operation needs a DAV client (not available in dry run)")
        return self._dav

    # -- listing -----------------------------------------------------------------------

    def _folders(self) -> list[ContactFolder]:
        base = f"/users/{quote(self._src, safe='@')}"
        out = [ContactFolder(DEFAULT_ID, "Contacts", f"{base}/contacts", DEFAULT_SLUG, True)]
        known = self._state.collection_slugs(self._src, self.kind)
        used = set(RESERVED_SLUGS) | set(known.values())
        stack = list(reversed(list(self._graph.iter_pages(
            f"{base}/contactFolders", params={"$top": PAGE_SIZE}))))
        while stack:
            folder = stack.pop()
            name = str(folder.get("displayName") or "Contacts")
            fid = folder["id"]
            slug = known.get(fid)
            if slug is None:
                slug_base = slugify(name)
                slug, n = slug_base, 1
                while slug in used:
                    n += 1
                    slug = f"{slug_base}-{n}"
                used.add(slug)
                if not self._dry_run:  # plan creates nothing, not even state rows
                    self._state.set_collection_slug(self._src, self.kind, fid, slug)
            qfid = quote(fid, safe="")
            out.append(ContactFolder(fid, name, f"{base}/contactFolders/{qfid}/contacts", slug))
            children = list(self._graph.iter_pages(
                f"{base}/contactFolders/{qfid}/childFolders", params={"$top": PAGE_SIZE}))
            stack.extend(reversed(children))
        return out

    def _contacts(self, folder: ContactFolder):
        return self._graph.iter_pages(folder.contacts_path, params={"$top": PAGE_SIZE})

    def _photo(self, contact_id: str) -> tuple[bytes | None, bool]:
        """Return ``(photo, ok)``; ``ok`` is False when the fetch failed for a reason other
        than "no photo", so the caller can leave the contact eligible for a retry."""
        try:
            user = quote(self._src, safe="@")
            cid = quote(contact_id, safe="")
            data = self._graph.get_bytes(f"/users/{user}/contacts/{cid}/photo/$value")
        except GraphError as exc:
            if exc.status != 404:  # 404 = no photo (ISC-111); anything else: go on without
                self._photo_errors += 1
                log.warning("%s: photo for a contact unavailable: %s", self._src, exc)
                return None, False
            return None, True
        return data or None, True

    # -- plan --------------------------------------------------------------------------

    def plan(self) -> CollectionsPlan:
        out = CollectionsPlan(mailbox=self._src, kind=self.kind)
        for folder in self._folders():
            count = sum(1 for _ in self._contacts(folder))
            out.collections.append(
                CollectionPlan(folder.folder_id, folder.name, folder.slug, count))
        return out

    # -- migrate -----------------------------------------------------------------------

    def migrate(self) -> CollectionsResult:
        started = time.monotonic()
        result = CollectionsResult(mailbox=self._src, kind=self.kind, dry_run=self._dry_run)
        try:
            self._migrate(result)
        except GraphError as exc:
            result.errors.append(f"listing contact folders failed: {exc}")
        finally:
            if self._photo_errors:
                result.warnings.append(f"{self._photo_errors} contact photo(s) could not be "
                                       "fetched and were left out")
            result.duration_s = round(time.monotonic() - started, 3)
            self._progress.finish(self._key)
        return result

    def _migrate(self, result: CollectionsResult) -> None:
        if not self._dry_run:
            try:
                home = self._require_dav().addressbook_home_exists()
            except DavError as exc:
                result.errors.append(f"SOGo address book home check failed: {exc}"
                                     + (SOGO_403_HINT if exc.status == 403 else ""))
                return
            if not home:  # ISC-107
                result.errors.append(
                    f"SOGo address book home for {self._src} not found (404); "
                    "skipping contacts for this mailbox")
                return
        for folder in self._folders():
            category: str | None = None
            slug = folder.slug
            if not self._dry_run and not folder.is_default:
                try:
                    slug = self._require_dav().ensure_addressbook(folder.name, slug=folder.slug)
                except DavError as exc:
                    slug, category = DEFAULT_SLUG, folder.name
                    result.fallbacks.append(
                        f"address book {folder.name!r} could not be created ({exc.status}); "
                        f"its contacts went to 'personal' with category {folder.name!r}")
            cr = CollectionResult(folder.folder_id, folder.name, slug)
            result.collections.append(cr)
            try:
                contacts = list(self._contacts(folder))
            except GraphError as exc:
                cr.error = f"listing contacts failed: {exc}"
                continue
            self._progress.start(self._key, len(contacts))
            for contact in contacts:
                cr.listed += 1
                self._migrate_contact(contact, cr, category)
                self._progress.advance(self._key)

    def _migrate_contact(self, contact: dict, cr: CollectionResult,
                         category: str | None) -> None:
        gid = contact["id"]
        last_modified = contact.get("lastModifiedDateTime")
        if last_modified and self._state.contact_last_modified(self._src, gid) == last_modified:
            cr.unchanged += 1
            return
        if self._dry_run:
            cr.would_put += 1
            return
        if category:
            cats = list(contact.get("categories") or [])
            if category not in cats:
                cats.append(category)
            contact = {**contact, "categories": cats}
        photo, photo_ok = self._photo(gid) if self._cfg.contacts_photos else (None, True)
        try:
            conv = contacts_conv.convert_contact(contact, photo)
        except Exception as exc:  # a converter bug must not end the whole run
            self._fail(cr, gid, last_modified, f"conversion failed: {exc.__class__.__name__}")
            return
        try:
            self._require_dav().put_contact(cr.slug, conv.uid, conv.vcf)
        except DavError as exc:
            self._fail(cr, gid, last_modified, str(exc))
            return
        # a contact whose photo could not be fetched is stored without its last_modified
        # so the next run tries again (the vCard itself is complete apart from PHOTO)
        self._state.mark_contact(self._src, gid, cr.slug,
                                 last_modified if photo_ok else None, STATUS_DONE)
        cr.put += 1

    def _fail(self, cr: CollectionResult, gid: str, last_modified: str | None,
              error: str) -> None:
        self._state.mark_contact(self._src, gid, cr.slug, last_modified, STATUS_FAILED, error)
        cr.failed += 1
        log.warning("%s: contact in %s failed: %s", self._src, cr.name, error)

    # -- verify ------------------------------------------------------------------------

    def verify(self) -> CollectionsVerify:
        """Per destination address book: Graph count (summed over the folders that feed
        it) against the CardDAV resource count (ISC-91)."""
        out = CollectionsVerify(mailbox=self._src, kind=self.kind)
        try:
            folders = self._folders()
            existing = {s for s, _ in self._require_dav().list_addressbooks()}
        except (GraphError, DavError) as exc:
            out.errors.append(f"listing address books failed: {exc}")
            return out
        per_book: dict[str, tuple[list[str], int]] = {}
        for folder in folders:
            slug = folder.slug
            if not folder.is_default and slug not in existing:
                out.fallbacks.append(f"{folder.name!r} has no own address book; "
                                     "counted under 'personal'")
                slug = DEFAULT_SLUG
            try:
                count = sum(1 for _ in self._contacts(folder))
            except GraphError as exc:
                out.errors.append(f"{folder.name}: listing contacts failed: {exc}")
                continue
            names, total = per_book.get(slug, ([], 0))
            per_book[slug] = ([*names, folder.name], total + count)
        counts = self._state.contact_counts(self._src)
        for slug, (names, graph_count) in per_book.items():
            try:
                dav_count = self._require_dav().count_resources("Contacts", slug)
            except DavError as exc:
                if exc.status != 404:
                    out.errors.append(f"{slug}: {exc}")
                    continue
                dav_count = 0
            c = counts.get(slug, {})
            failed = c.get(STATUS_FAILED, 0)
            expected = graph_count - failed
            out.collections.append(CollectionVerify(
                name=" + ".join(names), slug=slug, graph_count=graph_count,
                done=c.get(STATUS_DONE, 0), failed=failed, dav_count=dav_count,
                expected=expected, mismatch=dav_count != expected))
        return out
