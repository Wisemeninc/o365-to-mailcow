"""SOGo CalDAV/CardDAV client (just enough for PUT-based migration).

Every URL is built from ``https://{host}/SOGo/dav/{user}/`` and checked against the
configured host before the request is sent. Errors carry the status and the first 200
characters of the body, never headers (ISC-108), so credentials cannot leak.
"""

from __future__ import annotations

import logging
import re
import xml.etree.ElementTree as ET  # noqa: S405 - parses replies from our own mailcow host
from urllib.parse import quote, unquote, urlsplit
from xml.sax.saxutils import escape

import requests

from . import USER_AGENT

log = logging.getLogger(__name__)

TIMEOUT = (15.0, 120.0)
DAV = "{DAV:}"
CALDAV_NS = "urn:ietf:params:xml:ns:caldav"
CARDDAV_NS = "urn:ietf:params:xml:ns:carddav"
KINDS = {"calendar": "Calendar", "contacts": "Contacts"}

PROPFIND_BODY = (
    b'<?xml version="1.0" encoding="utf-8"?>'
    b'<d:propfind xmlns:d="DAV:"><d:prop><d:displayname/><d:resourcetype/></d:prop>'
    b"</d:propfind>"
)


class DavError(Exception):
    """A DAV request failed. ``status`` is 0 for network errors."""

    def __init__(self, status: int, message: str) -> None:
        super().__init__(f"DAV HTTP {status}: {message[:200]}")
        self.status = status
        self.body = message[:200]


def slugify(name: str, max_len: int = 40) -> str:
    """Lowercase, runs of anything but ``[a-z0-9]`` become ``-``, trimmed, max 40 chars."""
    slug = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")[:max_len].strip("-")
    return slug or "collection"


class SogoDav:
    """CalDAV/CardDAV access to one SOGo user with Basic auth (an app password)."""

    def __init__(self, host: str, user: str, password: str,
                 session: requests.Session | None = None, verify: bool = True) -> None:
        self._host = host.lower()
        self._user = user
        self._base = f"https://{host}/SOGo/dav/{quote(user, safe='@')}/"
        self._session = session or requests.Session()
        self._auth = (user, password)
        self._verify = verify

    # -- transport ---------------------------------------------------------------------

    def _request(self, method: str, path: str, body: bytes | None = None,
                 headers: dict[str, str] | None = None) -> requests.Response:
        url = self._base + path
        if urlsplit(url).netloc.lower() != self._host:
            raise DavError(0, "refusing request to a host other than the mailcow host")
        hdrs = {"User-Agent": USER_AGENT}
        if headers:
            hdrs.update(headers)
        try:
            return self._session.request(method, url, data=body, headers=hdrs,
                                         auth=self._auth, timeout=TIMEOUT,
                                         verify=self._verify)
        except requests.RequestException as exc:
            raise DavError(0, f"network error: {exc.__class__.__name__}") from exc

    @staticmethod
    def _fail(resp: requests.Response) -> DavError:
        return DavError(resp.status_code, resp.text[:200])

    def _propfind(self, path: str, depth: str) -> requests.Response:
        return self._request("PROPFIND", path, PROPFIND_BODY,
                             {"Depth": depth, "Content-Type": "application/xml; charset=utf-8"})

    @staticmethod
    def _parse(xml: bytes) -> list[tuple[str, str | None, set[str]]]:
        """Multistatus -> ``[(href, displayname, resourcetype tags)]``."""
        # Replies come from the configured mailcow host over verified TLS; Python's expat
        # (>= 2.4) refuses entity expansion attacks and ElementTree resolves no externals.
        root = ET.fromstring(xml)  # noqa: S314  # nosec B314
        out: list[tuple[str, str | None, set[str]]] = []
        for resp in root.iter(f"{DAV}response"):
            href = (resp.findtext(f"{DAV}href") or "").strip()
            name: str | None = None
            types: set[str] = set()
            for propstat in resp.iter(f"{DAV}propstat"):
                status = propstat.findtext(f"{DAV}status") or ""
                if " 200 " not in f"{status} ":
                    continue
                prop = propstat.find(f"{DAV}prop")
                if prop is None:
                    continue
                dn = prop.find(f"{DAV}displayname")
                if dn is not None and dn.text:
                    name = dn.text
                rt = prop.find(f"{DAV}resourcetype")
                if rt is not None:
                    types.update(child.tag for child in rt)
            out.append((href, name, types))
        return out

    def _home_exists(self, kind: str) -> bool:
        resp = self._propfind(f"{kind}/", "0")
        if resp.status_code == 404:
            return False
        if resp.status_code != 207:
            raise self._fail(resp)
        return True

    def _collections(self, kind: str) -> list[tuple[str, str]]:
        resp = self._propfind(f"{kind}/", "1")
        if resp.status_code != 207:
            raise self._fail(resp)
        home = urlsplit(self._base + f"{kind}/").path.rstrip("/")
        out: list[tuple[str, str]] = []
        for href, name, types in self._parse(resp.content):
            path = unquote(urlsplit(href).path).rstrip("/")
            if path == unquote(home) or f"{DAV}collection" not in types:
                continue
            slug = path.rsplit("/", 1)[-1]
            out.append((slug, name or slug))
        return out

    # -- homes -------------------------------------------------------------------------

    def calendar_home_exists(self) -> bool:
        return self._home_exists("Calendar")

    def addressbook_home_exists(self) -> bool:
        return self._home_exists("Contacts")

    # -- calendars ---------------------------------------------------------------------

    def list_calendars(self) -> list[tuple[str, str]]:
        return self._collections("Calendar")

    def ensure_calendar(self, name: str, slug: str | None = None) -> str:
        """Return the slug of a calendar collection, creating it with MKCALENDAR."""
        slug = slug or slugify(name)
        if any(s == slug for s, _ in self.list_calendars()):
            return slug
        body = (
            '<?xml version="1.0" encoding="utf-8"?>'
            f'<c:mkcalendar xmlns:d="DAV:" xmlns:c="{CALDAV_NS}"><d:set><d:prop>'
            f"<d:displayname>{escape(name)}</d:displayname>"
            "</d:prop></d:set></c:mkcalendar>"
        ).encode()
        resp = self._request("MKCALENDAR", f"Calendar/{quote(slug)}/", body,
                             {"Content-Type": "application/xml; charset=utf-8"})
        if resp.status_code not in (200, 201):
            raise self._fail(resp)
        log.info("created calendar %s (%s)", slug, name)
        return slug

    def put_event(self, slug: str, uid: str, ics: bytes) -> None:
        self._put(f"Calendar/{quote(slug)}/{quote(uid, safe='')}.ics", ics,
                  "text/calendar; charset=utf-8")

    # -- address books -----------------------------------------------------------------

    def list_addressbooks(self) -> list[tuple[str, str]]:
        return self._collections("Contacts")

    def ensure_addressbook(self, name: str, slug: str | None = None) -> str:
        """Return the slug of an address book, creating it with extended MKCOL (RFC 5689)."""
        slug = slug or slugify(name)
        if any(s == slug for s, _ in self.list_addressbooks()):
            return slug
        body = (
            '<?xml version="1.0" encoding="utf-8"?>'
            f'<d:mkcol xmlns:d="DAV:" xmlns:card="{CARDDAV_NS}"><d:set><d:prop>'
            "<d:resourcetype><d:collection/><card:addressbook/></d:resourcetype>"
            f"<d:displayname>{escape(name)}</d:displayname>"
            "</d:prop></d:set></d:mkcol>"
        ).encode()
        resp = self._request("MKCOL", f"Contacts/{quote(slug)}/", body,
                             {"Content-Type": "application/xml; charset=utf-8"})
        if resp.status_code not in (200, 201):
            raise self._fail(resp)
        log.info("created address book %s (%s)", slug, name)
        return slug

    def put_contact(self, slug: str, uid: str, vcf: bytes) -> None:
        self._put(f"Contacts/{quote(slug)}/{quote(uid, safe='')}.vcf", vcf,
                  "text/vcard; charset=utf-8")

    # -- shared ------------------------------------------------------------------------

    def _put(self, path: str, data: bytes, content_type: str) -> None:
        resp = self._request("PUT", path, data, {"Content-Type": content_type})
        if resp.status_code not in (200, 201, 204):
            raise self._fail(resp)

    def count_resources(self, kind: str, slug: str) -> int:
        """Number of non-collection members of ``{Calendar|Contacts}/{slug}/``."""
        folder = KINDS.get(kind.lower(), kind)
        resp = self._propfind(f"{folder}/{quote(slug)}/", "1")
        if resp.status_code != 207:
            raise self._fail(resp)
        return sum(1 for _, _, types in self._parse(resp.content)
                   if f"{DAV}collection" not in types)
