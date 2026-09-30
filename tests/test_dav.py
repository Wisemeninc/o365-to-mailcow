"""SOGo DAV client over a mocked transport."""

from __future__ import annotations

import base64

import pytest
import responses

from o365_to_mailcow.dav import DavError, SogoDav, slugify

BASE = "https://mail.example.net/SOGo/dav/alice@example.net/"
PW = "app-password-xyz"


def multistatus(*items: tuple[str, str | None, bool, str | None]) -> str:
    """items: (href, displayname, is_collection, extra resourcetype element)."""
    parts = []
    for href, name, coll, extra in items:
        rt = ("<d:collection/>" if coll else "") + (extra or "")
        dn = f"<d:displayname>{name}</d:displayname>" if name else ""
        parts.append(
            f"<d:response><d:href>{href}</d:href><d:propstat><d:prop>{dn}"
            f"<d:resourcetype>{rt}</d:resourcetype></d:prop>"
            "<d:status>HTTP/1.1 200 OK</d:status></d:propstat></d:response>")
    return ('<?xml version="1.0" encoding="utf-8"?><d:multistatus xmlns:d="DAV:" '
            'xmlns:c="urn:ietf:params:xml:ns:caldav">' + "".join(parts) + "</d:multistatus>")


HOME = "/SOGo/dav/alice@example.net/Calendar/"
CALS = multistatus(
    (HOME, None, True, None),
    (HOME + "personal/", "Personal Calendar", True, "<c:calendar/>"),
    (HOME + "holidays/", "Holidays", True, "<c:calendar/>"),
)


@pytest.fixture
def dav():
    return SogoDav("mail.example.net", "alice@example.net", PW)


def test_slugify():
    assert slugify("Team Calendar (2024)!") == "team-calendar-2024"
    assert slugify("Ümlaut") == "mlaut"
    assert slugify("***") == "collection"
    assert len(slugify("x" * 100)) == 40
    assert slugify("a" * 39 + " b") == "a" * 39


@responses.activate
def test_home_exists_isc_107(dav):
    responses.add("PROPFIND", BASE + "Calendar/", status=207, body=multistatus())
    responses.add("PROPFIND", BASE + "Contacts/", status=404)
    assert dav.calendar_home_exists() is True
    assert dav.addressbook_home_exists() is False
    req = responses.calls[0].request
    assert req.headers["Depth"] == "0"
    user, pw = base64.b64decode(req.headers["Authorization"].split()[1]).decode().split(":")
    assert (user, pw) == ("alice@example.net", PW)


@responses.activate
def test_home_check_auth_error_raises(dav):
    responses.add("PROPFIND", BASE + "Calendar/", status=401, body="Unauthorized")
    with pytest.raises(DavError) as exc:
        dav.calendar_home_exists()
    assert exc.value.status == 401 and PW not in str(exc.value)


@responses.activate
def test_list_calendars(dav):
    responses.add("PROPFIND", BASE + "Calendar/", status=207, body=CALS)
    assert dav.list_calendars() == [("personal", "Personal Calendar"), ("holidays", "Holidays")]
    assert responses.calls[0].request.headers["Depth"] == "1"


@responses.activate
def test_ensure_calendar_existing_makes_no_mkcalendar(dav):
    responses.add("PROPFIND", BASE + "Calendar/", status=207, body=CALS)
    assert dav.ensure_calendar("Holidays") == "holidays"
    assert [c.request.method for c in responses.calls] == ["PROPFIND"]


@responses.activate
def test_ensure_calendar_creates_with_displayname_isc_57(dav):
    responses.add("PROPFIND", BASE + "Calendar/", status=207, body=CALS)
    responses.add("MKCALENDAR", BASE + "Calendar/team-a-b/", status=201)
    assert dav.ensure_calendar("Team <A&B>") == "team-a-b"
    body = responses.calls[1].request.body.decode()
    assert "<d:displayname>Team &lt;A&amp;B&gt;</d:displayname>" in body
    assert "urn:ietf:params:xml:ns:caldav" in body


@responses.activate
def test_ensure_calendar_refused(dav):
    responses.add("PROPFIND", BASE + "Calendar/", status=207, body=CALS)
    responses.add("MKCALENDAR", BASE + "Calendar/work/", status=405, body="Not Allowed")
    with pytest.raises(DavError) as exc:
        dav.ensure_calendar("Work")
    assert exc.value.status == 405


@responses.activate
def test_put_event_url_and_content_type_isc_80(dav):
    responses.put(BASE + "Calendar/personal/040000008200E00074C5B7101A82E008%2Fx.ics",
                  status=201)
    dav.put_event("personal", "040000008200E00074C5B7101A82E008/x", b"BEGIN:VCALENDAR")
    req = responses.calls[0].request
    assert req.headers["Content-Type"] == "text/calendar; charset=utf-8"
    assert req.body == b"BEGIN:VCALENDAR"


@responses.activate
def test_put_4xx_error_body_trimmed_no_headers_isc_108(dav):
    responses.put(BASE + "Calendar/personal/u1.ics", status=412,
                  body="Precondition Failed " + "y" * 400,
                  headers={"WWW-Authenticate": "Basic realm=secret-header"})
    with pytest.raises(DavError) as exc:
        dav.put_event("personal", "u1", b"x")
    assert exc.value.status == 412
    assert len(exc.value.body) == 200
    assert "secret-header" not in str(exc.value) and PW not in str(exc.value)


@responses.activate
def test_count_resources_isc_82(dav):
    body = multistatus(
        (HOME + "personal/", "Personal", True, "<c:calendar/>"),
        (HOME + "personal/a.ics", None, False, None),
        (HOME + "personal/b.ics", None, False, None),
    )
    responses.add("PROPFIND", BASE + "Calendar/personal/", status=207, body=body)
    assert dav.count_resources("Calendar", "personal") == 2
    responses.add("PROPFIND", BASE + "Contacts/nope/", status=404)
    with pytest.raises(DavError) as exc:
        dav.count_resources("contacts", "nope")
    assert exc.value.status == 404


@responses.activate
def test_ensure_addressbook_extended_mkcol_isc_84(dav):
    responses.add("PROPFIND", BASE + "Contacts/", status=207, body=multistatus())
    responses.add("MKCOL", BASE + "Contacts/suppliers/", status=201)
    assert dav.ensure_addressbook("Suppliers") == "suppliers"
    body = responses.calls[1].request.body.decode()
    assert "<d:collection/><card:addressbook/>" in body
    assert "urn:ietf:params:xml:ns:carddav" in body


@pytest.mark.parametrize("status", [403, 405, 415])
@responses.activate
def test_ensure_addressbook_refused(dav, status):
    responses.add("PROPFIND", BASE + "Contacts/", status=207, body=multistatus())
    responses.add("MKCOL", BASE + "Contacts/suppliers/", status=status)
    with pytest.raises(DavError):
        dav.ensure_addressbook("Suppliers")


@responses.activate
def test_put_contact_isc_90(dav):
    responses.put(BASE + "Contacts/personal/abc.vcf", status=204)
    dav.put_contact("personal", "abc", b"BEGIN:VCARD")
    assert responses.calls[0].request.headers["Content-Type"] == "text/vcard; charset=utf-8"


def test_requests_stay_on_mailcow_host():
    seen: list[tuple[str, dict]] = []

    class Capture:
        def request(self, method, url, **kw):
            seen.append((url, kw))
            raise AssertionError("stop")

    d = SogoDav("mail.example.net", "x@evil.example/../#", "pw", session=Capture())
    with pytest.raises(AssertionError):
        d.put_event("personal", "u", b"x")
    url, kw = seen[0]
    assert url.startswith("https://mail.example.net/SOGo/dav/")
    assert kw["timeout"] and kw["verify"] is True
