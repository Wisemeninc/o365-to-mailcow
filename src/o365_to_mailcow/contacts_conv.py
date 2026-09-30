"""Graph contact to vCard 3.0 conversion.

The serializer is intentionally hand-written instead of using a runtime vCard library so
the project can control escaping, UTF-8 octet folding, and field-level mapping behavior
exactly. Mapping choices are:

- Outlook's birthday year sentinel ``1604`` means "no year" and is exported as
  ``BDAY:--MMDD``.
- Exchange legacy distinguished-name addresses (no ``@``) are dropped from ``EMAIL``.
- ``FN`` fallback chain is: ``displayName`` -> ``givenName + surname`` -> first valid
  email -> ``(no name)``. ``N`` is always emitted because vCard 3.0 requires it.
- ``UID`` is a UUID5 of the Graph contact ``id`` so a re-run PUTs to the same resource.
- Graph photos are JPEG, so ``PHOTO`` is always ``TYPE=JPEG`` with inline base64.
- ``REV`` is ``lastModifiedDateTime`` normalised to whole-second UTC.
- Malformed optional fields are skipped with a log warning (never the value itself);
  only a missing ``id`` is fatal.
"""

from __future__ import annotations

import base64
import logging
import re
import uuid
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from datetime import UTC

from dateutil import parser as dateutil_parser

log = logging.getLogger(__name__)

# C0 controls other than TAB/LF/CR (and DEL) are invalid in vCard text and are dropped.
_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
# URI values cannot carry an escaped newline, so every C0 control (and DEL) is dropped.
_URI_CONTROL_CHARS = re.compile(r"[\x00-\x1f\x7f]")


@dataclass(frozen=True)
class ConvertedContact:
    """A vCard ready to PUT: ``uid`` names the resource, ``last_modified`` feeds state."""

    uid: str
    vcf: bytes
    last_modified: str | None


def contact_uid(graph_id: str) -> str:
    """Deterministic vCard UID for a Graph contact id (UUID5 in the URL namespace)."""
    if not isinstance(graph_id, str) or not graph_id.strip():
        raise ValueError("graph_id must be a non-empty string")
    seed = f"o365-to-mailcow:contact:{graph_id}"
    return str(uuid.uuid5(uuid.NAMESPACE_URL, seed))


def convert_contact(contact: dict, photo: bytes | None = None) -> ConvertedContact:
    """Convert a Graph v1.0 ``contact`` object (plus optional JPEG bytes) to vCard 3.0.

    Raises ``ValueError`` when ``id`` is missing and ``TypeError`` for a non-dict contact
    or non-bytes photo; every other field is optional.
    """
    if not isinstance(contact, dict):
        raise TypeError("contact must be a dict")

    graph_id = contact.get("id")
    if not isinstance(graph_id, str) or not graph_id.strip():
        raise ValueError("contact.id must be a non-empty string")

    if photo is not None and not isinstance(photo, (bytes, bytearray)):
        raise TypeError("photo must be bytes or None")
    photo_bytes = bytes(photo) if photo is not None else None

    uid = contact_uid(graph_id)
    lines = ["BEGIN:VCARD", "VERSION:3.0", f"UID:{uid}"]

    given_name = _optional_text(contact, "givenName")
    middle_name = _optional_text(contact, "middleName")
    surname = _optional_text(contact, "surname")
    honorific = _optional_text(contact, "title")
    generation = _optional_text(contact, "generation")
    display_name = _optional_text(contact, "displayName")

    emails = _email_addresses(contact.get("emailAddresses"))
    first_email = emails[0] if emails else None
    full_name = _full_name(display_name, given_name, surname, first_email)

    lines.append(f"FN:{_escape_text(full_name)}")
    name_parts = [surname, given_name, middle_name, honorific, generation]
    lines.append("N:" + _join_structured([part or "" for part in name_parts]))

    nickname = _optional_text(contact, "nickName")
    if nickname:
        lines.append(f"NICKNAME:{_escape_text(nickname)}")

    company_name = _optional_text(contact, "companyName")
    department = _optional_text(contact, "department")
    if company_name or department:
        lines.append("ORG:" + _join_structured([company_name or "", department or ""]))

    job_title = _optional_text(contact, "jobTitle")
    if job_title:
        lines.append(f"TITLE:{_escape_text(job_title)}")

    for email in emails:
        lines.append(f"EMAIL;TYPE=INTERNET:{_escape_text(email)}")

    _append_phone_lines(lines, "businessPhones", "WORK", contact.get("businessPhones"))
    _append_phone_lines(lines, "homePhones", "HOME", contact.get("homePhones"))

    mobile = _optional_text(contact, "mobilePhone")
    if mobile:
        lines.append(f"TEL;TYPE=CELL:{_escape_text(mobile)}")

    _append_address_line(lines, "homeAddress", "HOME", contact.get("homeAddress"))
    _append_address_line(lines, "businessAddress", "WORK", contact.get("businessAddress"))
    _append_address_line(lines, "otherAddress", "OTHER", contact.get("otherAddress"))

    bday_value = _birthday(contact.get("birthday"))
    if bday_value:
        lines.append(f"BDAY:{bday_value}")

    notes = _optional_text(contact, "personalNotes", strip=False)
    if notes:
        lines.append(f"NOTE:{_escape_text(notes)}")

    business_home_page = _optional_text(contact, "businessHomePage")
    if business_home_page:
        # URL is a URI value, not TEXT: no RFC 2426 escaping, but a control character
        # (a stray newline) would break the line structure, so those are removed.
        url = _URI_CONTROL_CHARS.sub("", business_home_page)
        if url:
            lines.append(f"URL:{url}")

    categories = _categories(contact.get("categories"))
    if categories:
        lines.append(f"CATEGORIES:{categories}")

    if photo_bytes:
        encoded = base64.b64encode(photo_bytes).decode("ascii")
        lines.append(f"PHOTO;ENCODING=b;TYPE=JPEG:{encoded}")

    rev = _normalise_rev(contact.get("lastModifiedDateTime"))
    if rev:
        lines.append(f"REV:{rev}")

    lines.append("END:VCARD")

    raw_last_modified = contact.get("lastModifiedDateTime")
    if raw_last_modified is not None and not isinstance(raw_last_modified, str):
        _warn_skip("lastModifiedDateTime", "str", raw_last_modified)
        raw_last_modified = None

    return ConvertedContact(uid=uid, vcf=_serialize_vcard(lines), last_modified=raw_last_modified)


def _optional_text(data: Mapping[str, object], field: str, *, strip: bool = True) -> str | None:
    value = data.get(field)
    if value is None:
        return None
    if not isinstance(value, str):
        _warn_skip(field, "str", value)
        return None
    text = value.strip() if strip else value
    if not text:
        return None
    return text


def _email_addresses(raw: object) -> list[str]:
    if raw is None:
        return []
    if not isinstance(raw, list):
        _warn_skip("emailAddresses", "list", raw)
        return []

    addresses: list[str] = []
    seen: set[str] = set()
    for index, entry in enumerate(raw):
        if not isinstance(entry, Mapping):
            _warn_skip(f"emailAddresses[{index}]", "mapping", entry)
            continue

        value = entry.get("address")
        if value is None:
            continue
        if not isinstance(value, str):
            _warn_skip(f"emailAddresses[{index}].address", "str", value)
            continue

        address = value.strip()
        if not address or "@" not in address:
            continue
        lowered = address.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        addresses.append(address)
    return addresses


def _append_phone_lines(lines: list[str], field: str, phone_type: str, raw: object) -> None:
    if raw is None:
        return
    if not isinstance(raw, list):
        _warn_skip(field, "list", raw)
        return
    for index, item in enumerate(raw):
        if not isinstance(item, str):
            _warn_skip(f"{field}[{index}]", "str", item)
            continue
        phone = item.strip()
        if phone:
            lines.append(f"TEL;TYPE={phone_type}:{_escape_text(phone)}")


def _append_address_line(lines: list[str], field: str, adr_type: str, raw: object) -> None:
    if raw is None:
        return
    if not isinstance(raw, Mapping):
        _warn_skip(field, "mapping", raw)
        return

    street = _nested_optional_text(raw, field, "street")
    city = _nested_optional_text(raw, field, "city")
    state = _nested_optional_text(raw, field, "state")
    postal_code = _nested_optional_text(raw, field, "postalCode")
    country = _nested_optional_text(raw, field, "countryOrRegion")

    if not any([street, city, state, postal_code, country]):
        return

    components = ["", "", street, city, state, postal_code, country]
    value = _join_structured([part or "" for part in components])
    lines.append(f"ADR;TYPE={adr_type}:{value}")


def _nested_optional_text(data: Mapping[str, object], parent: str, field: str) -> str | None:
    value = data.get(field)
    if value is None:
        return None
    if not isinstance(value, str):
        _warn_skip(f"{parent}.{field}", "str", value)
        return None
    text = value.strip()
    if not text:
        return None
    return text


def _birthday(raw: object) -> str | None:
    if raw is None:
        return None
    if not isinstance(raw, str):
        _warn_skip("birthday", "str", raw)
        return None
    try:
        parsed = dateutil_parser.isoparse(raw)
    except (TypeError, ValueError, OverflowError):
        log.warning("Skipping unparseable contact birthday")
        return None
    if parsed.year == 1604:
        return f"--{parsed.month:02d}{parsed.day:02d}"
    return f"{parsed.year:04d}-{parsed.month:02d}-{parsed.day:02d}"


def _normalise_rev(raw: object) -> str | None:
    if raw is None:
        return None
    if not isinstance(raw, str):
        _warn_skip("lastModifiedDateTime", "str", raw)
        return None
    try:
        parsed = dateutil_parser.isoparse(raw)
    except (TypeError, ValueError, OverflowError):
        log.warning("Skipping unparseable contact lastModifiedDateTime")
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=UTC)
    normalised = parsed.astimezone(UTC).replace(microsecond=0)
    return normalised.strftime("%Y-%m-%dT%H:%M:%SZ")


def _categories(raw: object) -> str | None:
    if raw is None:
        return None
    if not isinstance(raw, list):
        _warn_skip("categories", "list", raw)
        return None

    escaped: list[str] = []
    for index, item in enumerate(raw):
        if not isinstance(item, str):
            _warn_skip(f"categories[{index}]", "str", item)
            continue
        value = item.strip()
        if value:
            escaped.append(_escape_text(value))
    if not escaped:
        return None
    return ",".join(escaped)


def _full_name(
    display_name: str | None,
    given_name: str | None,
    surname: str | None,
    first_email: str | None,
) -> str:
    if display_name:
        return display_name
    fallback = " ".join(part for part in [given_name, surname] if part)
    if fallback:
        return fallback
    if first_email:
        return first_email
    return "(no name)"


def _escape_text(value: str) -> str:
    """Escape a TEXT value per RFC 2426 section 4 (backslash, semicolon, comma, newline)."""
    escaped = _CONTROL_CHARS.sub("", value)
    escaped = escaped.replace("\\", "\\\\")
    escaped = escaped.replace(";", "\\;").replace(",", "\\,")
    escaped = escaped.replace("\r\n", "\\n").replace("\r", "\\n").replace("\n", "\\n")
    return escaped


def _join_structured(parts: Iterable[str]) -> str:
    return ";".join(_escape_text(part) for part in parts)


def _serialize_vcard(lines: list[str]) -> bytes:
    physical_lines: list[str] = []
    for line in lines:
        physical_lines.extend(_fold_line(line))
    return ("\r\n".join(physical_lines) + "\r\n").encode("utf-8")


def _fold_line(line: str) -> list[str]:
    if len(line.encode("utf-8")) <= 75:
        return [line]

    folded: list[str] = []
    remainder = line
    first_line = True
    while remainder:
        max_octets = 75 if first_line else 74
        piece, remainder = _utf8_prefix(remainder, max_octets)
        if first_line:
            folded.append(piece)
            first_line = False
        else:
            folded.append(f" {piece}")
    return folded


def _utf8_prefix(text: str, max_octets: int) -> tuple[str, str]:
    used = 0
    index = 0
    for character in text:
        size = len(character.encode("utf-8"))
        if used + size > max_octets:
            break
        used += size
        index += 1
    if index == 0:
        return text[:1], text[1:]
    return text[:index], text[index:]


def _warn_skip(field: str, expected: str, value: object) -> None:
    log.warning(
        "Skipping malformed contact field %s: expected %s, got %s",
        field,
        expected,
        type(value).__name__,
    )
