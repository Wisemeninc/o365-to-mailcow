"""Graph contact -> vCard 3.0 (ISC-85..89, ISC-109, ISC-110).

Every test parses the output back with vobject (test-only dependency) so the assertions
are about what a CardDAV client will read, not about string layout, except where the
criterion is itself about layout (escaping, folding).
"""

from __future__ import annotations

import base64
import json
import uuid
from pathlib import Path
from typing import Any

import pytest
import vobject

from o365_to_mailcow.contacts_conv import ConvertedContact, contact_uid, convert_contact

FIXTURES = Path(__file__).parent / "fixtures" / "graph"


def load(name: str) -> dict[str, Any]:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def parse(converted: ConvertedContact) -> Any:
    return vobject.readOne(converted.vcf.decode("utf-8"))


def physical_lines(vcf: bytes) -> list[bytes]:
    assert vcf.endswith(b"\r\n")
    return vcf[:-2].split(b"\r\n")


def unfold(vcf: bytes) -> str:
    return vcf.decode("utf-8").replace("\r\n ", "")


def minimal(**fields: Any) -> dict[str, Any]:
    return {"id": "AAMkAD-test-contact", **fields}


# -- ISC-85: UID ---------------------------------------------------------------------------


def test_isc85_uid_is_uuid5_of_graph_id() -> None:
    contact = load("contact_full.json")
    converted = convert_contact(contact)
    expected = str(uuid.uuid5(uuid.NAMESPACE_URL, "o365-to-mailcow:contact:" + contact["id"]))
    assert converted.uid == expected == contact_uid(contact["id"])
    assert parse(converted).uid.value == expected


def test_isc85_uid_is_deterministic_and_distinct() -> None:
    assert contact_uid("a") == contact_uid("a")
    assert contact_uid("a") != contact_uid("b")
    assert convert_contact(minimal()).vcf == convert_contact(minimal()).vcf


def test_isc85_version_is_3_0_and_crlf() -> None:
    converted = convert_contact(load("contact_full.json"))
    lines = physical_lines(converted.vcf)
    assert lines[0] == b"BEGIN:VCARD"
    assert lines[1] == b"VERSION:3.0"
    assert lines[-1] == b"END:VCARD"
    assert b"\n" not in converted.vcf.replace(b"\r\n", b"")


def test_last_modified_is_passed_through_raw() -> None:
    contact = load("contact_full.json")
    assert convert_contact(contact).last_modified == contact["lastModifiedDateTime"]
    assert convert_contact(minimal()).last_modified is None


# -- ISC-86: property mapping --------------------------------------------------------------


def test_isc86_full_contact_every_property() -> None:
    card = parse(convert_contact(load("contact_full.json")))

    assert card.fn.value == "Ada Lovelace"
    n = card.n.value
    assert (n.family, n.given, n.additional, n.prefix, n.suffix) == (
        "Lovelace", "Ada", "Byron", "Countess", "III",
    )
    assert card.nickname.value == "Enchantress"
    assert card.org.value == ["Analytical Engines Inc", "Research"]
    assert card.title.value == "Mathematician"

    emails = [e.value for e in card.contents["email"]]
    assert emails == ["ADA@example.com", "ada.lovelace@example.org"]  # case-insensitive dedupe
    assert all(e.params["TYPE"] == ["INTERNET"] for e in card.contents["email"])

    tels = sorted((t.params["TYPE"][0], t.value) for t in card.contents["tel"])
    assert tels == [("CELL", "+1 555 0102"), ("HOME", "+1 555 0101"), ("WORK", "+1 555 0100")]

    adrs = {a.params["TYPE"][0]: a.value for a in card.contents["adr"]}
    assert set(adrs) == {"HOME", "WORK", "OTHER"}
    home = adrs["HOME"]
    assert (home.box, home.extended, home.street, home.city, home.region, home.code,
            home.country) == ("", "", "1 Home St", "London", "LN", "H0M3", "UK")
    assert adrs["WORK"].street == "2 Work Rd"
    assert adrs["OTHER"].city == "Oxford"

    assert card.bday.value == "1980-05-17"
    assert card.note.value == "Met at conference."
    assert card.url.value == "https://example.com/ada"
    assert card.categories.value == ["VIP", "Engineering"]
    assert card.rev.value == "2026-09-20T10:34:56Z"  # fractional seconds dropped


def test_isc86_empty_phone_and_empty_address_are_skipped() -> None:
    converted = convert_contact(
        minimal(
            displayName="X",
            businessPhones=["", "  "],
            homeAddress={"street": "", "city": None},
        )
    )
    card = parse(converted)
    assert "tel" not in card.contents
    assert "adr" not in card.contents


def test_isc86_org_with_department_only() -> None:
    card = parse(convert_contact(minimal(displayName="X", department="Ops")))
    assert card.org.value == ["", "Ops"]


def test_absent_optional_properties_are_not_emitted() -> None:
    text = convert_contact(minimal(displayName="Only Name")).vcf.decode()
    for prop in ("NICKNAME", "ORG", "TITLE", "EMAIL", "TEL", "ADR", "BDAY", "NOTE", "URL",
                 "CATEGORIES", "PHOTO", "REV"):
        assert f"\r\n{prop}" not in text, prop


# -- ISC-87: FN fallback -------------------------------------------------------------------


def test_isc87_fn_falls_back_to_first_email() -> None:
    card = parse(convert_contact(load("contact_email_only.json")))
    assert card.fn.value == "email.only@example.com"
    assert "n" in card.contents  # N is mandatory in vCard 3.0


def test_isc87_fn_falls_back_to_given_and_surname() -> None:
    card = parse(convert_contact(minimal(givenName="Grace", surname="Hopper")))
    assert card.fn.value == "Grace Hopper"


def test_isc87_fn_last_resort() -> None:
    card = parse(convert_contact(minimal(displayName="   ")))
    assert card.fn.value == "(no name)"


def test_isc87_fn_skips_legacy_dn_when_choosing_email() -> None:
    contact = minimal(emailAddresses=[{"address": "/o=ExchangeLabs/cn=x"}, {"address": "a@b.c"}])
    assert parse(convert_contact(contact)).fn.value == "a@b.c"


# -- ISC-88: photo -------------------------------------------------------------------------


def test_isc88_photo_embedded_base64_jpeg_and_folded() -> None:
    photo = bytes(range(256)) * 20  # 5120 bytes -> long base64 line that must fold
    converted = convert_contact(minimal(displayName="P"), photo=photo)
    raw = unfold(converted.vcf)
    assert "PHOTO;ENCODING=b;TYPE=JPEG:" + base64.b64encode(photo).decode() in raw
    card = parse(converted)
    assert card.photo.value == photo
    assert all(len(line) <= 75 for line in physical_lines(converted.vcf))


def test_isc88_empty_photo_is_ignored() -> None:
    assert b"PHOTO" not in convert_contact(minimal(displayName="P"), photo=b"").vcf


# -- ISC-89: escaping and folding ----------------------------------------------------------


def test_isc89_text_escaping_round_trips() -> None:
    nasty = "Doe, John; Jr.\\ \nsecond line\r\nthird"
    converted = convert_contact(
        minimal(displayName=nasty, personalNotes=nasty, companyName="A;B", department="C,D")
    )
    raw = unfold(converted.vcf)
    assert "FN:Doe\\, John\\; Jr.\\\\ \\nsecond line\\nthird\r\n" in raw
    assert "ORG:A\\;B;C\\,D\r\n" in raw
    card = parse(converted)
    normalised = nasty.replace("\r\n", "\n")
    assert card.fn.value == normalised
    assert card.note.value == normalised
    assert card.org.value == ["A;B", "C,D"]


def test_isc89_structured_components_escape_separators() -> None:
    converted = convert_contact(minimal(surname="O;Brien", givenName="Ann,Marie"))
    assert "N:O\\;Brien;Ann\\,Marie;;;\r\n" in unfold(converted.vcf)
    n = parse(converted).n.value
    assert (n.family, n.given) == ("O;Brien", "Ann,Marie")


def test_isc89_categories_escape_commas_inside_values() -> None:
    converted = convert_contact(minimal(displayName="C", categories=["a,b", "c"]))
    assert "CATEGORIES:a\\,b,c\r\n" in unfold(converted.vcf)


def test_isc89_long_lines_fold_at_75_octets_without_splitting_utf8() -> None:
    note = ("Grüße aus Köln, €uro; emoji 😀 " * 15).strip()
    converted = convert_contact(minimal(displayName="F", personalNotes=note))
    lines = physical_lines(converted.vcf)
    assert any(line.startswith(b" ") for line in lines), "expected folded continuation lines"
    for line in lines:
        assert len(line) <= 75
        line.decode("utf-8")  # a split multi-byte character would raise here
    # Removing exactly one CRLF+SPACE per fold restores the logical line.
    escaped = note.replace(",", "\\,").replace(";", "\\;")
    assert f"\r\nNOTE:{escaped}\r\n" in unfold(converted.vcf)
    assert parse(converted).note.value == note


def test_control_characters_are_removed() -> None:
    converted = convert_contact(minimal(displayName="Bell\x07Name", businessHomePage="http://x\n/y"))
    card = parse(converted)
    assert card.fn.value == "BellName"
    assert card.url.value == "http://x/y"


# -- ISC-109, ISC-110 ----------------------------------------------------------------------


def test_isc109_legacy_dn_addresses_dropped() -> None:
    converted = convert_contact(load("contact_legacy_dn.json"))
    card = parse(converted)
    assert [e.value for e in card.contents["email"]] == ["legacy.user@example.net"]
    assert b"/o=ExchangeLabs" not in converted.vcf


def test_isc110_birthday_year_1604_omits_year() -> None:
    converted = convert_contact(load("contact_legacy_dn.json"))
    assert b"\r\nBDAY:--0517\r\n" in converted.vcf
    assert parse(converted).bday.value == "--0517"


# -- robustness ------------------------------------------------------------------------------


def test_malformed_optional_fields_do_not_raise(caplog: pytest.LogCaptureFixture) -> None:
    contact = minimal(
        displayName="Robust",
        emailAddresses="x",
        businessPhones=[None, 5, "+49 1"],
        homePhones="nope",
        mobilePhone=12345,
        birthday="garbage",
        homeAddress="str",
        businessAddress={"street": 7, "city": "Berlin"},
        categories=[1, "ok"],
        lastModifiedDateTime=12,
        nickName=["list"],
    )
    converted = convert_contact(contact)
    card = parse(converted)
    assert card.fn.value == "Robust"
    assert [t.value for t in card.contents["tel"]] == ["+49 1"]
    assert card.contents["adr"][0].value.city == "Berlin"
    assert card.categories.value == ["ok"]
    assert "bday" not in card.contents
    assert converted.last_modified is None
    assert "Skipping" in caplog.text


def test_missing_id_raises() -> None:
    with pytest.raises(ValueError, match="id"):
        convert_contact({"displayName": "No Id"})
    with pytest.raises(ValueError):
        contact_uid("")


def test_non_dict_contact_raises_type_error() -> None:
    with pytest.raises(TypeError):
        convert_contact(["not", "a", "dict"])  # type: ignore[arg-type]
