"""ISC-61, ISC-62: Windows time-zone names map to IANA; unknown names return None."""

from __future__ import annotations

import zoneinfo

import pytest

from o365_to_mailcow.tz import windows_to_iana, zone


@pytest.mark.parametrize(
    ("windows", "iana"),
    [
        ("W. Europe Standard Time", "Europe/Berlin"),
        ("Romance Standard Time", "Europe/Paris"),
        ("GMT Standard Time", "Europe/London"),
        ("Pacific Standard Time", "America/Los_Angeles"),
        ("UTC", "UTC"),
    ],
)
def test_isc61_windows_names_map_to_iana(windows: str, iana: str) -> None:
    assert windows_to_iana(windows) == iana


@pytest.mark.parametrize(
    "alias", ["UTC", "tzone://Microsoft/Utc", "Coordinated Universal Time", "Etc/UTC", " utc "]
)
def test_utc_aliases_normalise_to_utc(alias: str) -> None:
    assert windows_to_iana(alias) == "UTC"


@pytest.mark.parametrize("iana", ["Europe/Berlin", "America/New_York", "Asia/Kolkata"])
def test_iana_names_pass_through(iana: str) -> None:
    assert windows_to_iana(iana) == iana


@pytest.mark.parametrize(
    "unknown",
    [None, "", "   ", "tzone://Microsoft/Custom", "Mars Standard Time", "../../etc/passwd",
     "/etc/localtime", "Europe/../Europe/Berlin"],
)
def test_isc62_unknown_names_return_none(unknown: str | None) -> None:
    assert windows_to_iana(unknown) is None


def test_zone_loads_iana() -> None:
    assert zone("Europe/Berlin") == zoneinfo.ZoneInfo("Europe/Berlin")


def test_zone_rejects_unknown() -> None:
    with pytest.raises(zoneinfo.ZoneInfoNotFoundError):
        zone("Not/AZone")
