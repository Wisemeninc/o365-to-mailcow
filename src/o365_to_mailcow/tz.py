"""Windows -> IANA time-zone mapping for Graph events.

Graph reports an event's original zone in ``originalStartTimeZone`` /
``originalEndTimeZone`` using Windows zone names ("W. Europe Standard Time"), sometimes
IANA names, sometimes the literal "UTC" or "tzone://Microsoft/Utc", and for zones the
organiser built by hand, "tzone://Microsoft/Custom" (which carries no usable rules).

The mapping table is CLDR's windowsZones data as shipped by ``tzlocal`` (already a
runtime dependency), so no zone list is maintained here. Unknown names return ``None``
so the caller can fall back to UTC and report it, rather than guessing a zone.
"""

from __future__ import annotations

import zoneinfo
from functools import lru_cache

from tzlocal.windows_tz import win_tz

UTC_NAME = "UTC"

# Names Graph uses for UTC that are not keys of the CLDR table (or map to Etc/UTC,
# which we normalise to plain "UTC" for readability in the generated ICS).
_UTC_ALIASES = frozenset(
    {
        "utc",
        "tzone://microsoft/utc",
        "coordinated universal time",
        "etc/utc",
        "gmt",
        "etc/gmt",
        "z",
    }
)


def windows_to_iana(name: str | None) -> str | None:
    """Return the IANA zone name for a Windows (or IANA) zone name, or None if unknown.

    Resolution order: UTC aliases -> "UTC"; CLDR Windows table; the name itself when it
    already is a loadable IANA key. Anything else (including "tzone://Microsoft/Custom",
    empty strings and None) is unknown.
    """
    if name is None:
        return None
    cleaned = name.strip()
    if not cleaned:
        return None
    if cleaned.lower() in _UTC_ALIASES:
        return UTC_NAME
    mapped = win_tz.get(cleaned)
    if mapped is not None:
        return UTC_NAME if mapped in ("Etc/UTC", "UTC") else mapped
    if _is_iana(cleaned):
        return cleaned
    return None


def zone(name: str) -> zoneinfo.ZoneInfo:
    """Load an IANA zone; raises ``zoneinfo.ZoneInfoNotFoundError`` for unknown names."""
    return _load_zone(name)


def _is_iana(name: str) -> bool:
    # Reject path-like or relative keys before touching the tz database: ZoneInfo raises
    # ValueError on those, and a Graph field should never be interpreted as a path.
    if name.startswith("/") or ".." in name.split("/"):
        return False
    try:
        _load_zone(name)
    except (zoneinfo.ZoneInfoNotFoundError, ValueError, IsADirectoryError):
        return False
    return True


@lru_cache(maxsize=256)
def _load_zone(name: str) -> zoneinfo.ZoneInfo:
    return zoneinfo.ZoneInfo(name)
