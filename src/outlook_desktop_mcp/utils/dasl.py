"""Locale-independent helpers for Outlook DASL filters."""
from __future__ import annotations

from datetime import datetime

# Sentinel date used by OLE Automation: 30 December 1899 (NOT 31 December).
_OA_EPOCH = datetime(1899, 12, 30)


def to_oa_date(dt: datetime) -> float:
    """Convert a naive local-time ``datetime`` to an OLE Automation date number.

    DASL date string parsing in Outlook (``Items.Restrict``) depends on the
    user's Windows regional settings (``%m/%d/%Y`` vs ``%d/%m/%Y``). Passing
    the numeric OLE Automation date instead is fully locale-independent.

    The caller should pass a naive datetime in local time (Outlook's expected
    time zone for restrict filters). Timezone-aware datetimes are converted
    to the local timezone first.
    """
    if dt.tzinfo is not None:
        dt = dt.astimezone().replace(tzinfo=None)
    delta = dt - _OA_EPOCH
    return delta.total_seconds() / 86400.0


def dasl_date_literal(dt: datetime) -> str:
    """Return a DASL literal (no quotes) for ``dt`` using OA date numbers."""
    return f"{to_oa_date(dt):.6f}"
