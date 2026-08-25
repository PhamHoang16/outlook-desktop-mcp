"""macOS contact record parsing for AppleScript bridge output."""
from __future__ import annotations

from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM


def clean_applescript_value(value: str) -> str:
    """Replace AppleScript's 'missing value' with empty string."""
    v = value.strip()
    return "" if v == "missing value" else v


def parse_mac_contact_records(
    raw: str,
    query: str = "",
    limit: int = 50,
    *,
    clean=clean_applescript_value,
) -> list[dict]:
    """Parse DELIM/RECORD_DELIM contact lines from AppleScript."""
    query_lower = query.lower().strip()
    results: list[dict] = []
    for record in raw.split(RECORD_DELIM):
        record = record.strip()
        if not record:
            continue
        parts = record.split(DELIM)
        if len(parts) < 3:
            continue
        entry_id = parts[0].strip()
        full_name = clean(parts[1]) or "(no name)"
        email = clean(parts[2])
        if query_lower:
            haystack = f"{full_name} {email}".lower()
            if query_lower not in haystack:
                continue
        results.append({
            "entry_id": entry_id,
            "full_name": full_name,
            "email": email,
            "company": clean(parts[3]) if len(parts) > 3 else "",
            "phone": clean(parts[4]) if len(parts) > 4 else "",
            "job_title": clean(parts[5]) if len(parts) > 5 else "",
        })
        if len(results) >= limit:
            break
    return results
