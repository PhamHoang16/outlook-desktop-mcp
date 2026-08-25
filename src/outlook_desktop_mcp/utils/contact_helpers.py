"""Shared helpers for contact MCP tools (Windows and macOS)."""
from __future__ import annotations

import json
from typing import Any

CONTACT_COUNT_MIN = 1
CONTACT_COUNT_MAX = 200
CONTACT_COUNT_DEFAULT_LIST = 50
CONTACT_COUNT_DEFAULT_SEARCH = 20


def clamp_contact_count(count: int) -> int:
    """Clamp list/search result limit to the supported range."""
    return min(max(CONTACT_COUNT_MIN, count), CONTACT_COUNT_MAX)


def normalize_account_key(account: str) -> str:
    """Normalize account substring for cache keys (case-insensitive)."""
    return account.lower().strip()


def normalize_search_query(query: str) -> str | None:
    """Return stripped query, or None if empty after strip."""
    stripped = query.strip()
    return stripped if stripped else None


def should_cache_contact_result(
    result: str,
    *,
    tool: str,
    allow_empty_list: bool = True,
) -> bool:
    """Decide whether a successful contact tool JSON response may be cached."""
    if not result or result.startswith("Error"):
        return False
    try:
        data = json.loads(result)
    except json.JSONDecodeError:
        return False

    if "resolve" in tool:
        return isinstance(data, dict) and data.get("resolved") is True

    if isinstance(data, list):
        if not allow_empty_list and len(data) == 0:
            return False
        return True

    return isinstance(data, dict)


def contact_matches_query(contact: dict, query: str) -> bool:
    """True if query is a case-insensitive substring of name/email/company."""
    needle = query.casefold()
    if not needle:
        return False
    for key in ("full_name", "email", "company"):
        if needle in (contact.get(key) or "").casefold():
            return True
    return False


def pick_best_contact_match(contacts: list[dict], name: str) -> dict | None:
    """Prefer exact then partial full_name match; else first result."""
    if not contacts:
        return None
    needle = name.casefold()
    for contact in contacts:
        if (contact.get("full_name") or "").casefold() == needle:
            return contact
    for contact in contacts:
        if needle in (contact.get("full_name") or "").casefold():
            return contact
    return contacts[0]


def sort_contacts_by_name(contacts: list[dict]) -> list[dict]:
    """Sort contact dicts by full_name (case-insensitive)."""
    return sorted(
        contacts,
        key=lambda c: (c.get("full_name") or "").lower(),
    )
