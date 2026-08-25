"""Unit tests for ContactCache (no Outlook required)."""
import time

import pytest

from outlook_desktop_mcp.utils.contact_cache import (
    CONTACT_CACHE_TTL_SECONDS,
    ContactCache,
)


def test_default_ttl_is_seven_days():
    assert CONTACT_CACHE_TTL_SECONDS == 7 * 24 * 3600


def test_set_and_get_within_ttl():
    cache = ContactCache(ttl_seconds=60)
    key = cache.make_key("search_contacts", query="alice", count=10, account="")
    cache.set(key, '[{"full_name": "Alice"}]')
    assert cache.get(key) == '[{"full_name": "Alice"}]'


def test_get_missing_returns_none():
    cache = ContactCache()
    assert cache.get("search_contacts:nonexistent") is None


def test_expired_entry_removed():
    cache = ContactCache(ttl_seconds=0.01)
    key = cache.make_key("list_contacts", count=5, account="")
    cache.set(key, "[]")
    time.sleep(0.02)
    assert cache.get(key) is None
    assert len(cache) == 0


def test_make_key_stable_for_same_params():
    cache = ContactCache()
    k1 = cache.make_key("search_contacts", query="bob", count=20, account="Work")
    k2 = cache.make_key("search_contacts", account="Work", count=20, query="bob")
    assert k1 == k2


def test_make_key_differs_by_store_id():
    cache = ContactCache()
    k1 = cache.make_key("list_contacts", store_id="store-a", count=10, account="")
    k2 = cache.make_key("list_contacts", store_id="store-b", count=10, account="")
    assert k1 != k2


def test_make_key_differs_for_different_params():
    cache = ContactCache()
    k1 = cache.make_key("search_contacts", query="bob", count=20)
    k2 = cache.make_key("search_contacts", query="carol", count=20)
    assert k1 != k2


def test_clear_removes_all():
    cache = ContactCache()
    key = cache.make_key("resolve_recipient", name="John", account="")
    cache.set(key, "{}")
    cache.clear()
    assert cache.get(key) is None
    assert len(cache) == 0


def test_lru_evicts_oldest_when_over_max_entries():
    cache = ContactCache(ttl_seconds=3600, max_entries=2)
    k1 = cache.make_key("a", x=1)
    k2 = cache.make_key("a", x=2)
    k3 = cache.make_key("a", x=3)
    cache.set(k1, "one")
    cache.set(k2, "two")
    cache.set(k3, "three")
    assert cache.get(k1) is None
    assert cache.get(k2) == "two"
    assert cache.get(k3) == "three"
