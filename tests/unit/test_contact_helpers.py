"""Unit tests for contact_helpers."""
import json

import pytest

from outlook_desktop_mcp.utils.contact_helpers import (
    clamp_contact_count,
    contact_matches_query,
    normalize_account_key,
    normalize_search_query,
    pick_best_contact_match,
    should_cache_contact_result,
    sort_contacts_by_name,
)


def test_clamp_contact_count():
    assert clamp_contact_count(0) == 1
    assert clamp_contact_count(500) == 200
    assert clamp_contact_count(10) == 10


def test_normalize_account_key():
    assert normalize_account_key("  Work  ") == "work"
    assert normalize_account_key("") == ""


def test_normalize_search_query():
    assert normalize_search_query("  bob  ") == "bob"
    assert normalize_search_query("   ") is None
    assert normalize_search_query("") is None


def test_should_cache_resolve_only_success():
    ok = json.dumps({"resolved": True, "email": "a@b.com"})
    fail = json.dumps({"resolved": False, "name": "x"})
    assert should_cache_contact_result(ok, tool="resolve_recipient")
    assert not should_cache_contact_result(fail, tool="resolve_recipient")


def test_should_cache_empty_list_when_disabled():
    empty = json.dumps([])
    assert should_cache_contact_result(empty, tool="list_contacts")
    assert not should_cache_contact_result(
        empty, tool="mac:list_contacts", allow_empty_list=False
    )


def test_contact_matches_query_chinese():
    contact = {"full_name": "周尔康", "email": "zhou@example.com", "company": "BOCM"}
    assert contact_matches_query(contact, "周尔康")
    assert not contact_matches_query(contact, "张三")


def test_pick_best_contact_match_prefers_exact():
    contacts = [
        {"full_name": "周尔康 Jr", "email": "j@x.com"},
        {"full_name": "周尔康", "email": "z@x.com"},
    ]
    best = pick_best_contact_match(contacts, "周尔康")
    assert best["email"] == "z@x.com"


def test_sort_contacts_by_name():
    data = [
        {"full_name": "Zoe"},
        {"full_name": "Alice"},
    ]
    assert sort_contacts_by_name(data)[0]["full_name"] == "Alice"
