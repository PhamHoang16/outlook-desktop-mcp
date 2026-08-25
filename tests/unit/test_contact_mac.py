"""Unit tests for macOS contact record parsing."""
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM
from outlook_desktop_mcp.utils.contact_mac import parse_mac_contact_records


def _record(entry_id, name, email, company="", phone="", title=""):
    return (
        f"{entry_id}{DELIM}{name}{DELIM}{email}{DELIM}"
        f"{company}{DELIM}{phone}{DELIM}{title}{RECORD_DELIM}"
    )


def test_parse_filters_by_query():
    raw = _record("1", "Alice Smith", "alice@example.com") + _record(
        "2", "Bob Jones", "bob@example.com"
    )
    results = parse_mac_contact_records(raw, query="alice", limit=10)
    assert len(results) == 1
    assert results[0]["email"] == "alice@example.com"


def test_parse_skips_short_records():
    raw = f"broken{RECORD_DELIM}" + _record("1", "Jane", "jane@example.com")
    results = parse_mac_contact_records(raw, limit=10)
    assert len(results) == 1
