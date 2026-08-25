"""Unit tests for AppleScript helpers."""
from datetime import datetime

import pytest

from outlook_desktop_mcp.utils.applescript_helpers import (
    DELIM,
    RECORD_DELIM,
    InvalidEntryIdError,
    escape,
    format_date,
    resolve_folder_ref,
    split_delimited_record,
    validate_mac_entry_id,
)


def test_validate_mac_entry_id_accepts_digits():
    assert validate_mac_entry_id("42") == "42"
    assert validate_mac_entry_id(42) == "42"
    assert validate_mac_entry_id("  123  ") == "123"


@pytest.mark.parametrize(
    "bad",
    [
        "1; do shell script \"rm -rf /\"",
        '1"\nset x to "evil',
        "abc",
        "",
        "   ",
        "1.0",
        "-1",
        "1" * 33,
        "1\n2",
    ],
)
def test_validate_mac_entry_id_rejects_injection(bad):
    with pytest.raises(InvalidEntryIdError):
        validate_mac_entry_id(bad)


def test_escape_handles_quotes_and_newlines():
    out = escape('hello "world"\nline2\\backslash')
    assert '\\"world\\"' in out
    assert "\\n" in out
    assert "\\\\backslash" in out


def test_escape_strips_unicode_line_separators_and_nul():
    out = escape("a\u2028b\u2029c\x00d")
    assert "\u2028" not in out
    assert "\u2029" not in out
    assert "\x00" not in out
    assert "a" in out and "d" in out


def test_delim_uses_ascii_control_chars():
    assert DELIM == "\x1f"
    assert RECORD_DELIM == "\x1e"


def test_format_date_emits_helper_call():
    expr = format_date(datetime(2026, 3, 22, 14, 5, 9))
    assert expr == "(my _odm_make_date(2026, 3, 22, 14, 5, 9))"


def test_resolve_folder_ref_built_in():
    assert resolve_folder_ref("Inbox") == "inbox"
    assert resolve_folder_ref(" sent ") == "sent items"


def test_resolve_folder_ref_custom_is_escaped():
    out = resolve_folder_ref('Quarterly "Reports"')
    assert out.startswith('mail folder "')
    assert '\\"Reports\\"' in out


def test_validate_mac_entry_id_nfkc_normalizes_fullwidth_digits():
    assert validate_mac_entry_id("４２") == "42"


def test_split_delimited_record_tail_field_may_contain_delim():
    record = f"a{DELIM}b{DELIM}c{DELIM}tail{DELIM}still-tail"
    parts = split_delimited_record(record, DELIM, 4)
    assert parts == ["a", "b", "c", f"tail{DELIM}still-tail"]


def test_split_delimited_record_wrong_count_returns_none():
    assert split_delimited_record(f"a{DELIM}b", DELIM, 3) is None
