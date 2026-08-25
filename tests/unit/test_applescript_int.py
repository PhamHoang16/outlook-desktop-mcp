"""Unit tests for coerce_script_int (AppleScript injection defense)."""
import pytest

from outlook_desktop_mcp.utils.applescript_helpers import (
    InvalidScriptIntError,
    coerce_script_int,
)


@pytest.mark.parametrize(
    "value,expected",
    [
        (10, 10),
        ("10", 10),
        ("  42  ", 42),
        (None, 5),
        ("", 5),
        (10.0, 10),
    ],
)
def test_accepts_valid_ints(value, expected):
    assert coerce_script_int(value, default=5, lo=1, hi=100) == expected


def test_clamps_low_and_high():
    assert coerce_script_int(-5, default=10, lo=1, hi=100) == 1
    assert coerce_script_int(99999, default=10, lo=1, hi=100) == 100


@pytest.mark.parametrize(
    "value",
    [
        "10\nend tell\ndo shell script \"id\"",
        "abc",
        "10; do shell script",
        object(),
        [1, 2],
    ],
)
def test_rejects_non_numeric_strings_and_types(value):
    with pytest.raises(InvalidScriptIntError):
        coerce_script_int(value, default=5)


def test_rejects_bool_explicitly():
    with pytest.raises(InvalidScriptIntError, match="boolean"):
        coerce_script_int(True, default=5)


def test_rejects_non_whole_floats():
    with pytest.raises(InvalidScriptIntError, match="whole"):
        coerce_script_int(3.5, default=5)


def test_includes_name_in_error():
    with pytest.raises(InvalidScriptIntError, match="count"):
        coerce_script_int("xx", default=5, name="count")
