"""Unit tests for DASL date helpers."""
from datetime import datetime, timezone

import pytest

from outlook_desktop_mcp.utils.dasl import dasl_date_literal, to_oa_date


def test_oa_epoch_is_zero():
    assert to_oa_date(datetime(1899, 12, 30)) == 0.0


def test_known_oa_date():
    # 1 Jan 2026 should be > 45000 (well-known landmark)
    val = to_oa_date(datetime(2026, 1, 1, 0, 0, 0))
    assert 45000.0 < val < 50000.0


def test_dasl_literal_is_unquoted_number():
    s = dasl_date_literal(datetime(2026, 3, 22, 14, 0, 0))
    assert s.replace(".", "").lstrip("-").isdigit()
    assert "'" not in s
    assert "/" not in s


def test_tz_aware_datetime_converted_to_local():
    dt_utc = datetime(2026, 3, 22, 14, 0, 0, tzinfo=timezone.utc)
    val = to_oa_date(dt_utc)
    assert isinstance(val, float)


@pytest.mark.parametrize(
    "dt",
    [
        datetime(2026, 1, 1),
        datetime(2026, 3, 22, 14, 30, 0),
        datetime(2030, 12, 31, 23, 59, 59),
    ],
)
def test_round_trip_monotonic(dt):
    a = to_oa_date(dt)
    b = to_oa_date(dt.replace(year=dt.year + 1))
    assert b > a
