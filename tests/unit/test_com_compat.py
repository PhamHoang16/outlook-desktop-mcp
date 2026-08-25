"""Static-audit fixes for Outlook COM compatibility helpers.

Mirrors the Outlook VBA docs (e.g. TaskItem.ReminderTime,
AppointmentItem.Move) cited in the audit; these are pure-Python helpers
so they run on every CI matrix entry (Linux/macOS/Windows).
"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone
from unittest.mock import MagicMock

import pytest

from outlook_desktop_mcp.utils.com_compat import (
    coerce_outlook_datetime,
    compute_task_reminder_time,
    save_then_move,
)


def test_coerce_outlook_datetime_passthrough_naive():
    dt = datetime(2026, 5, 20, 14, 30)
    assert coerce_outlook_datetime(dt) is dt


def test_coerce_outlook_datetime_iso_with_space():
    out = coerce_outlook_datetime("2026-05-20 14:30")
    assert out == datetime(2026, 5, 20, 14, 30)


def test_coerce_outlook_datetime_iso_t_separator():
    out = coerce_outlook_datetime("2026-05-20T14:30:00")
    assert out == datetime(2026, 5, 20, 14, 30, 0)


def test_coerce_outlook_datetime_empty_returns_none():
    assert coerce_outlook_datetime("") is None
    assert coerce_outlook_datetime("   ") is None
    assert coerce_outlook_datetime(None) is None


def test_coerce_outlook_datetime_strips_tzinfo():
    dt_aware = datetime(2026, 5, 20, 14, 30, tzinfo=timezone.utc)
    result = coerce_outlook_datetime(dt_aware)
    assert result.tzinfo is None


def test_coerce_outlook_datetime_rejects_bad_string():
    with pytest.raises(ValueError, match="ISO 8601"):
        coerce_outlook_datetime("05/20/2026 2:30 PM")


def test_coerce_outlook_datetime_rejects_bad_type():
    with pytest.raises(TypeError):
        coerce_outlook_datetime(12345)


def test_compute_task_reminder_time_no_reminder():
    due = datetime(2026, 5, 20, 9, 0)
    assert compute_task_reminder_time(due, 0) is None
    assert compute_task_reminder_time(due, -5) is None
    assert compute_task_reminder_time(due, None) is None


def test_compute_task_reminder_time_subtracts_from_due():
    due = datetime(2026, 5, 20, 9, 0)
    out = compute_task_reminder_time(due, 30)
    assert out == due - timedelta(minutes=30)


def test_compute_task_reminder_time_no_due_falls_back_to_now():
    out = compute_task_reminder_time(None, 15)
    assert out is not None
    delta = out - datetime.now()
    assert timedelta(minutes=14) <= delta <= timedelta(minutes=16)


def test_save_then_move_saves_then_moves():
    item = MagicMock()
    item.Save = MagicMock()
    moved = MagicMock()
    item.Move = MagicMock(return_value=moved)
    folder = MagicMock()

    result = save_then_move(item, folder)

    item.Save.assert_called_once_with()
    item.Move.assert_called_once_with(folder)
    assert result is moved


def test_save_then_move_returns_original_when_folder_none():
    item = MagicMock()
    item.Save = MagicMock()
    item.Move = MagicMock()

    result = save_then_move(item, None)

    item.Save.assert_called_once_with()
    item.Move.assert_not_called()
    assert result is item


def test_save_then_move_returns_original_when_move_returns_none():
    item = MagicMock()
    item.Save = MagicMock()
    item.Move = MagicMock(return_value=None)
    folder = MagicMock()

    result = save_then_move(item, folder)

    assert result is item


def test_save_called_before_move():
    """Regression: Move() on unpersisted CreateItem object can fail."""
    item = MagicMock()
    calls: list[str] = []
    item.Save = MagicMock(side_effect=lambda: calls.append("save"))
    item.Move = MagicMock(side_effect=lambda f: calls.append("move") or MagicMock())
    folder = MagicMock()

    save_then_move(item, folder)

    assert calls == ["save", "move"]
