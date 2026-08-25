"""Outlook COM-compatibility helpers.

Centralises corrections informed by static analysis against the official
Outlook VBA / Object Model docs:

* Property assignments expecting an OLE ``DATE`` should receive a Python
  ``datetime`` (pywin32 marshals to OA date); raw locale-formatted strings
  are interpreted by Outlook using the current Windows regional settings
  and can silently misread day/month — same root cause as DASL date
  literals we already normalised in :mod:`utils.dasl`.

* ``AppointmentItem.Move`` / ``MailItem.Move`` / ``TaskItem.Move`` operate
  on a persisted item; freshly ``CreateItem``'d objects must be ``Save``'d
  first or the move can fail / lose properties.

* ``TaskItem`` exposes ``ReminderTime`` (absolute datetime); the
  ``ReminderMinutesBeforeStart`` property only exists on
  ``AppointmentItem``.

* ``AppointmentItem.Respond`` shows a UI dialog unless ``fNoUI=True``;
  unattended automation must pass the flag and then ``Send`` the returned
  response item explicitly.

These helpers contain no COM imports so they unit-test cleanly on Linux.
"""
from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any


def coerce_outlook_datetime(value: Any) -> datetime | None:
    """Coerce a user-supplied value to a naive ``datetime`` for COM assignment.

    Outlook's Object Model accepts strings for ``Start``/``End``/``DueDate``,
    but parses them through the user's Windows locale (``%m/%d/%Y`` vs
    ``%d/%m/%Y``), which has caused real bugs in non-en-US environments.
    Passing a Python ``datetime`` lets pywin32 marshal an unambiguous OLE
    automation date value.

    Returns ``None`` for empty input so callers can detect "leave unchanged".
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            return value.astimezone().replace(tzinfo=None)
        return value
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        try:
            dt = datetime.fromisoformat(stripped)
        except ValueError as e:
            raise ValueError(
                f"Date string must be ISO 8601 (e.g. '2026-02-25T14:00'), got {value!r}"
            ) from e
        if dt.tzinfo is not None:
            return dt.astimezone().replace(tzinfo=None)
        return dt
    raise TypeError(
        f"Date value must be a datetime or ISO 8601 string, got {type(value).__name__}"
    )


def compute_task_reminder_time(
    due_date: datetime | None, reminder_minutes: int
) -> datetime | None:
    """Return the absolute ``ReminderTime`` for a ``TaskItem``.

    ``TaskItem`` has no ``ReminderMinutesBeforeStart`` — only
    ``ReminderTime`` (an absolute datetime). When the caller supplies
    "remind me N minutes before due", subtract from the due datetime.

    Returns ``None`` when no reminder should be set (either ``reminder_minutes``
    <= 0 or there's no due date to anchor the reminder to).
    """
    if reminder_minutes is None or reminder_minutes <= 0:
        return None
    if due_date is None:
        # Without a due date, fall back to "remind N minutes from now" so the
        # user still gets a reminder; otherwise the request is silently dropped.
        return datetime.now() + timedelta(minutes=reminder_minutes)
    return due_date - timedelta(minutes=reminder_minutes)


def save_then_move(item, destination_folder):
    """Persist ``item`` then move it to ``destination_folder``.

    Required for items obtained from ``Outlook.Application.CreateItem`` —
    those exist only in memory and ``Move`` on an unpersisted item can fail
    or silently lose changes depending on Outlook version. This helper:

    1. Calls ``Save()`` so the item lives in the default folder for its
       :class:`OlItemType` and gets a real ``EntryID``.
    2. Calls ``Move(destination_folder)`` and returns the moved item
       (the COM return value), which has a *new* ``EntryID``.

    No-op if ``destination_folder`` is None — returns the saved item.
    """
    item.Save()
    if destination_folder is None:
        return item
    moved = item.Move(destination_folder)
    return moved if moved is not None else item
