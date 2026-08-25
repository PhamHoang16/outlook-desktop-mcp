"""COM error formatting and structured error JSON helpers."""
import json
import logging

_logger = logging.getLogger("outlook_desktop_mcp.errors")


def _com_hresult_and_message(e: Exception) -> tuple[int, str] | None:
    try:
        import pythoncom

        if isinstance(e, pythoncom.com_error):
            hr, msg, *_ = e.args
            return int(hr), str(msg)
    except Exception:
        pass
    try:
        import pywintypes

        if isinstance(e, pywintypes.com_error):
            hr, msg, *_ = e.args
            return int(hr), str(msg)
    except Exception:
        pass
    hresult = getattr(e, "hresult", None)
    if hresult is not None:
        return int(hresult), str(e)
    return None


def _com_error_hint(hresult: int) -> str:
    code = hresult & 0xFFFFFFFF
    if code == 0x80020009:
        return (
            "Outlook rejected the COM call (often DASL Restrict or address-book "
            "access). Allow programmatic access in Trust Center, or sync the "
            "offline address book."
        )
    if code in (0x80070005, 0x80004005):
        return (
            "Access denied — approve Outlook's programmatic access prompt for "
            "address information."
        )
    if code == 0x8004010F:
        return "Item not found — EntryID may be stale or the item was moved."
    if code == 0x80020005:
        return "Property type mismatch — check argument types for this item."
    if code == 0x80010108:
        return "Outlook RPC disconnected — restart Outlook and retry."
    return ""


def format_com_error(e: Exception) -> str:
    parsed = _com_hresult_and_message(e)
    if parsed is not None:
        hr, msg = parsed
        hint = _com_error_hint(hr)
        if hint:
            return f"COM Error (0x{hr & 0xFFFFFFFF:08X}): {msg}. {hint}"
        return f"COM Error (0x{hr & 0xFFFFFFFF:08X}): {msg}"
    _logger.warning("Unexpected non-COM exception: %s: %s", type(e).__name__, e)
    return f"An unexpected error occurred ({type(e).__name__}: {e})."


def tool_error_json(
    message: str,
    *,
    code: str = "error",
    retriable: bool = False,
    **extra,
) -> str:
    """Return a JSON-serialized error envelope for MCP tool responses.

    Use this for structured failures (timeouts, COM busy, validation errors)
    so LLM clients can branch on ``code`` / ``retriable`` instead of parsing
    free-form strings.
    """
    payload: dict[str, object] = {"error": message, "code": code, "retriable": retriable}
    payload.update(extra)
    return json.dumps(payload, ensure_ascii=False)


def format_bridge_exception(exc: Exception, *, action: str) -> str:
    """Return a structured JSON error for known bridge exceptions.

    Recognises ``ComBridgeTimeoutError``, ``ComBridgeBusyError``, and
    ``ComBridgeDisconnectedError`` by their ``code``/``retriable`` attributes.
    Falls back to ``format_com_error`` for unrecognised exceptions, preserving
    the legacy ``"Error <action>: <message>"`` string contract.

    ``action`` describes what the caller was trying to do
    (e.g. ``"sending email"``).
    """
    code = getattr(exc, "code", None)
    retriable = bool(getattr(exc, "retriable", False))
    if isinstance(code, str):
        extras: dict[str, object] = {}
        if code == "com_timeout":
            extras["warning"] = (
                "Outlook may have completed the operation after the timeout. "
                "Verify in the Outlook UI before retrying — do not assume the "
                "operation failed."
            )
        return tool_error_json(
            f"Error {action}: {exc}",
            code=code,
            retriable=retriable,
            **extras,
        )
    return f"Error {action}: {format_com_error(exc)}"
