"""Unit tests for structured error envelopes."""
import json

from outlook_desktop_mcp.com_bridge import (
    ComBridgeBusyError,
    ComBridgeDisconnectedError,
    ComBridgeTimeoutError,
)
from outlook_desktop_mcp.utils.errors import format_bridge_exception, tool_error_json


def test_tool_error_json_basic():
    text = tool_error_json("nope", code="bad_input", retriable=False, field="to")
    data = json.loads(text)
    assert data == {
        "error": "nope",
        "code": "bad_input",
        "retriable": False,
        "field": "to",
    }


def test_format_timeout_includes_warning_and_action():
    text = format_bridge_exception(
        ComBridgeTimeoutError("Outlook COM operation '_send' timed out after 60s."),
        action="sending email",
    )
    data = json.loads(text)
    assert data["code"] == "com_timeout"
    assert data["retriable"] is False
    assert "sending email" in data["error"]
    assert "Verify in the Outlook UI" in data["warning"]


def test_format_busy_is_retriable():
    text = format_bridge_exception(
        ComBridgeBusyError("busy"), action="listing emails"
    )
    data = json.loads(text)
    assert data["code"] == "com_busy"
    assert data["retriable"] is True
    assert "warning" not in data


def test_format_disconnected_is_retriable():
    text = format_bridge_exception(
        ComBridgeDisconnectedError("rpc dead"), action="reading email"
    )
    data = json.loads(text)
    assert data["code"] == "com_disconnected"
    assert data["retriable"] is True


def test_format_unknown_falls_back_to_string():
    text = format_bridge_exception(ValueError("bad account"), action="sending email")
    # Falls through to format_com_error path; legacy "Error <action>: ..." string.
    assert text.startswith("Error sending email:")
