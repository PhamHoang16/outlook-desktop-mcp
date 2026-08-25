"""Unit tests for OutlookBridge error classification and disconnect detection.

These tests do not start the COM thread (which would need pywin32 and a
running Outlook). They exercise the pure-Python helpers that classify
exceptions for retry / disconnect handling.
"""
import pytest

from outlook_desktop_mcp.com_bridge import (
    ComBridgeBusyError,
    ComBridgeDisconnectedError,
    ComBridgeTimeoutError,
    _is_rpc_disconnected,
)


class _FakeHResultError(Exception):
    def __init__(self, hresult: int):
        super().__init__("fake com error")
        self.hresult = hresult


@pytest.mark.parametrize(
    "hresult,expected",
    [
        (0x80010108, True),   # RPC_E_DISCONNECTED
        (-2147023174, True),  # 0x800706BA RPC server unavailable (signed)
        (0x80020009, False),  # DISP_E_EXCEPTION (different category)
        (0, False),
    ],
)
def test_is_rpc_disconnected(hresult, expected):
    assert _is_rpc_disconnected(_FakeHResultError(hresult)) is expected


def test_is_rpc_disconnected_args_fallback():
    class _LegacyError(Exception):
        pass

    err = _LegacyError(0x800706BA, "rpc unavailable", None, None)
    assert _is_rpc_disconnected(err) is True


def test_is_rpc_disconnected_no_hresult():
    assert _is_rpc_disconnected(ValueError("not com")) is False


def test_bridge_error_codes_and_retriable():
    assert ComBridgeBusyError("busy").code == "com_busy"
    assert ComBridgeBusyError("busy").retriable is True
    assert ComBridgeTimeoutError("slow").code == "com_timeout"
    assert ComBridgeTimeoutError("slow").retriable is False
    assert ComBridgeDisconnectedError("gone").code == "com_disconnected"
    assert ComBridgeDisconnectedError("gone").retriable is True
