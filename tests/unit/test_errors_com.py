"""Unit tests for enhanced COM error formatting."""
from outlook_desktop_mcp.utils.errors import format_com_error


class _FakeComError(Exception):
    def __init__(self, hresult: int, msg: str):
        super().__init__(hresult, msg, None, None)
        self.hresult = hresult


def test_format_com_error_mapi_not_found_hint():
    err = _FakeComError(-2147221233, "not found")  # 0x8004010F
    text = format_com_error(err)
    assert "0x8004010F" in text
    assert "not found" in text.lower() or "EntryID" in text


def test_format_com_error_rpc_disconnected_hint():
    err = _FakeComError(-2147417848, "disconnected")  # 0x80010108
    text = format_com_error(err)
    assert "0x80010108" in text
    assert "restart Outlook" in text


def test_format_non_com_includes_type():
    text = format_com_error(ValueError("bad account"))
    assert "ValueError" in text
