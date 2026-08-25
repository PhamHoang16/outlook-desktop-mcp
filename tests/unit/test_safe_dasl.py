"""Unit tests for _safe_dasl input sanitization."""
import sys
import types


def _load_safe_dasl():
    """Import only the _safe_dasl function from server.py without importing
    pywin32 (which isn't installed in CI / on Linux).
    """
    import importlib.util
    import pathlib

    path = pathlib.Path(__file__).resolve().parents[2] / "src" / "outlook_desktop_mcp" / "server.py"
    source = path.read_text(encoding="utf-8")

    src_func = []
    capture = False
    for line in source.splitlines():
        if line.startswith("def _safe_dasl"):
            capture = True
        if capture:
            src_func.append(line)
            if capture and src_func and src_func[-1].startswith("    return "):
                break
    ns: dict = {}
    exec("\n".join(src_func), ns)
    return ns["_safe_dasl"]


_safe_dasl = _load_safe_dasl()


def test_wildcards_escaped():
    out = _safe_dasl("100% off")
    assert "[%]" in out
    assert "%" not in out.replace("[%]", "")


def test_underscore_escaped():
    out = _safe_dasl("file_name")
    assert "[_]" in out


def test_bracket_escaped():
    out = _safe_dasl("[urgent]")
    assert "[[]" in out


def test_quotes_doubled():
    assert "''" in _safe_dasl("it's")
    assert '""' in _safe_dasl('"quoted"')


def test_empty_string_round_trip():
    assert _safe_dasl("") == ""
