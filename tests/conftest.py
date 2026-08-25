"""Shared pytest fixtures for outlook-desktop-mcp tests."""
import os
import sys

import pytest

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC_ROOT = os.path.join(REPO_ROOT, "src")

if SRC_ROOT not in sys.path:
    sys.path.insert(0, SRC_ROOT)


# Script-style legacy validation files: keep available for manual `python tests/<name>.py`
# runs but exclude from pytest's automatic collection (they predate pytest and contain
# ``def test_*`` helpers that aren't real pytest tests).
collect_ignore = [
    "phase1_com_test.py",
    "calendar_com_test.py",
    "extras_com_test.py",
    "phase3_mcp_test.py",
    "calendar_mcp_test.py",
    "extras_mcp_test.py",
]


def outlook_integration_enabled() -> bool:
    return os.environ.get("RUN_OUTLOOK_INTEGRATION", "").strip() in ("1", "true", "yes")


@pytest.fixture
def repo_root() -> str:
    return REPO_ROOT


@pytest.fixture
def stdio_server_params():
    """Build StdioServerParameters for the Windows MCP server module."""
    from mcp.client.stdio import StdioServerParameters

    def _factory(module: str = "outlook_desktop_mcp.server"):
        return StdioServerParameters(
            command=sys.executable,
            args=["-m", module],
            cwd=REPO_ROOT,
            env={**os.environ, "PYTHONPATH": SRC_ROOT},
        )

    return _factory


@pytest.fixture
def skip_without_outlook_integration():
    if not outlook_integration_enabled():
        pytest.skip("Set RUN_OUTLOOK_INTEGRATION=1 to run Outlook integration tests")
