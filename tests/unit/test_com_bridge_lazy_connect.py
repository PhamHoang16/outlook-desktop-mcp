"""Lifecycle tests for OutlookBridge lazy connection.

Regression cover for two field failures:

* Corporate security policy blocks ``Namespace.CurrentUser`` (Exchange address
  book) with E_ABORT. The bridge used to read it during startup, so the MCP
  server refused to start at all on locked-down machines.
* The bridge used to Dispatch Outlook during ``start()``, so the server could
  not come up before Outlook did.

These tests fake ``pythoncom`` / ``win32com.client`` so they run anywhere.
"""
import asyncio
import sys
import threading
import types

import pytest

from outlook_desktop_mcp.com_bridge import ComBridgeDisconnectedError, OutlookBridge


class _FakeCOMError(Exception):
    """Stands in for pythoncom.com_error (E_ABORT / 'Operation aborted')."""


def _install_fake_pythoncom(monkeypatch):
    mod = types.ModuleType("pythoncom")
    mod.COINIT_APARTMENTTHREADED = 0x2
    mod.com_error = _FakeCOMError
    mod.CoInitializeEx = lambda flags: None
    mod.CoInitialize = lambda: None
    mod.CoUninitialize = lambda: None
    monkeypatch.setitem(sys.modules, "pythoncom", mod)
    return mod


def _install_fake_win32com(monkeypatch, dispatch):
    """Install win32com.client whose Dispatch delegates to ``dispatch``."""
    win32com = types.ModuleType("win32com")
    client = types.ModuleType("win32com.client")
    client.Dispatch = dispatch
    win32com.client = client
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    return client


class _HostileNamespace:
    """MAPI namespace on a machine where corporate policy blocks the address book."""

    def __init__(self):
        self.inbox_reads = 0

    @property
    def CurrentUser(self):
        raise _FakeCOMError(-2147467260, "Operation aborted", None, None)

    @property
    def DefaultStore(self):
        raise _FakeCOMError(-2147467260, "Operation aborted", None, None)

    def GetDefaultFolder(self, _enum):
        self.inbox_reads += 1
        return "inbox-folder"


class _FakeOutlook:
    def __init__(self, namespace):
        self._namespace = namespace

    def GetNamespace(self, _kind):
        return self._namespace


@pytest.fixture
def bridge():
    b = OutlookBridge()
    yield b
    b.stop()


def _run(coro):
    return asyncio.run(coro)


def test_start_does_not_dispatch_outlook(monkeypatch, bridge):
    """start() must bring the thread up without touching Outlook at all."""
    _install_fake_pythoncom(monkeypatch)
    calls = []

    def _dispatch(progid):
        calls.append(progid)
        return _FakeOutlook(_HostileNamespace())

    _install_fake_win32com(monkeypatch, _dispatch)

    bridge.start(timeout=5)

    assert calls == [], "start() Dispatched Outlook; connection must be lazy"
    assert bridge._outlook is None


def test_start_succeeds_while_outlook_is_closed(monkeypatch, bridge):
    """The server must come up before Outlook does."""
    _install_fake_pythoncom(monkeypatch)

    def _dispatch(progid):
        raise _FakeCOMError(-2147221005, "Invalid class string", None, None)

    _install_fake_win32com(monkeypatch, _dispatch)

    bridge.start(timeout=5)  # must not raise

    assert bridge._thread is not None and bridge._thread.is_alive()


def test_startup_never_reads_the_address_book(monkeypatch, bridge):
    """Regression: CurrentUser/DefaultStore raise E_ABORT under corporate policy.

    Neither start() nor a normal call may touch them.
    """
    _install_fake_pythoncom(monkeypatch)
    namespace = _HostileNamespace()
    _install_fake_win32com(monkeypatch, lambda progid: _FakeOutlook(namespace))

    bridge.start(timeout=5)

    def _work(outlook, ns):
        return ns.GetDefaultFolder(6)

    assert _run(bridge.call(_work)) == "inbox-folder"
    assert namespace.inbox_reads == 1


def test_first_call_connects_once_and_reuses_the_connection(monkeypatch, bridge):
    _install_fake_pythoncom(monkeypatch)
    calls = []
    namespace = _HostileNamespace()

    def _dispatch(progid):
        calls.append(progid)
        return _FakeOutlook(namespace)

    _install_fake_win32com(monkeypatch, _dispatch)
    bridge.start(timeout=5)

    def _work(outlook, ns):
        return "ok"

    assert _run(bridge.call(_work)) == "ok"
    assert calls == ["Outlook.Application"]

    assert _run(bridge.call(_work)) == "ok"
    assert calls == ["Outlook.Application"], "second call re-Dispatched Outlook"


def test_connect_failure_surfaces_actionable_error(monkeypatch, bridge):
    _install_fake_pythoncom(monkeypatch)

    def _dispatch(progid):
        raise _FakeCOMError(-2147221005, "Invalid class string", None, None)

    _install_fake_win32com(monkeypatch, _dispatch)
    bridge.start(timeout=5)

    def _work(outlook, ns):
        return "unreachable"

    with pytest.raises(ComBridgeDisconnectedError) as exc:
        _run(bridge.call(_work))

    assert "Outlook" in str(exc.value)
    assert bridge._outlook is None, "failed connect left a half-set proxy behind"


def test_call_fails_fast_when_thread_is_not_running(bridge):
    """No 60s wait for a thread that was never started."""
    def _work(outlook, ns):
        return "unreachable"

    with pytest.raises(ComBridgeDisconnectedError):
        _run(bridge.call(_work))


def test_restart_after_stop_drops_stale_state(monkeypatch, bridge):
    _install_fake_pythoncom(monkeypatch)
    calls = []
    _install_fake_win32com(
        monkeypatch,
        lambda progid: (calls.append(progid), _FakeOutlook(_HostileNamespace()))[1],
    )

    bridge.start(timeout=5)

    def _work(outlook, ns):
        return "ok"

    assert _run(bridge.call(_work)) == "ok"
    bridge.stop()

    bridge.start(timeout=5)
    assert bridge._outlook is None, "restart kept a proxy from the dead apartment"
    assert _run(bridge.call(_work)) == "ok"
    assert len(calls) == 2, "restart did not force a fresh Dispatch"
