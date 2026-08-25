"""
COM Threading Bridge
====================
Runs all Outlook COM calls on a dedicated STA (Single-Threaded Apartment)
thread so the async MCP event loop never touches COM objects directly.

Every COM function passed to bridge.call() receives (outlook, namespace, ...)
as its first two arguments — the live COM objects that only exist on the
COM thread.

The bridge serializes work onto a single STA thread. To avoid one slow
operation cascading into wedging the whole server, ``call`` supports a
per-call timeout and refuses to queue a new request while a previous one
is still in flight (the COM thread cannot be cancelled mid-COM-call).

When Outlook crashes or is force-closed, the cached ``Outlook.Application``
and ``MAPI`` namespace references become RPC-disconnected. The bridge
detects those errors and re-Dispatches once before propagating the failure.

Connecting to Outlook is *lazy*: ``start()`` only brings the COM thread up,
and the first queued call performs the Dispatch. This keeps the MCP server
startable when Outlook is not running yet, and avoids touching the Exchange
address book at startup — some corporate security policies block
``Namespace.CurrentUser``, which used to abort startup with E_ABORT.
"""
import asyncio
import logging
import os
import queue
import sys
import threading

logger = logging.getLogger("outlook_desktop_mcp.com_bridge")

DEFAULT_CALL_TIMEOUT = float(os.environ.get("OUTLOOK_MCP_COM_TIMEOUT", "60"))
DEFAULT_START_TIMEOUT = float(os.environ.get("OUTLOOK_MCP_COM_START_TIMEOUT", "15"))

# HRESULTs that mean "the Outlook process / RPC channel is gone". Re-Dispatch
# may recover; we only retry once per call to avoid masking persistent errors.
_RPC_DISCONNECTED_HRESULTS = frozenset({
    0x80010108,  # RPC_E_DISCONNECTED
    0x800706BA,  # RPC server unavailable
    0x800706BE,  # remote procedure call failed
    0x80010105,  # RPC_E_SERVERFAULT
    0x800706BF,  # RPC failed, did not execute
    0x800401FD,  # CO_E_OBJNOTCONNECTED
})


def _is_rpc_disconnected(exc: Exception) -> bool:
    """Return True if ``exc`` indicates the Outlook process / RPC channel is gone."""
    hresult = getattr(exc, "hresult", None)
    if hresult is None:
        args = getattr(exc, "args", ())
        if args and isinstance(args[0], int):
            hresult = args[0]
    if hresult is None:
        return False
    return (int(hresult) & 0xFFFFFFFF) in _RPC_DISCONNECTED_HRESULTS


class ComBridgeBusyError(RuntimeError):
    """Raised when a new COM request is submitted while another is in flight."""

    code = "com_busy"
    retriable = True


class ComBridgeTimeoutError(TimeoutError):
    """Raised when a COM call exceeds its timeout.

    Important: COM has no cancellation. The underlying operation (Send,
    Delete, Move, etc.) may still complete on the Outlook side after the
    caller sees this error. Callers performing destructive operations should
    surface ``retriable=False`` and prompt the user to verify in the Outlook
    UI before retrying.
    """

    code = "com_timeout"
    retriable = False


class ComBridgeDisconnectedError(RuntimeError):
    """Raised when Outlook RPC is disconnected and re-Dispatch failed."""

    code = "com_disconnected"
    retriable = True


class OutlookBridge:
    """Manages a dedicated COM thread for Outlook operations."""

    def __init__(self):
        self._thread: threading.Thread | None = None
        self._request_queue: queue.Queue = queue.Queue()
        self._outlook = None
        self._namespace = None
        self._ready = threading.Event()
        self._shutdown = threading.Event()
        self._init_error: Exception | None = None
        self._alive = threading.Event()
        self._in_flight_lock = threading.Lock()
        self._in_flight_label: str | None = None

    def start(self, timeout: float = DEFAULT_START_TIMEOUT):
        """Start the COM thread. Call once at server startup.

        Waits only for the COM apartment to initialize — not for a connection
        to Outlook. Outlook is Dispatched lazily on the first queued call, so
        this succeeds even when Outlook is not running yet.
        """
        if self._thread is not None and self._thread.is_alive():
            logger.warning("OutlookBridge.start() called twice; ignoring.")
            return
        # Reset state so a restart after stop() behaves like a fresh start.
        self._shutdown.clear()
        self._ready.clear()
        self._alive.clear()
        self._init_error = None
        # A fresh thread means a fresh COM apartment; any cached proxy from a
        # previous thread is invalid.
        self._outlook = None
        self._namespace = None
        # Drop requests left over from a previous lifecycle so they do not
        # re-execute on the new thread.
        self._drain_pending("COM bridge restarted before this request ran")

        self._thread = threading.Thread(
            target=self._com_thread_main, daemon=True, name="outlook-com"
        )
        self._thread.start()
        if not self._ready.wait(timeout=timeout):
            if self._init_error:
                raise self._init_error
            raise RuntimeError(
                f"Outlook COM thread failed to initialize within {timeout}s."
            )
        if self._init_error:
            raise self._init_error

    def _connect(self) -> None:
        """Dispatch Outlook.Application and cache the MAPI namespace.

        Runs on the COM thread only. Deliberately touches nothing beyond
        Dispatch + GetNamespace: reading ``DefaultStore`` or ``CurrentUser``
        here hits the Exchange address book, which corporate policy can block
        with E_ABORT and which used to make startup fail outright.

        Never leaves a half-set _outlook/_namespace pair behind.
        """
        import win32com.client

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
        except Exception as e:
            self._outlook = None
            self._namespace = None
            raise ComBridgeDisconnectedError(
                "Could not connect to Outlook Desktop (Classic). Open Outlook "
                "and retry. The new Outlook (olk.exe) does not support COM."
            ) from e
        self._outlook = outlook
        self._namespace = namespace
        logger.debug("Connected to Outlook via COM")

    def _redispatch_outlook(self) -> None:
        """Drop cached references and re-Dispatch Outlook.Application.

        Called on the COM thread only, after detecting RPC_E_DISCONNECTED.
        Caller already holds the in-flight lock.
        """
        logger.warning("Outlook COM disconnected; attempting re-Dispatch")
        self._outlook = None
        self._namespace = None
        self._connect()
        logger.info("Outlook COM re-Dispatch succeeded")

    def _invoke(self, func, args, kwargs):
        """Run func once; connect lazily, and on RPC disconnect retry once."""
        if self._outlook is None:
            self._connect()
        try:
            return func(self._outlook, self._namespace, *args, **kwargs)
        except Exception as e:
            if not _is_rpc_disconnected(e):
                raise
            try:
                self._redispatch_outlook()
            except Exception as redispatch_err:
                raise ComBridgeDisconnectedError(
                    "Outlook RPC is disconnected and re-Dispatch failed. "
                    "Restart Outlook Desktop and try again."
                ) from redispatch_err
            return func(self._outlook, self._namespace, *args, **kwargs)

    def _com_thread_main(self):
        """Main loop for the COM thread.

        Does NOT connect to Outlook here — see _invoke / _connect. The thread
        only needs its COM apartment to be ready before it accepts work.
        """
        import pythoncom

        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        except Exception as e:
            self._init_error = e
            logger.error("COM apartment init failed: %s", e)
            self._ready.set()
            return

        self._alive.set()
        self._ready.set()
        try:
            while not self._shutdown.is_set():
                try:
                    func, args, kwargs, result_event, result_holder = (
                        self._request_queue.get(timeout=0.5)
                    )
                except queue.Empty:
                    continue
                try:
                    result_holder["value"] = self._invoke(func, args, kwargs)
                except Exception as e:
                    result_holder["error"] = e
                finally:
                    result_event.set()

            self._drain_pending("COM bridge is shutting down")
        finally:
            self._alive.clear()
            self._outlook = None
            self._namespace = None
            pythoncom.CoUninitialize()

    def _drain_pending(self, message: str):
        """Reject any queued-but-not-started requests with a clear error."""
        while True:
            try:
                _, _, _, result_event, result_holder = (
                    self._request_queue.get_nowait()
                )
            except queue.Empty:
                return
            result_holder["error"] = RuntimeError(message)
            result_event.set()

    async def call(self, func, *args, timeout: float | None = None, **kwargs):
        """
        Schedule a function on the COM thread and await its result.

        Refuses to queue if another call is currently *awaited* by the caller
        path: the STA thread can only run one COM call at a time and slow
        operations should not silently stack up. The caller should retry after
        a short delay.

        Note that on ``ComBridgeTimeoutError`` the in-flight COM call keeps
        running on the COM thread (COM has no cancellation). The bridge lock
        is released when the caller's ``call`` returns, so subsequent calls
        will queue behind the still-running operation; if you need hard
        back-pressure, catch the timeout and refuse new work yourself.

        ``timeout`` defaults to ``DEFAULT_CALL_TIMEOUT`` (60s).
        """
        timeout_val = DEFAULT_CALL_TIMEOUT if timeout is None else float(timeout)
        label = getattr(func, "__name__", "<anonymous>")

        # Fail fast rather than making the caller wait out the full timeout
        # for a thread that is not there to answer.
        if self._init_error:
            raise self._init_error
        if self._thread is None or not self._thread.is_alive():
            raise ComBridgeDisconnectedError(
                "COM bridge is not running. Restart the MCP server."
            )
        if not self._alive.is_set():
            raise ComBridgeDisconnectedError(
                "COM thread exited unexpectedly. Restart the MCP server."
            )

        if not self._in_flight_lock.acquire(blocking=False):
            raise ComBridgeBusyError(
                f"COM thread is busy with previous request "
                f"({self._in_flight_label or 'unknown'}). Retry shortly."
            )
        try:
            self._in_flight_label = label
            result_event = threading.Event()
            result_holder: dict = {}
            self._request_queue.put((func, args, kwargs, result_event, result_holder))

            loop = asyncio.get_running_loop()
            signaled = await loop.run_in_executor(
                None, lambda: result_event.wait(timeout=timeout_val)
            )
            if not signaled:
                raise ComBridgeTimeoutError(
                    f"Outlook COM operation '{label}' timed out after "
                    f"{timeout_val:.0f}s. Outlook may be waiting on a dialog. "
                    f"If the operation modifies data (send/delete/move), verify "
                    f"in Outlook before retrying — the call may still complete."
                )
            if "error" in result_holder:
                raise result_holder["error"]
            return result_holder.get("value")
        finally:
            self._in_flight_label = None
            self._in_flight_lock.release()

    def stop(self):
        """Signal the COM thread to shut down and reject pending requests."""
        self._shutdown.set()
        self._alive.clear()
        self._drain_pending("OutlookBridge stopped")
        if self._thread:
            self._thread.join(timeout=5)
