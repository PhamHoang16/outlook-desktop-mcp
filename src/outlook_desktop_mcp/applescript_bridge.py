"""
AppleScript Bridge
==================
Runs Outlook automation via osascript subprocess calls on macOS.
Each call is stateless — no persistent COM-like objects.

Every tool builds an AppleScript string and passes it to bridge.run(). The
bridge injects shared helper handlers (e.g. ``_odm_make_date``) at the top of
every script and enforces a maximum script size to avoid exceeding the
osascript argv limit.
"""
import asyncio
import logging
import os

from outlook_desktop_mcp.utils.applescript_helpers import ODM_DATE_HANDLER

logger = logging.getLogger("outlook_desktop_mcp.applescript_bridge")

STARTUP_TIMEOUT = float(os.environ.get("OUTLOOK_MCP_APPLESCRIPT_STARTUP_TIMEOUT", "10"))
SCRIPT_TIMEOUT = float(os.environ.get("OUTLOOK_MCP_APPLESCRIPT_TIMEOUT", "30"))

# osascript -e takes one argv slot. macOS argv limits are large but a single
# string >256KB starts to hit problems on some shells / process snapshots.
MAX_SCRIPT_BYTES = int(os.environ.get("OUTLOOK_MCP_APPLESCRIPT_MAX_BYTES", str(200 * 1024)))


class AppleScriptError(RuntimeError):
    """AppleScript returned a non-zero exit code."""


class AppleScriptTimeoutError(RuntimeError):
    """AppleScript subprocess exceeded the timeout."""


class AppleScriptScriptTooLargeError(RuntimeError):
    """Compiled AppleScript exceeded MAX_SCRIPT_BYTES."""


class AppleScriptBridge:
    """Manages AppleScript execution for Outlook on macOS."""

    def __init__(self):
        self._version: str | None = None

    async def start(self):
        """Verify Outlook is running and accessible. Call once at server startup."""
        try:
            self._version = await self.run(
                'tell application "Microsoft Outlook" to get version',
                timeout=STARTUP_TIMEOUT,
            )
            logger.info("AppleScript bridge ready. Outlook version: %s", self._version)
        except Exception as e:
            raise RuntimeError(
                f"Cannot connect to Microsoft Outlook via AppleScript. "
                f"Is Outlook running? Error: {e}"
            ) from e

    @staticmethod
    def _wrap_script(script: str) -> str:
        """Prefix shared helper handlers and validate size."""
        full = ODM_DATE_HANDLER + "\n" + script
        encoded = full.encode("utf-8")
        if len(encoded) > MAX_SCRIPT_BYTES:
            raise AppleScriptScriptTooLargeError(
                f"AppleScript exceeds limit ({len(encoded)} > {MAX_SCRIPT_BYTES} bytes). "
                "Reduce body size or pass the data through a file."
            )
        return full

    async def run(self, script: str, timeout: float = SCRIPT_TIMEOUT) -> str:
        """Execute an AppleScript and return stdout as a string.

        Raises AppleScriptError on non-zero exit, AppleScriptTimeoutError on
        timeout, AppleScriptScriptTooLargeError when the wrapped script exceeds
        MAX_SCRIPT_BYTES.
        """
        wrapped = self._wrap_script(script)
        proc = await asyncio.create_subprocess_exec(
            "osascript", "-e", wrapped,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        try:
            stdout, stderr = await asyncio.wait_for(
                proc.communicate(), timeout=timeout
            )
        except asyncio.TimeoutError:
            try:
                proc.kill()
                await proc.communicate()
            except Exception:
                pass
            raise AppleScriptTimeoutError(
                f"AppleScript timed out after {timeout}s"
            )

        if proc.returncode != 0:
            err = stderr.decode("utf-8", errors="replace").strip()
            lower = err.lower()
            if "application isn't running" in lower or "can't find process" in lower:
                raise AppleScriptError(
                    "Microsoft Outlook is not running. Start Outlook for Mac and retry."
                )
            raise AppleScriptError(f"AppleScript error: {err}")

        return stdout.decode("utf-8", errors="replace").rstrip("\n")

    async def run_lines(self, script: str, timeout: float = SCRIPT_TIMEOUT) -> list[str]:
        """Execute an AppleScript and return output split into non-empty lines."""
        result = await self.run(script, timeout=timeout)
        return [line for line in result.split("\n") if line.strip()]

    def stop(self):
        """No-op — AppleScript has no persistent resources to release."""
        pass
