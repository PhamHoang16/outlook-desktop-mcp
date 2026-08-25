"""In-memory cache for contact/address-book MCP query results."""
from __future__ import annotations

import hashlib
import json
import logging
import threading
import time
from collections import OrderedDict
from typing import Any

logger = logging.getLogger("outlook_desktop_mcp.contact_cache")

# Default: 7 days
CONTACT_CACHE_TTL_SECONDS = 7 * 24 * 3600
CONTACT_CACHE_MAX_ENTRIES = 256


class ContactCache:
    """Process-local TTL cache with LRU eviction for contact query responses."""

    def __init__(
        self,
        ttl_seconds: float = CONTACT_CACHE_TTL_SECONDS,
        max_entries: int = CONTACT_CACHE_MAX_ENTRIES,
    ):
        self._ttl = ttl_seconds
        self._max_entries = max(1, max_entries)
        self._entries: OrderedDict[str, tuple[float, str]] = OrderedDict()
        self._lock = threading.Lock()

    def make_key(self, tool_name: str, **params: Any) -> str:
        """Build a stable cache key from tool name and normalized parameters."""
        normalized = json.dumps(params, sort_keys=True, default=str)
        digest = hashlib.sha256(normalized.encode("utf-8")).hexdigest()
        return f"{tool_name}:{digest}"

    def get(self, key: str) -> str | None:
        """Return cached value if present and not expired, else None."""
        with self._lock:
            entry = self._entries.get(key)
            if entry is None:
                return None
            expires_at, value = entry
            if time.time() >= expires_at:
                del self._entries[key]
                return None
            self._entries.move_to_end(key)
            logger.info("contact cache hit: %s", key.split(":", 1)[0])
            return value

    def set(self, key: str, value: str) -> None:
        """Store a value with TTL from now; evict LRU entry if at capacity."""
        with self._lock:
            if key in self._entries:
                del self._entries[key]
            self._entries[key] = (time.time() + self._ttl, value)
            self._entries.move_to_end(key)
            while len(self._entries) > self._max_entries:
                self._entries.popitem(last=False)

    def clear(self) -> None:
        """Remove all entries (mainly for tests)."""
        with self._lock:
            self._entries.clear()

    def __len__(self) -> int:
        with self._lock:
            return len(self._entries)
