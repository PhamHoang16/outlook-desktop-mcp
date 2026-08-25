"""Unit tests for the _resolve_store / _require_store helpers."""
import importlib.util
import pathlib
import sys
import types

import pytest


def _load_resolve_helpers():
    path = pathlib.Path(__file__).resolve().parents[2] / "src" / "outlook_desktop_mcp" / "server.py"
    text = path.read_text(encoding="utf-8")
    # Extract from `class AmbiguousAccountError` through end of `_require_store`.
    start = text.index("class AmbiguousAccountError")
    end = text.index("def _resolve_account_object")
    snippet = text[start:end]
    ns: dict = {}
    exec(snippet, ns)
    return ns["_resolve_store"], ns["_require_store"], ns["AmbiguousAccountError"]


_resolve_store, _require_store, AmbiguousAccountError = _load_resolve_helpers()


class _Store:
    def __init__(self, display_name):
        self.DisplayName = display_name
        self.StoreID = display_name


class _Stores:
    def __init__(self, items):
        self._items = items

    @property
    def Count(self):
        return len(self._items)

    def Item(self, i):
        return self._items[i - 1]


class _Namespace:
    def __init__(self, stores, default_index=0):
        self.Stores = _Stores(stores)
        self.DefaultStore = stores[default_index] if stores else None


def test_default_when_account_empty():
    ns = _Namespace([_Store("A"), _Store("B")], default_index=1)
    assert _resolve_store(ns, "").DisplayName == "B"


def test_exact_match_preferred_over_substring():
    ns = _Namespace([_Store("Work"), _Store("Work Mailbox"), _Store("Network")])
    assert _resolve_store(ns, "Work").DisplayName == "Work"


def test_unique_substring_match():
    ns = _Namespace([_Store("Personal"), _Store("Engineering Team")])
    assert _resolve_store(ns, "engineering").DisplayName == "Engineering Team"


def test_ambiguous_substring_raises():
    ns = _Namespace([_Store("Engineering"), _Store("Engineering Archive")])
    with pytest.raises(AmbiguousAccountError):
        _resolve_store(ns, "engin")


def test_no_match_returns_none():
    ns = _Namespace([_Store("Work")])
    assert _resolve_store(ns, "nope") is None


def test_require_store_raises_on_missing():
    ns = _Namespace([_Store("Work")])
    with pytest.raises(ValueError):
        _require_store(ns, "nope")
