"""Helpers for building and parsing AppleScript safely."""
import re
import unicodedata
from datetime import datetime


# ---------------------------------------------------------------------------
# String escaping
# ---------------------------------------------------------------------------


def escape(text: str) -> str:
    """Escape a string for safe embedding inside AppleScript double quotes.

    Handles backslashes, double quotes, common control chars, and the Unicode
    line separators (U+2028, U+2029) that AppleScript may treat as statement
    terminators in some contexts.
    """
    text = unicodedata.normalize("NFKC", text)
    text = (
        text.replace("\u201c", '"')
        .replace("\u201d", '"')
        .replace("\u2018", "'")
        .replace("\u2019", "'")
    )
    text = text.replace("\\", "\\\\")
    text = text.replace('"', '\\"')
    text = text.replace("\n", "\\n")
    text = text.replace("\r", "\\r")
    text = text.replace("\t", "\\t")
    text = text.replace("\u2028", "\\n")
    text = text.replace("\u2029", "\\n")
    text = text.replace("\x00", "")
    text = text.replace("\x0b", " ")
    text = text.replace("\x0c", " ")
    return text


# ---------------------------------------------------------------------------
# Entry ID validation (defense against AppleScript injection)
# ---------------------------------------------------------------------------

_MAC_ENTRY_ID_RE = re.compile(r"\A\d{1,32}\Z")


class InvalidEntryIdError(ValueError):
    """Raised when an entry_id fails validation for AppleScript use."""


def validate_mac_entry_id(entry_id) -> str:
    """Return a sanitized macOS Outlook entry_id (numeric only).

    macOS Outlook entry IDs are numeric AppleScript object IDs. Any other
    characters — including embedded newlines, surrounding whitespace beyond a
    simple strip, or punctuation — are rejected to prevent AppleScript
    injection through string interpolation.
    """
    s = unicodedata.normalize("NFKC", str(entry_id)).strip()
    if not re.fullmatch(r"[0-9]{1,32}", s):
        raise InvalidEntryIdError(
            f"Invalid entry_id (must be a positive integer up to 32 digits): {entry_id!r}"
        )
    return s


def split_delimited_record(record: str, delim: str, field_count: int) -> list[str] | None:
    """Split a fixed-width delimited record without delimiter bleed into tail fields.

    Uses ``str.split(delim, field_count - 1)`` so delimiter bytes inside the
    last field (e.g. subject lines) do not shift earlier columns.
    """
    parts = record.split(delim, field_count - 1)
    if len(parts) != field_count:
        return None
    return parts


class InvalidScriptIntError(ValueError):
    """Raised when a numeric parameter cannot be coerced for AppleScript."""


def coerce_script_int(
    value,
    *,
    default: int,
    lo: int = 1,
    hi: int = 500,
    name: str = "value",
) -> int:
    """Coerce ``value`` to an ``int`` clamped to ``[lo, hi]`` for AppleScript use.

    Any string interpolated into an AppleScript snippet must be a real integer:
    if a client passes a string like ``"10\\nend tell\\ndo shell script ..."``,
    naive ``f"set maxCount to {value}"`` lets attacker bytes escape the literal
    and execute. This helper rejects bool, parses numeric strings strictly, and
    clamps the result so callers can safely embed it as ``{count}``.

    Returns ``default`` (also clamped) for ``None`` or empty strings, mirroring
    the behaviour of optional MCP parameters.
    """
    if value is None or value == "":
        coerced = default
    elif isinstance(value, bool):
        raise InvalidScriptIntError(
            f"{name} must be an integer, not a boolean: {value!r}"
        )
    elif isinstance(value, int):
        coerced = value
    elif isinstance(value, float):
        if not value.is_integer():
            raise InvalidScriptIntError(
                f"{name} must be a whole number: {value!r}"
            )
        coerced = int(value)
    elif isinstance(value, str):
        stripped = value.strip()
        if not re.fullmatch(r"-?\d+", stripped):
            raise InvalidScriptIntError(
                f"{name} must be an integer, got {value!r}"
            )
        coerced = int(stripped)
    else:
        raise InvalidScriptIntError(
            f"{name} must be an integer, got {type(value).__name__}"
        )
    return max(lo, min(coerced, hi))


# ---------------------------------------------------------------------------
# Locale-independent date handling
# ---------------------------------------------------------------------------

# Handler that constructs an AppleScript date deterministically from numeric
# components. Injected at the top of every script by AppleScriptBridge.
ODM_DATE_HANDLER = (
    "on _odm_make_date(y, m, d, h, mn, s)\n"
    "    set ddd to (current date)\n"
    "    set year of ddd to y\n"
    "    set month of ddd to m\n"
    "    set day of ddd to d\n"
    "    set hours of ddd to h\n"
    "    set minutes of ddd to mn\n"
    "    set seconds of ddd to s\n"
    "    return ddd\n"
    "end _odm_make_date\n"
)


def format_date(dt: datetime) -> str:
    """Return a locale-independent AppleScript expression for ``dt``.

    Relies on the ``_odm_make_date`` handler injected by AppleScriptBridge.
    """
    return (
        f"(my _odm_make_date({dt.year}, {dt.month}, {dt.day}, "
        f"{dt.hour}, {dt.minute}, {dt.second}))"
    )


def parse_date(text: str) -> str:
    """Parse an AppleScript date string into ISO 8601, or return raw on failure."""
    text = text.strip()
    text = re.sub(r"^\w+day,\s*", "", text)
    text = text.replace(" at ", " ")
    for fmt in (
        "%B %d, %Y %I:%M:%S %p",
        "%d. %B %Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
    ):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.isoformat()
        except ValueError:
            continue
    return text


# ---------------------------------------------------------------------------
# Folder references
# ---------------------------------------------------------------------------

# Locale-independent AppleScript folder keywords
FOLDER_MAP = {
    "inbox": "inbox",
    "sent": "sent items",
    "sentmail": "sent items",
    "sent items": "sent items",
    "drafts": "drafts",
    "deleted": "deleted items",
    "deleted items": "deleted items",
    "trash": "deleted items",
    "junk": "junk mail",
    "spam": "junk mail",
    "outbox": "outbox",
}


def resolve_folder_ref(folder_name: str) -> str:
    """Map a user-facing folder name to an AppleScript folder reference."""
    key = folder_name.lower().strip()
    if key in FOLDER_MAP:
        return FOLDER_MAP[key]
    return f'mail folder "{escape(folder_name)}"'


# ---------------------------------------------------------------------------
# Output framing
# ---------------------------------------------------------------------------

# ASCII Unit Separator / Record Separator. These bytes never appear in normal
# email content (subjects, bodies, attachment names), eliminating the
# collisions that ``|||``/``===`` had with real-world data.
DELIM = "\x1f"
RECORD_DELIM = "\x1e"
