"""Path-traversal and directory-allowlist guards for attachment saving."""
from __future__ import annotations

import os
import re
import tempfile

DEFAULT_SUBDIR = "outlook-mcp-attachments"

_UNSAFE_BASENAME = re.compile(r"[^A-Za-z0-9._\-+ ()\[\]]")
_DOUBLE_DOT = re.compile(r"\.{2,}")
_DANGEROUS_EXTS = {
    ".exe", ".bat", ".cmd", ".com", ".scr", ".lnk", ".ps1", ".vbs",
    ".js", ".jse", ".wsf", ".wsh", ".msi", ".cpl", ".reg",
}

# Default per-file cap for attachments the server reads from disk. Exchange Online
# rejects messages larger than ~25 MiB; oversized files almost always indicate a
# mistake or a path the LLM was tricked into supplying.
DEFAULT_MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024


class UnsafeAttachmentPath(ValueError):
    """Raised when a requested save location is rejected."""


def default_attachment_dir() -> str:
    """Return the default safe attachment directory under the user's profile."""
    home = os.path.expanduser("~")
    base = os.path.join(home, "Downloads")
    if not os.path.isdir(base):
        base = home
    return os.path.join(base, DEFAULT_SUBDIR)


def _is_under(child: str, parent: str) -> bool:
    try:
        common = os.path.commonpath([os.path.realpath(child), os.path.realpath(parent)])
    except (ValueError, OSError):
        return False
    return common == os.path.realpath(parent)


def _allowed_roots() -> list[str]:
    home = os.path.expanduser("~")
    roots = [home]
    tmp = tempfile.gettempdir()
    if tmp:
        roots.append(tmp)
    return [os.path.realpath(r) for r in roots]


def validate_save_directory(save_directory: str) -> str:
    """Normalize and validate ``save_directory`` for attachment writes.

    - Empty/None defaults to ``~/Downloads/outlook-mcp-attachments``.
    - Rejects UNC paths (``\\\\host\\share`` / ``//host/share``).
    - Rejects paths outside the user's home or the temp dir.

    Returns the absolute, validated directory path. Caller is responsible
    for creating it via :func:`ensure_save_directory`.
    """
    raw = (save_directory or "").strip()
    if not raw:
        return default_attachment_dir()

    if raw.startswith("\\\\") or raw.startswith("//"):
        raise UnsafeAttachmentPath(
            "UNC and network share paths are not allowed for attachment saves."
        )

    abs_path = os.path.abspath(os.path.expanduser(raw))

    if not any(_is_under(abs_path, root) for root in _allowed_roots()):
        raise UnsafeAttachmentPath(
            "Save directory must be inside the user's home directory or the system temp dir."
        )

    return abs_path


def ensure_save_directory(save_directory: str) -> str:
    """Validate and create the directory if needed; return the absolute path."""
    path = validate_save_directory(save_directory)
    os.makedirs(path, exist_ok=True)
    return path


def sanitize_attachment_filename(name: str, *, neutralize_dangerous_ext: bool = True) -> str:
    """Strip path components and unsafe characters from an attachment filename.

    - Normalises both ``/`` and ``\\`` directory separators (cross-platform).
    - Collapses repeated dots that could create ``..``-style traversal.
    - Strips leading dots so hidden-file names aren't produced.
    - Replaces characters outside ``[A-Za-z0-9._\\- +()\\[\\]]`` with ``_``.
    - Optionally renames dangerous Windows-execution extensions to ``.txt``
      to prevent accidental double-click execution after save.

    Always returns a non-empty filename (falls back to ``attachment``).
    """
    raw = (name or "").strip()
    if not raw:
        return "attachment"
    flat = raw.replace("\\", "/").rstrip("/")
    segments = [seg for seg in flat.split("/") if seg and seg not in (".", "..")]
    base = segments[-1] if segments else "attachment"
    base = _DOUBLE_DOT.sub(".", base)
    base = _UNSAFE_BASENAME.sub("_", base)
    base = base.lstrip(".")
    if not base:
        base = "attachment"
    if neutralize_dangerous_ext:
        stem, ext = os.path.splitext(base)
        if ext.lower() in _DANGEROUS_EXTS:
            base = f"{stem}{ext}.txt"
    return base


def validate_readable_attachment(
    path: str,
    *,
    max_bytes: int = DEFAULT_MAX_ATTACHMENT_BYTES,
    label: str = "attachment",
) -> str:
    """Validate that ``path`` is a regular file inside an allowed root.

    The MCP server reads files from disk to attach to outgoing mail. Without
    this guard an LLM prompt can attach arbitrary local files (SSH keys, AWS
    credentials, browser cookies) and exfiltrate them via ``send_email``.

    Rules:
    - UNC / network share paths are rejected up-front.
    - The path must resolve (via ``realpath``) to a regular file inside the
      user's home directory or the system temp dir.
    - The file size must not exceed ``max_bytes`` (default 25 MiB).

    Returns the canonical absolute path (after ``realpath`` resolution).
    Raises :class:`UnsafeAttachmentPath` with a user-actionable message on
    rejection. ``label`` is included in error messages so the caller can
    distinguish inline images from regular attachments.
    """
    if not isinstance(path, str) or not path:
        raise UnsafeAttachmentPath(f"{label} path must be a non-empty string")
    if path.startswith("\\\\") or path.startswith("//"):
        raise UnsafeAttachmentPath(
            f"UNC and network share paths are not allowed for {label}: {path!r}"
        )

    abs_path = os.path.abspath(os.path.expanduser(path))
    try:
        real_path = os.path.realpath(abs_path)
    except OSError as e:
        raise UnsafeAttachmentPath(
            f"Could not resolve {label} path {path!r}: {e}"
        ) from e

    if not os.path.isfile(real_path):
        raise UnsafeAttachmentPath(
            f"{label} file does not exist or is not a regular file: {path!r}"
        )

    roots = _allowed_roots()
    if not any(_is_under(real_path, root) for root in roots):
        raise UnsafeAttachmentPath(
            f"{label} path must be inside the user's home directory or the "
            f"system temp dir (got {path!r})."
        )

    try:
        size = os.path.getsize(real_path)
    except OSError as e:
        raise UnsafeAttachmentPath(
            f"Could not stat {label} {path!r}: {e}"
        ) from e
    if size > max_bytes:
        raise UnsafeAttachmentPath(
            f"{label} {path!r} is {size} bytes which exceeds the "
            f"{max_bytes}-byte limit; pass a smaller file or raise the limit."
        )

    return real_path


def resolve_attachment_path(save_directory: str, filename: str) -> str:
    """Return a validated absolute file path inside ``save_directory``.

    Re-checks containment after path normalization to defeat traversal even
    when the sanitized filename somehow contains directory separators.
    """
    directory = ensure_save_directory(save_directory)
    clean_name = sanitize_attachment_filename(filename)
    candidate = os.path.abspath(os.path.join(directory, clean_name))
    if not _is_under(candidate, directory):
        raise UnsafeAttachmentPath(
            "Resolved attachment path escapes save directory."
        )
    return candidate
