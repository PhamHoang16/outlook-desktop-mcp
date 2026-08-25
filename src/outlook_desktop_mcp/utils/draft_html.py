"""HTML body + inline-image helpers for draft email tools.

The Outlook draft tools accept any combination of:

- A plain-text ``body`` (always required as text fallback for clients that
  cannot render HTML).
- An optional ``html_body`` for rich formatting.
- An optional ``inline_images`` list: each entry either a string path or a
  dict ``{"path": "/abs/path/to.png", "cid": "img1", "placeholder": "{IMG1}"}``.

The helper normalises the inline image list, generates stable CIDs when none
are provided, substitutes placeholders inside the HTML, and appends any
remaining images at the end of the document. The resulting HTML references
each image as ``<img src="cid:CID">`` so MUAs render them inline once the
matching attachment has its ``PR_ATTACH_CONTENT_ID`` set.

This module deliberately performs no filesystem reads beyond ``os.path.isfile``
so it stays trivially testable in CI without real Outlook.
"""
from __future__ import annotations

import hashlib
import os
import re
from html import escape as html_escape
from typing import Iterable

_CID_REFERENCE_BOUNDARY = re.compile(r"[A-Za-z0-9._\-]")

from outlook_desktop_mcp.utils.attachment_safety import (
    UnsafeAttachmentPath,
    sanitize_attachment_filename,
    validate_readable_attachment,
)


_CID_RE = re.compile(r"^[A-Za-z0-9._\-]{1,60}$")


class InvalidInlineImage(ValueError):
    """Raised when an inline image entry is malformed or references a bad path."""


def _generate_cid(path: str, index: int) -> str:
    digest = hashlib.sha1(
        os.path.abspath(path).encode("utf-8"), usedforsecurity=False
    ).hexdigest()[:10]
    return f"img{index}_{digest}"


def _normalize_image_entry(entry, index: int) -> tuple[str, str, str | None]:
    """Return ``(cid, path, placeholder)`` for one inline_images entry."""
    if isinstance(entry, str):
        path = entry
        cid = None
        placeholder = None
    elif isinstance(entry, dict):
        path = entry.get("path") or entry.get("file")
        cid = entry.get("cid")
        placeholder = entry.get("placeholder")
    else:
        raise InvalidInlineImage(
            f"inline_images[{index}] must be a string path or dict, got {type(entry).__name__}"
        )

    if not path or not isinstance(path, str):
        raise InvalidInlineImage(f"inline_images[{index}].path is required")
    try:
        abs_path = validate_readable_attachment(
            path, label=f"inline_images[{index}]"
        )
    except UnsafeAttachmentPath as e:
        raise InvalidInlineImage(str(e)) from e
    if cid is None:
        cid = _generate_cid(abs_path, index)
    if not _CID_RE.match(cid):
        raise InvalidInlineImage(
            f"inline_images[{index}].cid must match [A-Za-z0-9._-]{{1,60}}: {cid!r}"
        )
    return cid, abs_path, placeholder


def plain_to_html(body: str) -> str:
    """Convert a plain-text body to safe HTML (escape + wrap paragraphs)."""
    normalized = (body or "").replace("\r\n", "\n").replace("\r", "\n")
    escaped = html_escape(normalized, quote=False)
    paragraphs = escaped.split("\n\n")
    return "".join(
        "<p>" + p.replace("\n", "<br>") + "</p>"
        for p in paragraphs
    ) or "<p></p>"


def _cid_already_referenced(html: str, cid: str) -> bool:
    """True if ``cid`` is referenced inside an existing ``cid:`` URI in html.

    Uses a non-CID-char boundary so that ``cid:logo1`` does not falsely match
    ``cid:logo10``.
    """
    needle = f"cid:{cid}"
    start = 0
    while True:
        idx = html.find(needle, start)
        if idx == -1:
            return False
        end = idx + len(needle)
        next_char = html[end:end + 1]
        if not next_char or not _CID_REFERENCE_BOUNDARY.match(next_char):
            return True
        start = end


def prepare_inline_html(
    body,
    html_body,
    inline_images: Iterable | None,
) -> tuple[str, list[tuple[str, str]]]:
    """Build the final HTML body and a list of ``(cid, abs_path)`` to attach.

    - ``html_body`` may be ``None`` (build from ``body``), the empty string
      (intentional clear: result is ``""`` plus appended inline-image tags
      if any), or a string of HTML used as-is.
    - When ``html_body`` is ``None``, ``body`` is converted to HTML with
      paragraph and line-break preservation.
    - Each inline image entry's placeholder (if any) is substituted with the
      ``<img>`` tag. Images whose CID is not already referenced in the HTML
      are appended at the end inside their own ``<p>``.
    """
    images = list(inline_images or [])
    normalized: list[tuple[str, str, str | None]] = []
    seen_cids: set[str] = set()
    for index, entry in enumerate(images):
        cid, abs_path, placeholder = _normalize_image_entry(entry, index)
        if cid in seen_cids:
            raise InvalidInlineImage(
                f"inline_images[{index}].cid duplicates an earlier image: {cid!r}"
            )
        seen_cids.add(cid)
        normalized.append((cid, abs_path, placeholder))

    # html_body=="" is intentional clear; only convert from plain body when
    # the caller did not supply any html_body at all (None).
    if html_body is None:
        html = plain_to_html(body)
    else:
        html = html_body

    for cid, _path, placeholder in normalized:
        img_tag = f'<img src="cid:{cid}" alt="">'
        if placeholder and placeholder in html:
            html = html.replace(placeholder, img_tag)
        elif not _cid_already_referenced(html, cid):
            html = html + f'<p>{img_tag}</p>'

    return html, [(cid, path) for cid, path, _ in normalized]


def suggest_attachment_basename(path: str) -> str:
    """Return a safe basename for a file path used as an inline-image attachment."""
    try:
        return sanitize_attachment_filename(os.path.basename(path) or "image")
    except UnsafeAttachmentPath:
        return "image"
