"""Unit tests for draft_html helpers."""
import os

import pytest

from outlook_desktop_mcp.utils.draft_html import (
    InvalidInlineImage,
    plain_to_html,
    prepare_inline_html,
    suggest_attachment_basename,
)


def test_plain_to_html_escapes_and_paragraphs():
    out = plain_to_html("Hello <world>\nLine 2\n\nPara 2")
    assert "&lt;world&gt;" in out
    assert "<br>Line 2" in out
    assert out.count("<p>") == 2


def test_plain_to_html_empty():
    assert plain_to_html("") == "<p></p>"


def test_prepare_no_images_uses_html_body():
    html, images = prepare_inline_html("plain", "<h1>Rich</h1>", None)
    assert html == "<h1>Rich</h1>"
    assert images == []


def test_prepare_no_images_falls_back_to_plain():
    html, _ = prepare_inline_html("hi\nthere", None, None)
    assert "<p>hi<br>there</p>" == html


def test_prepare_appends_unreferenced_image(tmp_path):
    img = tmp_path / "logo.png"
    img.write_bytes(b"\x89PNG fake")
    html, images = prepare_inline_html("see below", None, [str(img)])
    assert len(images) == 1
    cid, path = images[0]
    assert path == str(img)
    assert f'<img src="cid:{cid}"' in html


def test_prepare_replaces_placeholder(tmp_path):
    img = tmp_path / "logo.png"
    img.write_bytes(b"\x89PNG")
    body_html = "<p>Top {{LOGO}} bottom</p>"
    html, images = prepare_inline_html(
        "",
        body_html,
        [{"path": str(img), "cid": "logo1", "placeholder": "{{LOGO}}"}],
    )
    assert "{{LOGO}}" not in html
    assert '<img src="cid:logo1"' in html
    assert images == [("logo1", str(img))]


def test_prepare_respects_existing_cid_reference(tmp_path):
    img = tmp_path / "x.png"
    img.write_bytes(b"x")
    body_html = '<p><img src="cid:myimg"></p>'
    html, _ = prepare_inline_html(
        "",
        body_html,
        [{"path": str(img), "cid": "myimg"}],
    )
    # Should not be appended again
    assert html.count('cid:myimg') == 1


def test_prepare_rejects_unc():
    with pytest.raises(InvalidInlineImage):
        prepare_inline_html("", "", ["\\\\server\\share\\img.png"])


def test_prepare_rejects_missing_file():
    with pytest.raises(InvalidInlineImage):
        prepare_inline_html("", "", ["/nonexistent/path/to/img.png"])


def test_prepare_rejects_bad_cid(tmp_path):
    img = tmp_path / "x.png"
    img.write_bytes(b"x")
    with pytest.raises(InvalidInlineImage):
        prepare_inline_html("", "", [{"path": str(img), "cid": "bad cid!"}])


def test_prepare_rejects_bad_entry_type():
    with pytest.raises(InvalidInlineImage):
        prepare_inline_html("", "", [42])


def test_suggest_attachment_basename_strips_dirs():
    name = suggest_attachment_basename("/usr/bin/evil.exe")
    assert "/" not in name
    assert name.endswith(".txt")  # dangerous extension neutralized


def test_cid_substring_boundary_does_not_false_match(tmp_path):
    """cid:logo1 in html must not be considered a reference if html has cid:logo10."""
    img = tmp_path / "logo.png"
    img.write_bytes(b"x")
    body_html = '<p><img src="cid:logo10"></p>'
    html, images = prepare_inline_html(
        "",
        body_html,
        [{"path": str(img), "cid": "logo1"}],
    )
    assert "cid:logo10" in html
    assert "cid:logo1" in html
    assert images == [("logo1", str(img))]


def test_plain_to_html_normalizes_crlf():
    out = plain_to_html("line1\r\nline2\rline3")
    assert "<br>line2<br>line3" in out
    assert "\r" not in out


def test_empty_html_body_intentionally_clears():
    html, images = prepare_inline_html("plain body", "", None)
    assert html == ""
    assert images == []


def test_none_html_body_falls_back_to_plain():
    html, _ = prepare_inline_html("plain body", None, None)
    assert "<p>plain body</p>" == html


def test_prepare_rejects_duplicate_cids(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    img = tmp_path / "x.png"
    img.write_bytes(b"x")
    with pytest.raises(InvalidInlineImage, match="duplicates an earlier"):
        prepare_inline_html(
            "",
            "<p>{{A}} {{B}}</p>",
            [
                {"path": str(img), "cid": "shared", "placeholder": "{{A}}"},
                {"path": str(img), "cid": "shared", "placeholder": "{{B}}"},
            ],
        )


def test_prepare_rejects_outside_home(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path / "elsewhere"))
    outside = "/etc/hostname"
    if not os.path.isfile(outside):
        pytest.skip("/etc/hostname not present")
    with pytest.raises(InvalidInlineImage):
        prepare_inline_html("", "", [outside])
