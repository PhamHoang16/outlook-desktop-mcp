"""Unit tests for attachment path safety."""
import os
import tempfile

import pytest

from outlook_desktop_mcp.utils.attachment_safety import (
    UnsafeAttachmentPath,
    default_attachment_dir,
    resolve_attachment_path,
    sanitize_attachment_filename,
    validate_readable_attachment,
    validate_save_directory,
)


def test_default_directory_under_home():
    p = default_attachment_dir()
    assert p.startswith(os.path.expanduser("~"))


def test_empty_save_directory_uses_default():
    assert validate_save_directory("") == default_attachment_dir()


@pytest.mark.parametrize("unc", ["\\\\server\\share", "//host/share/path"])
def test_unc_paths_rejected(unc):
    with pytest.raises(UnsafeAttachmentPath):
        validate_save_directory(unc)


def test_save_outside_home_or_tmp_rejected():
    with pytest.raises(UnsafeAttachmentPath):
        validate_save_directory("/etc")


def test_save_in_tmp_allowed():
    p = validate_save_directory(tempfile.gettempdir())
    assert os.path.isabs(p)


@pytest.mark.parametrize(
    "raw,expected",
    [
        ("hello.pdf", "hello.pdf"),
        ("../etc/passwd", "passwd"),
        ("..\\..\\boot.ini", "boot.ini"),
        ("safe (1).docx", "safe (1).docx"),
        ("evil.exe", "evil.exe.txt"),
        ("script.bat", "script.bat.txt"),
        ("", "attachment"),
        (".hidden", "hidden"),
        ("...weird", "weird"),
        ("a/b/c.txt", "c.txt"),
    ],
)
def test_sanitize_attachment_filename(raw, expected):
    out = sanitize_attachment_filename(raw)
    assert "/" not in out and "\\" not in out
    assert out == expected


def test_resolve_attachment_path_under_directory(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    subdir = tmp_path / "downloads"
    out = resolve_attachment_path(str(subdir), "Report.pdf")
    assert out.endswith("Report.pdf")
    assert os.path.dirname(out) == os.path.realpath(str(subdir))


def test_resolve_attachment_rejects_traversal(tmp_path):
    p = resolve_attachment_path(str(tmp_path), "../../etc/passwd")
    assert os.path.dirname(p) == os.path.realpath(str(tmp_path))


def test_validate_readable_attachment_under_home(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    f = tmp_path / "report.pdf"
    f.write_bytes(b"hello")
    out = validate_readable_attachment(str(f))
    assert out == os.path.realpath(str(f))


def test_validate_readable_attachment_under_tmp_allowed():
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        tmp.write(b"x")
        path = tmp.name
    try:
        assert validate_readable_attachment(path) == os.path.realpath(path)
    finally:
        os.unlink(path)


def test_validate_readable_attachment_rejects_outside_home_and_tmp(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path / "home"))
    # /etc is outside both HOME and /tmp on Linux/macOS, even if /etc/hosts exists.
    candidate = "/etc/hosts"
    if not os.path.isfile(candidate):
        pytest.skip("/etc/hosts not present on this platform")
    with pytest.raises(UnsafeAttachmentPath):
        validate_readable_attachment(candidate)


def test_validate_readable_attachment_rejects_missing_file(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    with pytest.raises(UnsafeAttachmentPath):
        validate_readable_attachment(str(tmp_path / "does_not_exist.pdf"))


def test_validate_readable_attachment_rejects_unc():
    with pytest.raises(UnsafeAttachmentPath):
        validate_readable_attachment("\\\\server\\share\\file.txt")


def test_validate_readable_attachment_rejects_oversize(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    f = tmp_path / "big.bin"
    f.write_bytes(b"x" * 100)
    with pytest.raises(UnsafeAttachmentPath):
        validate_readable_attachment(str(f), max_bytes=10)


def test_validate_readable_attachment_label_in_error(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    with pytest.raises(UnsafeAttachmentPath, match=r"inline_images\[2\]"):
        validate_readable_attachment(
            str(tmp_path / "missing.png"), label="inline_images[2]"
        )
