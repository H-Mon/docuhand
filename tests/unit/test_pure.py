"""CI-safe unit tests — zero COM dependencies (pure logic only)."""

from pathlib import Path

import pytest

from docuhand.engine.office_app import (
    format_matches_extension,
    owner_lock_file,
    probe_writable,
    sniff_format,
)
from docuhand.errors import PathNotAllowedError
from docuhand.safety.allowlist import check_allowed, get_allowlist

OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP = b"PK\x03\x04"


class TestSniffFormat:
    def test_ole2(self):
        assert sniff_format(OLE2 + b"rest-of-cfb") == "ole2"

    def test_zip(self):
        assert sniff_format(ZIP + b"xxxx") == "zip"

    def test_rtf(self):
        assert sniff_format(b"{\\rtf1 hello") == "rtf"

    def test_html(self):
        assert sniff_format(b"  <html><body>hi</body></html>") == "html"

    def test_text(self):
        assert sniff_format(b"plain ascii text file\r\nline2") == "text"

    def test_unknown(self):
        assert sniff_format(b"\x00\x01\x02\x00\x01\x02") == "unknown"

    def test_doc_renamed_to_txt_keeps_ole2(self):
        # a real .doc renamed to .txt is still OLE2 — sniff ignores extension
        assert sniff_format(OLE2) == "ole2"


class TestFormatMatchesExtension:
    @pytest.mark.parametrize(
        ("ext", "detected", "expected"),
        [
            (".doc", "ole2", True),
            (".docx", "zip", True),
            (".txt", "ole2", False),  # renamed file mismatch
            (".rtf", "zip", False),
        ],
    )
    def test_matrix(self, ext, detected, expected):
        assert format_matches_extension(ext, detected) is expected


class TestProbeWritable:
    def test_new_file_is_writable(self, tmp_path):
        f = tmp_path / "a.txt"
        f.write_text("x")
        assert probe_writable(str(f)) is True

    def test_missing_file_not_writable(self, tmp_path):
        assert probe_writable(str(tmp_path / "nope.bin")) is False


class TestOwnerLockFile:
    def test_no_lockfile(self, tmp_path):
        f = tmp_path / "report.doc"
        f.write_text("x")
        assert owner_lock_file(str(f)) is None

    def test_word_lockfile_detected(self, tmp_path):
        # Word truncates the owner lockfile name to 8+3: report.doc -> ~$port.doc
        f = tmp_path / "report.doc"
        f.write_text("x")
        (tmp_path / "~$port.doc").write_text("x")
        assert owner_lock_file(str(f)) == "~$port.doc"


class TestAllowlist:
    def test_empty_allows_all(self, monkeypatch):
        monkeypatch.delenv("DOCUHAND_ALLOWLIST", raising=False)
        assert get_allowlist() == []
        check_allowed(r"D:\anywhere\file.doc")  # must not raise

    def test_inside_allowed(self, monkeypatch, tmp_path):
        monkeypatch.setenv("DOCUHAND_ALLOWLIST", str(tmp_path))
        check_allowed(str(tmp_path / "sub" / "file.doc"))  # must not raise

    def test_outside_rejected(self, monkeypatch, tmp_path):
        monkeypatch.setenv("DOCUHAND_ALLOWLIST", str(tmp_path / "docs"))
        with pytest.raises(PathNotAllowedError):
            check_allowed(r"C:\Windows\system.ini")
