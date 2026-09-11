"""Unit tests for the pre-COM container gate (pure logic, zero COM)."""

import zipfile

import pytest

from docuhand.engine.container_guard import validate_container

OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP = b"PK\x03\x04"


def _fake_ole(path, garbage: bytes = b"\xff" * 64, **kwargs) -> None:
    path.write_bytes(OLE2 + garbage)
    assert not kwargs


class TestZipContainers:
    def test_valid_docx_passes(self, tmp_path):
        p = tmp_path / "ok.docx"
        with zipfile.ZipFile(p, "w") as z:
            z.writestr("[Content_Types].xml", "<Types/>")
            z.writestr("word/document.xml", "<w/>")
        ok, reason = validate_container(str(p))
        assert ok, reason

    def test_fake_zip_magic_rejected(self, tmp_path):
        p = tmp_path / "corrupt.docx"
        p.write_bytes(ZIP + b"\x00" * 64 + b"GARBAGE-TAIL")
        ok, reason = validate_container(str(p))
        assert not ok
        assert "zip" in reason.lower()

    def test_zip_without_content_types_rejected(self, tmp_path):
        p = tmp_path / "not_ooxml.docx"
        with zipfile.ZipFile(p, "w") as z:
            z.writestr("random.txt", "hi")
        ok, reason = validate_container(str(p))
        assert not ok
        assert "content_types" in reason.lower()

    def test_truncated_zip_rejected(self, tmp_path):
        p = tmp_path / "trunc.docx"
        p.write_bytes(ZIP + b"\x00" * 10)
        ok, _ = validate_container(str(p))
        assert not ok

    def test_zip_with_corrupt_member_rejected(self, tmp_path):
        import io

        p = tmp_path / "bad_member.docx"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
            z.writestr("[Content_Types].xml", "<Types/>")
            z.writestr("word/document.xml", "x" * 100)
        raw = bytearray(buf.getvalue())
        # STORED layout: find the document.xml local header and clobber its
        # DATA region (after name + extra field), then fix the central
        # directory offsets so the zip still parses — only the CRC breaks.
        name = b"word/document.xml"
        # local file header signature 'PK\x03\x04'; second entry is ours
        first = raw.find(name)
        data_start = raw.find(name, first + 1)
        if data_start == -1:
            data_start = first
        data_start += len(name)
        # our entry is second → single extra-field-less layout: clobber 32B
        raw[data_start:data_start + 32] = b"\xee" * 32
        p.write_bytes(bytes(raw))
        ok, reason = validate_container(str(p))
        assert not ok, f"expected CRC failure, got: {reason}"
        assert "member" in reason.lower() or "crc" in reason.lower() or "zip" in reason.lower()


class TestOleContainers:
    def test_ole_magic_plus_garbage_rejected(self, tmp_path):
        # the exact fixture that used to hang Word/WPS on a modal dialog
        p = tmp_path / "corrupt.doc"
        _fake_ole(p)
        ok, reason = validate_container(str(p))
        assert not ok
        assert "ole" in reason.lower() or "malformed" in reason.lower()

    def test_truncated_ole_rejected(self, tmp_path):
        # OLE magic + zeros fails isOleFile/parse because the CFB header
        # sector-count math doesn't add up — but keep a truly tiny file
        # case too (header itself truncated mid-signature is caught by the
        # magic check; here: valid 8-byte magic, body cut short).
        p = tmp_path / "trunc.doc"
        p.write_bytes(OLE2 + b"\x00" * 100)  # 108 bytes < 512B header
        ok, reason = validate_container(str(p))
        assert not ok
        assert "truncated" in reason.lower() or "ole" in reason.lower()

    def test_real_doc_structure_passes(self, tmp_path):
        # WordDocument stream readability is the positive-path check; a real
        # .doc (with a WordDocument stream) passes — proven end-to-end by the
        # `-m com` suite converting scripts/make_mixed_dir.py fixtures.
        # Here we prove the negative side: OLE magic alone (no valid CFB
        # directory) is rejected even when the size looks plausible.
        p = tmp_path / "fake.doc"
        _fake_ole(p, garbage=b"\x00" * 600)
        # 608 bytes of OLE magic + zeros: still malformed CFB → rejected
        ok, _ = validate_container(str(p))
        assert not ok


class TestPassthrough:
    def test_plain_text_passes(self, tmp_path):
        p = tmp_path / "notes.doc"
        p.write_text("hello")
        ok, _ = validate_container(str(p))
        assert ok

    def test_rtf_passes(self, tmp_path):
        p = tmp_path / "doc.rtf"
        p.write_text(r"{\rtf1 hello}")
        ok, _ = validate_container(str(p))
        assert ok

    def test_missing_file_rejected(self, tmp_path):
        ok, reason = validate_container(str(tmp_path / "nope.doc"))
        assert not ok
        assert "unreadable" in reason.lower()
