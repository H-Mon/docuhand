"""Unit tests for detect_encryption — synthetic fixtures, zero COM.

Real-password fixtures can't be built without the engine (SaveAs2
Password= is unreliable through late binding), so we forge the container
tells directly: a real CFB with a WordDocument stream whose FIB carries
fEncrypted=0x0100, plus the OOXML EncryptedPackage variant.
"""

import io
import sys
import tempfile
from pathlib import Path

import olefile
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docuhand.engine.container_guard import _fib_is_encrypted, detect_encryption

OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _write_cfb(path: Path, streams: dict[str, bytes]) -> None:
    """Build a real compound file with the given top-level streams."""
    from helpers_cfb import write_cfb

    write_cfb(str(path), streams)


class TestFibCheck:
    def test_clear_fib(self):
        # wIdent 0xA5EC + flags without fEncrypted
        head = b"\xec\xa5" + b"\x00" * 8 + b"\x00\x00" + b"\x00" * 2
        assert _fib_is_encrypted(head) is False

    def test_other_flags_do_not_trigger(self):
        # 0x0004 = fComplex; encrypted bit is 0x0100, so this is clear
        head = b"\xec\xa5" + b"\x00" * 8 + b"\x04\x00" + b"\x00" * 2
        assert _fib_is_encrypted(head) is False

    def test_fEncrypted_true(self):
        # 0x0100 little-endian = b"\x00\x01" at offset 0x0A
        head = b"\xec\xa5" + b"\x00" * 8 + b"\x00\x01" + b"\x00\x00"
        assert _fib_is_encrypted(head) is True

    def test_wrong_magic_rejected(self):
        head = b"\x00\x00" + b"\x00" * 8 + b"\x01\x01" + b"\x00\x00"
        assert _fib_is_encrypted(head) is False

    def test_short_stream_rejected(self):
        assert _fib_is_encrypted(b"\xec\xa5\x00") is False


class TestDetectEncryption:
    def test_plain_zip_is_not_encrypted(self, tmp_path):
        p = tmp_path / "a.docx"
        p.write_bytes(b"PK\x03\x04" + b"\x00" * 100)
        assert detect_encryption(str(p)) is False

    def test_text_disguised_as_doc_is_not_encrypted(self, tmp_path):
        p = tmp_path / "a.doc"
        p.write_text("hello")
        assert detect_encryption(str(p)) is False

    def test_missing_file_not_encrypted(self, tmp_path):
        assert detect_encryption(str(tmp_path / "no.doc")) is False

    def test_real_doc_from_engine_fixture_is_not_encrypted(self):
        # the repo's own demo .doc was produced by a real engine unencrypted
        demo = Path(__file__).resolve().parents[2] / "examples" / "demo-files" / "DocuHand_Demo.doc"
        if not demo.exists():
            pytest.skip("run scripts/make_demo_docs.py first")
        assert detect_encryption(str(demo)) is False

    def test_encrypted_ooxml_style_cfb_detected(self, tmp_path):
        """Word's 'Encrypt with password' on a .docx wraps it in an OLE2
        container with an EncryptedPackage stream — detect via stream name."""
        p = tmp_path / "encrypted.docx"
        _write_cfb(p, {"EncryptedPackage": b"\x00" * 4200,
                       "EncryptedSummaryInformation": b"\x00" * 4096})
        assert detect_encryption(str(p)) is True

    def test_encrypted_legacy_doc_fib_detected(self, tmp_path):
        """Legacy .doc encryption = fEncrypted bit in the FIB."""
        fib = b"\xec\xa5" + b"\x00" * 8 + b"\x00\x01" + b"\x00" * 4
        p = tmp_path / "encrypted.doc"
        # >= 4096B so the stream lands in regular sectors (this writer
        # implements no mini-stream); FIB says fEncrypted → must detect
        _write_cfb(p, {"WordDocument": fib + b"\x00" * 4200, "1Table": b"\x00" * 4096})
        assert detect_encryption(str(p)) is True

    def test_unencrypted_legacy_doc_fib_clear(self, tmp_path):
        fib = b"\xec\xa5" + b"\x00" * 8 + b"\x04\x00" + b"\x00" * 4  # fComplex only
        p = tmp_path / "plain.doc"
        _write_cfb(p, {"WordDocument": fib + b"\x00" * 4200, "1Table": b"\x00" * 4096})
        assert detect_encryption(str(p)) is False
