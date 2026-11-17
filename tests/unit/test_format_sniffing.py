"""Format sniffing and extension mapping.

The container is not the product: HWP 5.0 (Hancom, Korea) uses the same OLE2
compound container as legacy .doc, so the friendly name must not assert one
of them. These tests pin that, plus the extension mapping that keeps
"matches_extension" honest for every OLE2/ZIP variant.
"""

from __future__ import annotations

from docuhand.engine.office_app import (
    FRIENDLY_FORMATS,
    format_matches_extension,
    sniff_format,
)

OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_MAGIC = b"PK\x03\x04"


def test_sniff_ole2():
    assert sniff_format(OLE2_MAGIC + b"\x00" * 8) == "ole2"


def test_sniff_zip():
    assert sniff_format(ZIP_MAGIC + b"\x00" * 8) == "zip"


def test_sniff_rtf():
    assert sniff_format(b"{\\rtf1\\ansi") == "rtf"


def test_sniff_html():
    assert sniff_format(b"<!DOCTYPE html><html>") == "html"


def test_sniff_text():
    assert sniff_format(b"plain ascii text") == "text"


def test_sniff_unknown():
    assert sniff_format(b"\x00\x01\x02\x03\xff\xfe") == "unknown"


def test_ole2_friendly_name_does_not_assert_a_single_product():
    """HWP 5.0 shares the OLE2 container - naming only .doc misreports it."""
    friendly = FRIENDLY_FORMATS["ole2"].lower()
    assert "hwp" in friendly
    # must not read as "this container is a .doc and nothing else"
    assert friendly.startswith("ole2")


def test_extension_mapping_covers_legacy_variants():
    assert format_matches_extension(".doc", "ole2")
    assert format_matches_extension(".dot", "ole2")
    # Korean formats: HWP 5.0 is OLE2, HWPX is zip
    assert format_matches_extension(".hwp", "ole2")
    assert format_matches_extension(".hwpx", "zip")


def test_extension_mapping_covers_ooxml_variants():
    assert format_matches_extension(".docx", "zip")
    assert format_matches_extension(".dotx", "zip")


def test_extension_mismatch_is_reported():
    assert not format_matches_extension(".xyz", "ole2")
    assert not format_matches_extension(".doc", "zip")
    assert not format_matches_extension(".docx", "ole2")


def test_extension_matching_is_case_insensitive():
    assert format_matches_extension(".DOC", "ole2")
    assert format_matches_extension(".HWP", "ole2")
