"""Pure planning for convert_documents — zero COM, fully unit-testable.

Kept out of the tool layer so CI can verify the naming/classification rules
without an Office installation. The engine only sees finished (source,
target, wd_format) triples.
"""

from __future__ import annotations

from pathlib import Path

# SaveFormat constants — locale-independent numbers (docs/pitfalls.md #1):
# wdFormatDocument = 0 (legacy .doc), wdFormatXMLDocument = 16 (.docx).
TO_FORMATS: dict[str, int] = {"doc": 0, "docx": 16}
SOURCE_EXTENSIONS = {".doc", ".docx"}


def normalize_to_format(to_format: str) -> str:
    """Accept 'docx', '.DOCX', ' doc ' → canonical key of TO_FORMATS."""
    fmt = (to_format or "").strip().lower().lstrip(".")
    if fmt not in TO_FORMATS:
        raise ValueError(f"Unsupported to_format: {to_format!r} (use 'docx' or 'doc')")
    return fmt


def resolve_target(source: Path, to_format: str, output_dir: Path | None = None) -> Path:
    """Same stem, new extension; relocated when output_dir is given.

    Never returns the source path itself: the extension always differs, so
    an in-place conversion can never clobber its own source.
    """
    target = source.with_suffix("." + to_format)
    if output_dir is not None:
        target = Path(output_dir) / target.name
    return target


def classify_source(extension: str, to_format: str) -> str:
    """One of 'convert' | 'already_target' | 'unsupported_source'."""
    ext = (extension or "").lower()
    if ext not in SOURCE_EXTENSIONS:
        return "unsupported_source"
    return "already_target" if ext == "." + to_format else "convert"
