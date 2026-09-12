"""Pure planning for export_pdf — page-range grammar, zero COM.

v0.1 grammar: ``"all"`` / ``None`` (whole document) or a single contiguous
1-based range ``"7"`` / ``"3-9"``.  Word's ExportAsFixedFormat only accepts
one From/To pair anyway; comma lists would silently lie about fidelity.
"""

from __future__ import annotations

from typing import Any

WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_ALL_DOCUMENT = 0
WD_EXPORT_FROM_TO = 3


def parse_page_range(page_range: Any) -> tuple[int, int] | None:
    """``"3-9"`` → (3, 9); ``"7"`` → (7, 7); ``None``/``"all"`` → None.

    Raises ValueError with an llm-oriented message on anything else.
    """
    if page_range is None:
        return None
    if not isinstance(page_range, str):
        raise ValueError(f"page_range must be a string like 'all', '7' or '3-9' (got {type(page_range).__name__})")
    text = page_range.strip().lower()
    if text in ("", "all", "*"):
        return None
    parts = text.replace("–", "-").replace("—", "-").split("-")
    if len(parts) == 1:
        try:
            page = int(parts[0])
        except ValueError:
            raise ValueError(f"page_range must be 'all', a page number, or 'from-to' (got {page_range!r})") from None
        if page < 1:
            raise ValueError(f"page numbers are 1-based (got {page_range!r})")
        return (page, page)
    if len(parts) == 2:
        try:
            lo, hi = int(parts[0]), int(parts[1])
        except ValueError:
            raise ValueError(f"page_range must be 'all', a page number, or 'from-to' (got {page_range!r})") from None
        if lo < 1 or hi < lo:
            raise ValueError(f"page_range must satisfy 1 <= from <= to (got {page_range!r})")
        return (lo, hi)
    raise ValueError(
        f"only a single contiguous page range is supported in v0.1 (got {page_range!r}); "
        "call export_pdf once per range"
    )
