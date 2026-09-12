"""Global write policy — the ``--read-only`` mode as data, not discipline.

Set ``DOCUHAND_READONLY=1`` to lock the server into read-only operation.
Blocked: document-mutating tools (edit_open_document, fill_template,
merge_documents, convert_documents).  Allowed: pure readers (inspect_document,
extract_content) and new-artifact exporters (export_pdf never touches the
source document).
"""

from __future__ import annotations

from ..errors import ReadOnlyError

_TRUTHY = {"1", "true", "yes", "on"}


def is_readonly() -> bool:
    raw = ""
    try:
        raw = __import__("os").environ.get("DOCUHAND_READONLY", "")
    except Exception:
        return False
    return raw.strip().lower() in _TRUTHY


def ensure_writable(tool: str) -> None:
    """Raise ReadOnlyError when the server runs in read-only mode."""
    if is_readonly():
        raise ReadOnlyError(tool)
