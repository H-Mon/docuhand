"""Pure planning for merge_documents — output naming, zero COM.

One template + N data rows → N filled files.  Naming: a pattern with
``{i}`` (1-based, zero-padded to 3 for sort order) and any ``{field}``
tokens from the row data.  Everything that can reach a filename is
sanitized (path separators, reserved characters, length) — a hostile
data value must not be able to write outside the output directory.
"""

from __future__ import annotations

import re
from pathlib import Path

_ILLEGAL = '<>:"/\\|?*'
_COMPONENT_CAP = 40
_PAD = 3


def sanitize_component(value: str) -> str:
    """Make one naming component filesystem-safe and short."""
    text = str(value)
    out: list[str] = []
    for ch in text:
        if ch in _ILLEGAL or ord(ch) < 0x20:
            out.append("_")
        else:
            out.append(ch)
    cleaned = re.sub(r"\s+", " ", "".join(out)).strip().strip(".")
    return cleaned[:_COMPONENT_CAP]


def build_name(pattern: str, index: int, row: dict) -> str:
    """Substitute {i} and {field} tokens, then sanitize the whole stem.

    ``{i}`` renders zero-padded to 3 digits so directory listings sort in
    row order regardless of how the OS sorts numbers.
    """
    text = str(pattern or "")
    text = text.replace("{i}", str(index).zfill(_PAD))
    for key, value in row.items():
        text = text.replace("{" + str(key) + "}", sanitize_component(str(value)))
    # the pattern itself is untrusted input too
    return sanitize_component(text) or str(index).zfill(_PAD)


def resolve_collision(directory: Path, stem: str, extension: str) -> Path:
    """First free '<stem>.<ext>', then '<stem>-2', '-3', ... (cap 999)."""
    candidate = directory / f"{stem}{extension}"
    n = 2
    while candidate.exists() and n <= 999:
        candidate = directory / f"{stem}-{n}{extension}"
        n += 1
    return candidate


def default_output_dir(template: Path) -> Path:
    return template.parent / ("merged_" + template.stem)
