"""Pure planning for fill_template — zero COM, fully unit-testable.

Template fill modes (v0.1):
- ``content_controls``  map a control's Tag/Title to a data key
- ``bookmarks``         replace the bookmark's range with the value
- ``placeholders``      find/replace literal ``{{key}}`` tokens

``auto`` runs all three in that order, so one template can mix mechanisms.
The engine only ever receives validated (template, output, data, mode).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

# Wildcard Find pattern for leftover {{...}} scan (\{ and \} are escaped
# braces in Word wildcard syntax; * = any run of characters).
LEFTOVER_WILDCARD = r"\{\{*\}\}"

MAX_KEY_LEN = 100
MAX_VALUE_LEN = 200_000
OUTPUT_FORMATS = {".doc": 0, ".docx": 16}  # SaveFormat numbers (pitfall #1)

_MODES = ("auto", "content_controls", "bookmarks", "placeholders")


def normalize_mode(mode: str) -> str:
    m = (mode or "auto").strip().lower()
    if m not in _MODES:
        raise ValueError(
            f"Unsupported mode: {mode!r} (use 'auto', 'placeholders', 'bookmarks' or 'content_controls')"
        )
    return m


def validate_data(data: Any) -> dict[str, str]:
    """Coerce to {str: str} and enforce key/value sanity rules.

    Keys must be 1..MAX_KEY_LEN chars without braces or newlines (they are
    spliced into a literal ``{{key}}`` search token and a Word Find argument,
    both of which break on those).  Values are stringified and length-capped.
    """
    if not isinstance(data, dict):
        raise ValueError("data must be a JSON object mapping keys to replacement text")
    clean: dict[str, str] = {}
    for key, value in data.items():
        k = str(key)
        if not (0 < len(k) <= MAX_KEY_LEN):
            raise ValueError(f"data key length must be 1..{MAX_KEY_LEN}: {k[:40]!r}")
        if any(ch in k for ch in "{}\r\n\x0b"):
            raise ValueError(
                f"data key must not contain braces or newlines: {k[:40]!r} "
                "(the key is used inside a literal {{key}} token)"
            )
        v = value if isinstance(value, str) else str(value)
        if len(v) > MAX_VALUE_LEN:
            raise ValueError(
                f"value for {k[:40]!r} exceeds {MAX_VALUE_LEN} chars; "
                "split the document work into multiple calls"
            )
        clean[k] = v
    return clean


def resolve_output(template: Path, output: str | None) -> Path:
    """Default output = '<stem>_filled' next to the template, same format.

    The template file itself is sacred — same-path output is rejected and
    only .doc/.docx targets are supported in v0.1.
    """
    t = template.expanduser()
    if output:
        out = Path(output).expanduser()
    else:
        out = t.with_name(t.stem + "_filled" + t.suffix.lower())
    ext = out.suffix.lower()
    if ext not in OUTPUT_FORMATS:
        raise ValueError(
            f"output_path extension must be .doc or .docx (got {ext or 'none'}); "
            "use export_pdf for PDF output"
        )
    try:
        same = out.exists() and out.samefile(t)
    except OSError:
        same = str(out.resolve()).lower() == str(t.resolve()).lower()
    if same:
        raise ValueError(
            f"output_path must differ from the template ({t}); "
            "fill_template never rewrites its own template"
        )
    return out


def parse_bookmark_names(bookmark_args: Any) -> list[str]:
    """Optional explicit bookmark list — None/omitted means 'all bookmarks'."""
    if bookmark_args is None:
        return []
    if not isinstance(bookmark_args, (list, tuple)):
        raise ValueError("bookmarks must be a list of bookmark names")
    names: list[str] = []
    for b in bookmark_args:
        s = str(b).strip()
        if not s:
            raise ValueError("bookmark names must be non-empty")
        names.append(s)
    return names
