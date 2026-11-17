"""Path allowlist — env-driven, enforced before any tool runs.

``DOCUHAND_ALLOWLIST``: os.pathsep-separated absolute directories.
Empty/absent (the v0.1 default) = allow everything: the user launches the
server themselves, which per spec IS the explicit consent. Setting the env
var locks the server to those subtrees.
"""

from __future__ import annotations

import os
from pathlib import Path

from ..errors import PathNotAllowedError


def get_allowlist() -> list[Path]:
    raw = os.environ.get("DOCUHAND_ALLOWLIST", "")
    entries: list[Path] = []
    for part in raw.split(os.pathsep):
        if part.strip():
            entries.append(Path(part.strip()).expanduser().resolve())
    return entries


def check_allowed(path: str) -> None:
    entries = get_allowlist()
    if not entries:
        return
    p = Path(path).expanduser()
    try:
        resolved = p.resolve()
    except OSError:
        resolved = p.absolute()
    for root in entries:
        try:
            resolved.relative_to(root)
            return
        except ValueError:
            continue
    raise PathNotAllowedError(str(p), [str(e) for e in entries])
