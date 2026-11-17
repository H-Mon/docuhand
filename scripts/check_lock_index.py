#!/usr/bin/env python
"""Fail if uv.lock points at anything other than the canonical PyPI index.

Why this exists: a regional mirror configured in the contributor's
environment (UV_INDEX_URL, pip.conf, or an OS-level mirror) is written
into uv.lock by `uv lock` / `uv run`. The lock file is published with
the repository, so every other user would then resolve through that
mirror - including users in countries where it is slow, blocked, or
untrusted.

An environment variable takes priority over `[tool.uv] index-url` in
pyproject.toml, so configuration alone cannot prevent this. Run this
before committing a changed lock file:

    python scripts/check_lock_index.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

CANONICAL = "https://pypi.org/"
REGISTRY = re.compile(r'registry\s*=\s*"(?P<url>[^"]+)"')
LOCK = Path(__file__).resolve().parent.parent / "uv.lock"


def main() -> int:
    if not LOCK.exists():
        print(f"{LOCK} not found - nothing to check.")
        return 0

    offenders = sorted(
        {
            m.group("url")
            for m in (REGISTRY.search(line) for line in
                      LOCK.read_text(encoding="utf-8").splitlines())
            if m and not m.group("url").startswith(CANONICAL)
        }
    )

    if offenders:
        print("uv.lock points at a non-canonical index:", file=sys.stderr)
        for url in offenders:
            print(f"  {url}", file=sys.stderr)
        print(
            "\nFix it with:\n"
            "  env -u UV_INDEX_URL uv lock\n"
            "or, on Windows:\n"
            "  Remove-Item Env:UV_INDEX_URL; uv lock",
            file=sys.stderr,
        )
        return 1

    print("uv.lock points at the canonical index.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
