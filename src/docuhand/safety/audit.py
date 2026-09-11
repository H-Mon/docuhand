"""JSONL audit log — every operation lands here, user can disable it.

Set ``DOCUHAND_AUDIT=off`` to disable, ``DOCUHAND_AUDIT_PATH`` to relocate
(default: %LOCALAPPDATA%/docuhand/audit.jsonl). Audit failures are swallowed:
logging must never break a document operation.
"""

from __future__ import annotations

import json
import os
import time
from pathlib import Path

_ENABLED = os.environ.get("DOCUHAND_AUDIT", "on").strip().lower() not in ("off", "0", "false", "no")


def _log_path() -> Path:
    env = os.environ.get("DOCUHAND_AUDIT_PATH")
    if env:
        return Path(env)
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    return Path(base) / "docuhand" / "audit.jsonl"


def record(
    tool: str,
    path: str,
    *,
    ok: bool,
    duration_ms: float,
    engine: str | None = None,
    error_code: str | None = None,
) -> None:
    if not _ENABLED:
        return
    entry = {
        "ts": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "tool": tool,
        "path": path,
        "ok": ok,
        "duration_ms": round(duration_ms),
        "engine": engine,
        "error_code": error_code,
    }
    try:
        target = _log_path()
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass
