"""extract_content — text / tables / outline out of legacy .doc as JSON."""

from __future__ import annotations

import time
from typing import Any

from ..engine import get_engine, get_worker
from ..engine.extraction import extract_content as engine_extract
from ..errors import DocuhandError, InvalidParamError, InternalError
from ..safety import allowlist, audit

TOOL_NAME = "extract_content"

_WHAT = ("text", "tables", "outline")


def run(path: str, what: str = "text") -> dict[str, Any]:
    t0 = time.perf_counter()

    what_norm = (what or "text").strip().lower()
    if what_norm not in _WHAT:
        return InvalidParamError(
            f"what must be one of {list(_WHAT)} (got {what!r})",
            llm_hint="'tables' pulls every table as a JSON grid; 'outline' returns the heading tree; 'text' is the full plain text.",
        ).to_dict()

    try:
        allowlist.check_allowed(path)
    except DocuhandError as exc:
        audit.record(TOOL_NAME, path, ok=False, duration_ms=0, error_code=exc.error_code)
        return exc.to_dict()

    worker = get_worker()
    engine = get_engine()
    try:
        result = worker.submit(engine_extract, engine, worker, path, what_norm)
    except DocuhandError as exc:
        audit.record(
            TOOL_NAME,
            path,
            ok=False,
            duration_ms=(time.perf_counter() - t0) * 1000,
            error_code=exc.error_code,
        )
        return exc.to_dict()
    except Exception as exc:
        err = InternalError(
            f"Unexpected error extracting content: {type(exc).__name__}: {exc}",
            llm_hint="Retry once. If it persists, report error_code and details to the DocuHand repository.",
            details={"path": path},
        )
        audit.record(TOOL_NAME, path, ok=False, duration_ms=(time.perf_counter() - t0) * 1000, error_code=err.error_code)
        return err.to_dict()

    audit.record(
        TOOL_NAME,
        path,
        ok=True,
        duration_ms=result.get("duration_ms") or 0,
        engine=result.get("engine"),
    )
    return result
