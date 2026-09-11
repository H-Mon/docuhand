"""inspect_document — pre-flight health check for Word documents."""

from __future__ import annotations

import time
from typing import Any

from ..engine import get_engine
from ..errors import DocuhandError, InternalError
from ..safety import allowlist, audit

TOOL_NAME = "inspect_document"


def run(path: str) -> dict[str, Any]:
    t0 = time.perf_counter()

    try:
        allowlist.check_allowed(path)
    except DocuhandError as exc:
        audit.record(TOOL_NAME, path, ok=False, duration_ms=0, error_code=exc.error_code)
        return exc.to_dict()

    try:
        result = get_engine().inspect_document(path)
    except DocuhandError as exc:
        audit.record(
            TOOL_NAME,
            path,
            ok=False,
            duration_ms=(time.perf_counter() - t0) * 1000,
            error_code=exc.error_code,
        )
        return exc.to_dict()
    except Exception as exc:  # unknown engine crash — still structured
        err = InternalError(
            f"Unexpected error: {type(exc).__name__}: {exc}",
            llm_hint="Retry once. If it persists, report error_code and details to the DocuHand repository.",
            details={"path": path},
        )
        audit.record(
            TOOL_NAME,
            path,
            ok=False,
            duration_ms=(time.perf_counter() - t0) * 1000,
            error_code=err.error_code,
        )
        return err.to_dict()

    audit.record(
        TOOL_NAME,
        path,
        ok=True,
        duration_ms=(time.perf_counter() - t0) * 1000,
        engine=result.get("engine"),
    )
    return result
