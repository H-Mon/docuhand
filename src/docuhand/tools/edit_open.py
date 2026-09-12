"""edit_open_document — edit a document the user has open RIGHT NOW.

Attaches to the live Word/WPS instance on the STA worker, resolves the
document by path or window title, validates the whole operation batch up
front (edit_plan.py), applies it, and reports the saved/dirty state.
Default is NOT to save ('ask') — the agent tells the user what changed.
"""

from __future__ import annotations

import time
from typing import Any

from ..engine import edit_plan, get_worker
from ..engine.live_edit import edit_open_document as engine_edit
from ..errors import DocuhandError, InvalidParamError, InternalError
from ..safety import allowlist, audit
from ..safety.policy import ensure_writable

TOOL_NAME = "edit_open_document"


def run(
    path_or_title: str,
    operations: list[dict[str, Any]],
    save: str = "ask",
) -> dict[str, Any]:
    t0 = time.perf_counter()

    # -- validate the whole batch BEFORE any COM call --------------------------
    try:
        clean_ops = edit_plan.validate_operations(operations)
    except ValueError as exc:
        return InvalidParamError(str(exc), llm_hint="Fix the operations list and retry.").to_dict()
    if save not in ("ask", "save"):
        return InvalidParamError(
            f"save must be 'ask' or 'save' (got {save!r})",
            llm_hint="Default 'ask' never writes without the user knowing; 'save' persists the changes.",
        ).to_dict()

    # -- safety gates ------------------------------------------------------------
    # allowlist applies to the file path when resolvable; title matches are
    # allowed only because the user opened that document themselves.
    p = None
    from pathlib import Path

    try:
        p = Path(path_or_title).expanduser()
        if p.exists():
            allowlist.check_allowed(str(p.resolve()))
    except DocuhandError as exc:
        audit.record(TOOL_NAME, str(path_or_title), ok=False, duration_ms=0, error_code=exc.error_code)
        return exc.to_dict()
    except OSError:
        pass  # a window title, not a path — the user opened it themselves

    try:
        ensure_writable(TOOL_NAME)
    except DocuhandError as exc:
        audit.record(TOOL_NAME, str(path_or_title), ok=False, duration_ms=0, error_code=exc.error_code)
        return exc.to_dict()

    # -- execute on the STA worker -------------------------------------------------
    worker = get_worker()
    try:
        result = worker.submit(engine_edit, path_or_title, clean_ops, save)
    except DocuhandError as exc:
        audit.record(
            TOOL_NAME,
            str(path_or_title),
            ok=False,
            duration_ms=(time.perf_counter() - t0) * 1000,
            error_code=exc.error_code,
        )
        return exc.to_dict()
    except Exception as exc:
        err = InternalError(
            f"Unexpected error editing open document: {type(exc).__name__}: {exc}",
            llm_hint="Retry once. If it persists, report error_code and details to the DocuHand repository.",
            details={"path_or_title": path_or_title},
        )
        audit.record(TOOL_NAME, str(path_or_title), ok=False, duration_ms=(time.perf_counter() - t0) * 1000, error_code=err.error_code)
        return err.to_dict()

    audit.record(
        TOOL_NAME,
        result.get("document", str(path_or_title)),
        ok=True,
        duration_ms=result.get("duration_ms") or 0,
        engine=result.get("engine"),
    )
    return result
