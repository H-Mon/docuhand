"""fill_template — fill bookmarks / {{placeholders}} / content controls.

Tool layer: validate + plan with pure helpers (engine/template_plan.py),
gate through the safety stack (read-only policy, allowlist, audit), then
run one engine fill on the COM STA worker. The template file itself is
never modified; failures come back as structured errors.
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

from ..engine import get_engine, get_worker
from ..engine.template_plan import (
    normalize_mode,
    parse_bookmark_names,
    resolve_output,
    validate_data,
)
from ..errors import DocuhandError, FileNotFoundError, InternalError, InvalidParamError, TemplateFillError
from ..safety import allowlist, audit
from ..safety.policy import ensure_writable

TOOL_NAME = "fill_template"


def run(
    template_path: str,
    data: dict[str, Any],
    output_path: str | None = None,
    mode: str = "auto",
    bookmarks: list[str] | None = None,
) -> dict[str, Any]:
    t0 = time.perf_counter()

    # -- argument validation ---------------------------------------------------
    try:
        norm_mode = normalize_mode(mode)
        clean_data = validate_data(data)
        names = parse_bookmark_names(bookmarks)
    except ValueError as exc:
        return InvalidParamError(str(exc), llm_hint="Fix the argument and retry.").to_dict()

    template = Path(template_path).expanduser()
    if not template.exists():
        return FileNotFoundError(str(template)).to_dict()

    try:
        output = resolve_output(template, output_path)
    except ValueError as exc:
        return InvalidParamError(
            str(exc),
            llm_hint="Choose a different output_path (extension .doc/.docx); the template is never rewritten.",
        ).to_dict()

    # -- safety gates ------------------------------------------------------------
    try:
        ensure_writable(TOOL_NAME)
        allowlist.check_allowed(str(template))
        allowlist.check_allowed(str(output))
    except DocuhandError as exc:
        audit.record(TOOL_NAME, str(template), ok=False, duration_ms=0, error_code=exc.error_code)
        return exc.to_dict()

    # -- execute -------------------------------------------------------------------
    worker = get_worker()
    engine = get_engine()
    try:
        result = worker.submit(
            engine.fill_document,
            str(template),
            str(output),
            clean_data,
            norm_mode,
            names,
        )
    except DocuhandError as exc:
        audit.record(
            TOOL_NAME,
            str(template),
            ok=False,
            duration_ms=(time.perf_counter() - t0) * 1000,
            error_code=exc.error_code,
        )
        return exc.to_dict()
    except Exception as exc:  # unknown engine crash — still structured
        err = InternalError(
            f"Unexpected error filling {template.name}: {type(exc).__name__}: {exc}",
            llm_hint="Retry once. If it persists, report error_code and details to the DocuHand repository.",
            details={"template": str(template)},
        )
        audit.record(TOOL_NAME, str(template), ok=False, duration_ms=(time.perf_counter() - t0) * 1000, error_code=err.error_code)
        return err.to_dict()

    audit.record(TOOL_NAME, str(template), ok=True, duration_ms=result.get("duration_ms") or 0, engine=result.get("engine"))

    # A fill that silently skipped keys is a warning, not a crash — surface it.
    problems: list[str] = []
    if result.get("missing_keys"):
        problems.append(f"{len(result['missing_keys'])} data key(s) found no target in the template")
    if result.get("leftover_placeholders"):
        problems.append(f"{len(result['leftover_placeholders'])} {{{{...}}}} placeholder(s) remain unfilled")
    if problems and norm_mode == "auto":
        incomplete = TemplateFillError(
            str(output),
            "; ".join(problems),
            missing=result.get("missing_keys") or None,
        )
        # success-shaped payload plus the structured warning — the file IS
        # written; the agent decides whether it is good enough to keep.
        result["ok"] = True
        result["warning"] = incomplete.to_dict()["error"]
    return result
