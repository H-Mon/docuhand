"""export_pdf — print-fidelity PDF via the real Word/WPS layout engine.

Tool layer: parse the page range with pure helpers (engine/pdf_plan.py),
gate through allowlist + audit (export creates a NEW file but never
touches the source, so it stays available in read-only mode), then run
one engine export on the COM STA worker.
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

from ..engine import get_engine, get_worker
from ..engine.pdf_plan import parse_page_range
from ..errors import DocuhandError, FileNotFoundError, InternalError, InvalidParamError
from ..safety import allowlist, audit

TOOL_NAME = "export_pdf"


def run(
    path: str,
    output_path: str | None = None,
    page_range: str | None = None,
) -> dict[str, Any]:
    t0 = time.perf_counter()

    # -- argument validation ---------------------------------------------------
    src = Path(path).expanduser()
    if not src.exists():
        return FileNotFoundError(str(src)).to_dict()

    try:
        pages = parse_page_range(page_range)
    except ValueError as exc:
        return InvalidParamError(str(exc), llm_hint="Use page_range='all', '7' or '3-9', or omit it.").to_dict()

    if output_path:
        dst = Path(output_path).expanduser()
        if dst.suffix.lower() != ".pdf":
            return InvalidParamError(
                f"output_path must end in .pdf (got {dst.suffix or 'none'})",
                llm_hint="PDF export writes a .pdf file; pick a name like 'report.pdf'.",
            ).to_dict()
    else:
        dst = src.with_suffix(".pdf")

    try:
        if dst.exists() and dst.samefile(src):
            return InvalidParamError(
                "output_path must differ from the source document",
                llm_hint="Pick another name or omit output_path to write next to the source.",
            ).to_dict()
    except OSError:
        pass

    # -- safety gates ------------------------------------------------------------
    try:
        allowlist.check_allowed(str(src))
        allowlist.check_allowed(str(dst))
    except DocuhandError as exc:
        audit.record(TOOL_NAME, str(src), ok=False, duration_ms=0, error_code=exc.error_code)
        return exc.to_dict()

    # -- execute -------------------------------------------------------------------
    worker = get_worker()
    engine = get_engine()
    page_from, page_to = pages if pages else (None, None)
    try:
        result = worker.submit(engine.export_pdf, str(src), str(dst), page_from, page_to)
    except DocuhandError as exc:
        audit.record(
            TOOL_NAME,
            str(src),
            ok=False,
            duration_ms=(time.perf_counter() - t0) * 1000,
            error_code=exc.error_code,
        )
        return exc.to_dict()
    except Exception as exc:  # unknown engine crash — still structured
        err = InternalError(
            f"Unexpected error exporting {src.name}: {type(exc).__name__}: {exc}",
            llm_hint="Retry once. If it persists, report error_code and details to the DocuHand repository.",
            details={"source": str(src)},
        )
        audit.record(TOOL_NAME, str(src), ok=False, duration_ms=(time.perf_counter() - t0) * 1000, error_code=err.error_code)
        return err.to_dict()

    audit.record(TOOL_NAME, str(src), ok=True, duration_ms=result.get("duration_ms") or 0, engine=result.get("engine"))
    return result
