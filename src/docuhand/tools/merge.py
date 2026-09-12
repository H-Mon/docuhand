"""merge_documents — one template × N data rows → N filled files.

Mail-merge style batch generation - one template, N data rows, N filled
rebuilt as a product tool). Per-row fault isolation: one bad row fails
alone and the batch continues — the convert_documents survival guarantee,
applied to generation.
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

from ..engine import get_engine, get_worker
from ..engine.merge_plan import build_name, default_output_dir, resolve_collision
from ..engine.template_plan import validate_data
from ..errors import DocuhandError, InternalError, InvalidParamError
from ..safety import allowlist, audit
from ..safety.policy import ensure_writable

TOOL_NAME = "merge_documents"


def run(
    template_path: str,
    data_rows: list[dict[str, Any]],
    output_dir: str | None = None,
    name_pattern: str = "{i}_{name}",
) -> dict[str, Any]:
    t0 = time.perf_counter()

    # -- argument validation ---------------------------------------------------
    template = Path(template_path).expanduser()
    if not template.exists():
        from ..errors import FileNotFoundError

        return FileNotFoundError(str(template)).to_dict()

    if not isinstance(data_rows, list) or not data_rows:
        return InvalidParamError(
            "data_rows must be a non-empty array of row objects",
            llm_hint="Each row is one output document: [{\"name\": \"张三\", ...}, ...].",
        ).to_dict()
    if len(data_rows) > 500:
        return InvalidParamError(
            f"at most 500 rows per call (got {len(data_rows)})",
            llm_hint="Split very large merges into several calls; the COM session stays healthier.",
        ).to_dict()

    cleaned_rows: list[dict[str, str]] = []
    for i, row in enumerate(data_rows):
        try:
            cleaned_rows.append(validate_data(row))
        except ValueError as exc:
            return InvalidParamError(
                f"data_rows[{i}]: {exc}",
                llm_hint="Fix this row and retry.",
            ).to_dict()

    out_dir = Path(output_dir).expanduser() if output_dir else default_output_dir(template)
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return DocuhandError(
            f"Cannot create output directory {out_dir}: {exc}",
            llm_hint="Check the path (drive letter, permissions) or pass a different output_dir.",
            details={"output_dir": str(out_dir)},
        ).to_dict()

    ext = template.suffix.lower()
    if ext not in (".doc", ".docx"):
        return InvalidParamError(
            f"template must be .doc or .docx (got {ext or 'none'})",
            llm_hint="Pick a Word template file; PDF templates are not supported.",
        ).to_dict()

    # -- safety gates ------------------------------------------------------------
    try:
        ensure_writable(TOOL_NAME)
        allowlist.check_allowed(str(template))
        allowlist.check_allowed(str(out_dir))
    except DocuhandError as exc:
        audit.record(TOOL_NAME, str(template), ok=False, duration_ms=0, error_code=exc.error_code)
        return exc.to_dict()

    # -- plan outputs (collisions resolved before any engine work) -----------------
    planned: list[tuple[dict[str, str], Path]] = []
    for i, row in enumerate(cleaned_rows, start=1):
        stem = build_name(name_pattern, i, row)
        target = resolve_collision(out_dir, stem, ext)
        planned.append((row, target))

    # -- execute per row with fault isolation ----------------------------------------
    worker = get_worker()
    engine = get_engine()
    results: list[dict[str, Any]] = []
    counts = {"merged": 0, "failed": 0, "with_warnings": 0}

    for i, (row, target) in enumerate(planned, start=1):
        try:
            r = worker.submit(engine.merge_one, str(template), str(target), row)
        except DocuhandError as exc:
            counts["failed"] += 1
            results.append({"row": i, "output": str(target), "status": "failed", "error": exc.to_dict()["error"]})
            audit.record(TOOL_NAME, str(template), ok=False, duration_ms=0, error_code=exc.error_code)
            continue
        except Exception as exc:
            err = InternalError(
                f"row {i}: {type(exc).__name__}: {exc}",
                llm_hint="Retry once; if it persists report error_code to the DocuHand repository.",
            )
            counts["failed"] += 1
            results.append({"row": i, "output": str(target), "status": "failed", "error": err.to_dict()["error"]})
            audit.record(TOOL_NAME, str(template), ok=False, duration_ms=0, error_code=err.error_code)
            continue

        counts["merged"] += 1
        entry: dict[str, Any] = {
            "row": i,
            "output": str(target),
            "status": "merged",
            "engine": r.get("engine"),
            "output_size_bytes": r.get("output_size_bytes"),
        }
        if r.get("row_warning"):
            counts["with_warnings"] += 1
            entry["warning"] = r["row_warning"]
        results.append(entry)
        audit.record(TOOL_NAME, str(template), ok=True, duration_ms=r.get("duration_ms") or 0, engine=r.get("engine"))

    return {
        "ok": counts["failed"] < len(planned),  # total engine failure only if every row failed
        "tool": TOOL_NAME,
        "template": str(template),
        "output_dir": str(out_dir),
        "summary": {
            **counts,
            "rows": len(planned),
            "duration_ms": round((time.perf_counter() - t0) * 1000),
        },
        "results": results,
    }
