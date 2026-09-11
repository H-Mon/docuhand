"""convert_documents — batch .doc ↔ .docx through the real engine.

Tool-layer responsibilities (the engine stays MCP-free):
- validate + normalize arguments into structured errors
- plan the batch with pure helpers (engine/convert_plan.py)
- run each conversion on the COM STA worker (threading contract)
- enforce the overwrite policy BEFORE any engine call
- keep every source file (never delete) and report that explicitly
- wrap every failure into a per-file structured entry — one bad file in a
  500-file batch must not kill the batch. Only a total engine launch
  failure aborts the whole run (nothing else can succeed either).
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

from ..engine import convert_plan, get_engine, get_worker
from ..engine.office_app import engine_order_names, sniff_format
from ..errors import (
    DocuhandError,
    EngineUnavailableError,
    FileNotFoundError,
    InternalError,
    InvalidParamError,
)
from ..safety import allowlist, audit

TOOL_NAME = "convert_documents"


def run(
    directory: str,
    to_format: str = "docx",
    recursive: bool = False,
    overwrite: str = "skip",
    output_dir: str | None = None,
) -> dict[str, Any]:
    t0 = time.perf_counter()

    # -- argument validation -------------------------------------------------
    try:
        fmt = convert_plan.normalize_to_format(to_format)
    except ValueError as exc:
        return InvalidParamError(
            str(exc),
            llm_hint="Use to_format='docx' (recommended, modern format) or 'doc'.",
        ).to_dict()

    if overwrite not in ("skip", "overwrite"):
        return InvalidParamError(
            f"overwrite must be 'skip' or 'overwrite', got {overwrite!r}",
            llm_hint="Default 'skip' never destroys an existing file; "
            "'overwrite' replaces existing target files deliberately.",
        ).to_dict()

    root = Path(directory).expanduser()
    if not root.exists():
        return FileNotFoundError(str(root)).to_dict()
    if not root.is_dir():
        return InvalidParamError(
            f"Not a directory: {root}",
            llm_hint="Pass the directory that contains the documents. "
            "For a single file, put it in a directory or use inspect_document first.",
        ).to_dict()
    root = root.resolve()  # canonical form in payloads; the engine needs absolutes

    out_root: Path | None = None
    if output_dir:
        out_root = Path(output_dir).expanduser()
        try:
            out_root.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            return DocuhandError(
                f"Cannot create output directory {out_root}: {exc}",
                llm_hint="Check the path (drive letter, permissions). "
                "Omit output_dir to convert in place, next to each source.",
                details={"output_dir": str(out_root)},
            ).to_dict()

    # -- safety gates ----------------------------------------------------------
    try:
        allowlist.check_allowed(str(root))
        if out_root is not None:
            allowlist.check_allowed(str(out_root))
    except DocuhandError as exc:
        audit.record(TOOL_NAME, str(root), ok=False, duration_ms=0, error_code=exc.error_code)
        return exc.to_dict()

    wd_format = convert_plan.TO_FORMATS[fmt]

    # -- plan ------------------------------------------------------------------
    pattern = root.rglob("*") if recursive else root.glob("*")
    try:
        candidates = sorted(p for p in pattern if p.is_file())
    except OSError as exc:
        return InternalError(
            f"Cannot list directory: {exc}",
            details={"directory": str(root)},
        ).to_dict()

    results: list[dict[str, Any]] = []
    counts = {
        "to_convert": 0,
        "already_target": 0,
        "skipped_existing": 0,
        "converted": 0,
        "failed": 0,
    }
    ignored = 0

    planned: list[tuple[Path, Path]] = []
    for src in candidates:
        kind = convert_plan.classify_source(src.suffix, fmt)
        if kind == "unsupported_source":
            ignored += 1
            continue
        if kind == "already_target":
            counts["already_target"] += 1
            results.append({"source": str(src), "target": None, "status": "already_target"})
            continue
        counts["to_convert"] += 1
        planned.append((src, convert_plan.resolve_target(src, fmt, out_root)))

    # -- execute -----------------------------------------------------------------
    worker = get_worker()
    engine = get_engine()
    aborted: dict[str, Any] | None = None

    for src, target in planned:
        if target.exists() and overwrite == "skip":
            counts["skipped_existing"] += 1
            results.append(
                {
                    "source": str(src),
                    "target": str(target),
                    "status": "skipped_existing",
                    "reason": "target exists (overwrite='skip')",
                }
            )
            continue

        try:
            info = worker.submit(engine.convert_document, str(src), str(target), wd_format)
        except EngineUnavailableError as exc:
            counts["failed"] += 1
            results.append(
                {"source": str(src), "target": str(target), "status": "failed", "error": exc.to_dict()["error"]}
            )
            audit.record(TOOL_NAME, str(src), ok=False, duration_ms=0, error_code=exc.error_code)
            aborted = exc.to_dict()
            break
        except DocuhandError as exc:
            counts["failed"] += 1
            results.append(
                {"source": str(src), "target": str(target), "status": "failed", "error": exc.to_dict()["error"]}
            )
            audit.record(TOOL_NAME, str(src), ok=False, duration_ms=0, error_code=exc.error_code)
            continue
        except Exception as exc:  # unknown engine crash — still structured
            err = InternalError(
                f"Unexpected error converting {src.name}: {type(exc).__name__}: {exc}",
                llm_hint="Retry once. If it persists, report error_code and details to the DocuHand repository.",
                details={"source": str(src)},
            )
            counts["failed"] += 1
            results.append(
                {"source": str(src), "target": str(target), "status": "failed", "error": err.to_dict()["error"]}
            )
            audit.record(TOOL_NAME, str(src), ok=False, duration_ms=0, error_code=err.error_code)
            continue

        # Verify the artifact really is the claimed container format.
        try:
            with open(target, "rb") as f:
                verified = sniff_format(f.read(8))
        except OSError:
            verified = "unreadable"

        counts["converted"] += 1
        results.append(
            {
                "source": str(src),
                "target": str(target),
                "status": "converted",
                "engine": info.get("engine"),
                "target_verified_format": verified,
                "target_size_bytes": target.stat().st_size if target.exists() else None,
                "duration_ms": info.get("duration_ms"),
            }
        )
        audit.record(
            TOOL_NAME,
            str(src),
            ok=True,
            duration_ms=info.get("duration_ms") or 0,
            engine=info.get("engine"),
        )

    total_found = counts["already_target"] + counts["to_convert"]
    payload: dict[str, Any] = {
        "ok": aborted is None,
        "tool": TOOL_NAME,
        "directory": str(root),
        "output_dir": str(out_root) if out_root else None,
        "to_format": fmt,
        "recursive": recursive,
        "overwrite": overwrite,
        "engine_order": engine_order_names(),
        "sources_deleted": False,
        "summary": {
            **counts,
            "doc_files_found": total_found,
            "ignored_other_extensions": ignored,
            "duration_ms": round((time.perf_counter() - t0) * 1000),
        },
        "results": results,
    }
    if aborted is not None:
        payload["error"] = aborted["error"]
        payload["llm_hint"] = (
            "No Office engine could be launched, so the batch aborted before "
            "touching the remaining files. Install/repair Microsoft Word or "
            "WPS Office (or close a stuck first-run dialog), then retry."
        )
    return payload
