"""Content extraction: text / tables / outline from legacy .doc via COM.

Tables come out as row-major string matrices with merged-cell repeats
(Horizontal/VerticalMerge cells repeat their origin text, which keeps the
grid rectangular — documented, lossless for data, and simple for JSON).
Runs on the STA worker with the same failover contract.
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

from ..errors import (
    EngineUnavailableError,
    FileNotFoundError,
    OpenFailedError,
    PasswordProtectedError,
)
from .com_utils import _g, _short_exc
from .container_guard import detect_encryption, validate_container
from .wd_constants import WD_DO_NOT_SAVE_CHANGES


def _clean_cell(text: str | None) -> str:
    """Cell text arrives with trailing \\r\\a markers — strip control noise."""
    if not text:
        return ""
    return text.replace("\x07", "").replace("\r", " ").replace("\n", " ").strip()


def _extract_tables(doc: Any, max_tables: int = 100) -> list[dict[str, Any]]:
    tables: list[dict[str, Any]] = []
    try:
        count = int(_g(lambda: doc.Tables.Count, 0) or 0)
    except Exception:
        count = 0
    for ti in range(min(count, max_tables)):
        table = _g(lambda ti=ti: doc.Tables.Item(ti + 1))
        if table is None:
            continue
        rows_n = int(_g(lambda table=table: table.Rows.Count, 0) or 0)
        cols_n = int(_g(lambda table=table: table.Columns.Count, 0) or 0)
        grid: list[list[str]] = []
        # Cell-by-cell read: row-major, tolerant of vertical merges (which
        # break row.Range.Cells iteration ordering on some engines).
        for r in range(1, rows_n + 1):
            row_cells: list[str] = []
            for c in range(1, cols_n + 1):
                # Cell.Range.Text — the portable path. Pitfall #15: on
                # WPS-as-Word 12 Cell.Text itself does not exist (AttributeError)
                # and Range indirection is required; the marker strip below
                # cleans the \r\x07 cell terminator that comes with it.
                cell_text = _g(lambda table=table, r=r, c=c: table.Cell(r, c).Range.Text, "")
                row_cells.append(_clean_cell(cell_text))
            grid.append(row_cells)
        tables.append(
            {
                "index": ti + 1,
                "rows": rows_n,
                "columns": cols_n,
                "data": grid,
            }
        )
    return tables


def _extract_outline(doc: Any, max_items: int = 200) -> list[dict[str, Any]]:
    """Headings via paragraph outline levels (locale-independent numbers)."""
    items: list[dict[str, Any]] = []
    try:
        paras = doc.Paragraphs
        total = int(_g(lambda: paras.Count, 0) or 0)
    except Exception:
        total = 0
    for i in range(min(total, 2000)):  # hard scan cap
        para = _g(lambda i=i: paras.Item(i + 1))
        if para is None:
            continue
        level = _g(lambda para=para: para.OutlineLevel, 10)
        if not isinstance(level, int) or level < 1 or level > 9:
            continue  # body text is level 10
        text = _g(lambda para=para: para.Range.Text, "")
        text = (text or "").replace("\r", "").replace("\x07", "").strip()
        if not text:
            continue
        items.append({"level": level, "text": text[:200]})
        if len(items) >= max_items:
            break
    return items


def extract_content_from_doc(doc: Any, what: str) -> dict[str, Any]:
    """Run the extraction selectors on an already-open document."""
    payload: dict[str, Any] = {"what": what}
    if what in ("text", "all"):
        text = _g(lambda: doc.Content.Text, "") or ""
        payload["text"] = text
        payload["text_chars"] = len(text)
    if what in ("tables", "all"):
        tables = _extract_tables(doc)
        payload["tables"] = tables
        payload["tables_count"] = len(tables)
    if what in ("outline", "all"):
        items = _extract_outline(doc)
        payload["outline"] = items
        payload["outline_count"] = len(items)
    return payload


def extract_content(
    engine: Any,
    worker: Any,
    path: str,
    what: str,
) -> dict[str, Any]:
    """Full tool flow with failover: probe order, open RO, extract, close.

    ``engine``/``worker`` come from the caller (tools layer) so this module
    stays import-cycle-free and testable with fakes.
    """
    t0 = time.perf_counter()
    src = Path(path).expanduser().resolve()
    if not src.exists():
        raise FileNotFoundError(str(src))
    ok, reason = validate_container(str(src))
    if not ok:
        raise OpenFailedError(str(src), "container pre-check", f"structurally invalid container: {reason}")
    if detect_encryption(str(src)):
        raise PasswordProtectedError(str(src), "container pre-check")

    order = engine._engine_order()
    attempts: list[dict[str, str]] = []
    launch_failures = 0

    for engine_name, progid in order:
        try:
            app = engine._ensure_app(progid, engine_name)
        except Exception as exc:
            launch_failures += 1
            attempts.append({"engine": engine_name, "stage": "launch", "error": _short_exc(exc)})
            engine._discard_app()
            continue
        try:
            doc = engine._open_read_only(app, str(src))
            try:
                payload = extract_content_from_doc(doc, what)
            finally:
                try:
                    doc.Close(WD_DO_NOT_SAVE_CHANGES)
                except Exception:
                    pass
        except PasswordProtectedError:
            raise
        except Exception as exc:
            attempts.append({"engine": engine_name, "stage": "extract", "error": _short_exc(exc)})
            engine._discard_app()
            continue

        payload.update(
            {
                "ok": True,
                "tool": "extract_content",
                "path": str(src),
                "engine": engine_name,
                "engine_version": engine._engine_version,
                "duration_ms": round((time.perf_counter() - t0) * 1000),
            }
        )
        return payload

    if launch_failures == len(order):
        raise EngineUnavailableError(attempts)
    raise OpenFailedError(str(src), "word→wps failover exhausted", "; ".join(a["error"] for a in attempts))
