"""Edit documents that are currently OPEN in a live Word/WPS instance.

The signature differentiator: every other engine method works on its own
hidden ``DispatchEx`` instance. This one ATTACHES to the instance the user
is running (``Dispatch`` — the running-object-table lookup — vs Launch)
and drives the visible document through the same STA worker.

Safety model:
- resolve_document matches by absolute path OR exact window title.
- The document must already exist in the attached instance's Documents
  collection; we never open files silently in the user's session.
- ``save`` default ``"ask"`` maps to: do NOT save, report dirty state and
  let the agent ask the user. ``"save"`` saves explicitly as the final op.
- Attach failure (no live instance) is a structured error, not a fallback
  to a hidden instance — silently editing a file the user cannot see
  would betray the "hands on the visible document" contract.
"""

from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

from ..errors import (
    DocuhandError,
    FileNotFoundError,
    InternalError,
)
from .com_utils import _g, _short_exc
from .templating import _restore_font, _snapshot_font

WD_DO_NOT_SAVE_CHANGES = 0
WD_SAVE_CHANGES = -1


def _rot_word_apps() -> list[Any]:
    """Every Word/KWPS instance registered in the Running Object Table.

    ``Dispatch("Word.Application")`` returns *one* registration — and after a
    crashed-or-quit engine leaves a stale zero-document instance behind
    (pitfall #2), that registration is frequently NOT the one holding the
    user's visible documents. Enumerating the ROT and keeping every
    candidate lets the caller pick the instance that actually has documents.
    """
    apps: list[Any] = []
    try:
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
        enum = rot.EnumRunning()
        while True:
            monikers = enum.Next(1)
            if not monikers:
                break
            for moniker in monikers:
                app = _app_from_moniker(rot, moniker)
                if app is not None:
                    apps.append(app)
    except Exception:
        # ROT enumeration is best-effort; the ProgID fallback below still works.
        pass
    return apps


def _app_from_moniker(rot: Any, moniker: Any) -> Any | None:
    """Bind a ROT moniker to a Word/WPS Application, or None if it isn't one.

    ``rot.GetObject(moniker)`` is the only pywin32 shape that works here:
    ``moniker.BindToObject`` rejects the CLSID moniker with
    ``E_INVALIDARG`` and ``PyIMoniker`` has no ``GetObject``.
    """
    try:
        unk = rot.GetObject(moniker)
        disp = unk.QueryInterface(pythoncom.IID_IDispatch)
        app = win32com.client.Dispatch(disp)
    except Exception:
        return None
    # WPS impersonates Word's ProgID, so the Name is not a reliable filter;
    # instead require that the object exposes a Documents collection.
    try:
        int(app.Documents.Count)
    except Exception:
        return None
    return app


def _attach_running_app() -> Any:
    """Attach to the user's running Word/WPS instance (ROT lookup).

    Prefers an instance that actually holds open documents: a stale
    zero-document registration must never shadow the live one (see
    docs/pitfalls.md #17).
    """
    pythoncom.CoInitialize()
    candidates = _rot_word_apps()
    # Prefer instances that already have documents open.
    with_docs = [a for a in candidates if int(_g(lambda a=a: a.Documents.Count, 0) or 0) > 0]
    if with_docs:
        return with_docs[0]
    if candidates:
        return candidates[0]
    # ROT enumeration found nothing usable — fall back to the ProgID lookup.
    try:
        return win32com.client.Dispatch("Word.Application")
    except Exception:
        pass
    try:
        return win32com.client.Dispatch("KWPS.Application")
    except Exception as exc:
        raise DocuhandError(
            "No running Word/WPS instance to attach to",
            llm_hint=(
                "edit_open_document works on documents the user has open "
                "right now. If the file is closed, use fill_template / "
                "convert_documents instead; if it should be open, ask the "
                "user to open it first."
            ),
            details={"com_error": _short_exc(exc)},
        ) from None


def _resolve_document(app: Any, path_or_title: str) -> tuple[Any, dict[str, Any]]:
    """Match by absolute path first, then exact/full window title."""
    docs = app.Documents
    count = int(_g(lambda: docs.Count, 0) or 0)
    wanted_path = str(Path(path_or_title).expanduser().resolve()).lower()
    wanted_title = path_or_title.strip().lower()

    candidates: list[dict[str, Any]] = []
    for i in range(count):
        doc = _g(lambda i=i: docs.Item(i + 1))
        if doc is None:
            continue
        full_name = _g(lambda doc=doc: doc.FullName, "") or ""
        name = _g(lambda doc=doc: doc.Name, "") or ""
        title = _g(lambda doc=doc: doc.ActiveWindow.Caption if doc.Windows.Count else name, "") or name
        candidates.append(
            {
                "doc": doc,
                "full_name": str(full_name),
                "name": str(name),
                "title": str(title),
                "saved": bool(_g(lambda doc=doc: doc.Saved, True)),
                "readonly": bool(_g(lambda doc=doc: doc.ReadOnly, False)),
            }
        )

    for c in candidates:
        if c["full_name"] and c["full_name"].lower() == wanted_path:
            return c["doc"], c
    for c in candidates:
        if wanted_title and wanted_title in (c["title"].lower(), c["name"].lower()):
            return c["doc"], c
    # one more chance: title minus hidden extension variants
    stem = Path(path_or_title).stem.lower()
    for c in candidates:
        if stem and stem == Path(c["name"]).stem.lower():
            return c["doc"], c

    raise FileNotFoundError(
        f"Document not found among the {count} open document(s): {path_or_title}"
    )


def _replace_all_two_step(doc: Any, token: str, value: str) -> int:
    """Pitfall #14 recipe against the LIVE document (see templating.py)."""
    n = 0
    while n < 500:
        find = _g(lambda: doc.Content.Find, None)
        if find is None:
            break
        try:
            find.ClearFormatting()
            find.Forward = True
            find.Wrap = 1
            hit = find.Execute(token)
        except Exception:
            break
        if not hit:
            break
        try:
            find.Parent.Text = value
            n += 1
        except Exception:
            break
    return n


def _apply_operations(doc: Any, operations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for i, op in enumerate(operations):
        kind = op["op"]
        if kind == "replace_all":
            n = _replace_all_two_step(doc, op["find"], op["replace"])
            results.append({"index": i, "op": kind, "find": op["find"], "replacements": n, "ok": True})
        elif kind == "insert_text":
            where = op["where"]
            text = op["text"]
            if where == "start":
                rng = doc.Range(0, 0)
                snap = _snapshot_font(_g(lambda rng=rng: rng.Font, {}) or {})
                rng.Text = text
                _restore_font(_g(lambda rng=rng: rng.Font, {}) or {}, snap, text)
            else:
                end = doc.Content.End
                rng = doc.Range(end - 1 if end > 0 else 0, end - 1 if end > 0 else 0)
                snap = _snapshot_font(_g(lambda rng=rng: rng.Font, {}) or {})
                rng.Text = text
                _restore_font(_g(lambda rng=rng: rng.Font, {}) or {}, snap, text)
            results.append({"index": i, "op": kind, "where": where, "chars": len(text), "ok": True})
        elif kind == "save":
            doc.Save()
            results.append({"index": i, "op": kind, "ok": True})
    return results


def edit_open_document(
    path_or_title: str,
    operations: list[dict[str, Any]],
    save: str = "ask",
) -> dict[str, Any]:
    """Attach to the user's live instance and edit an OPEN document."""
    t0 = time.perf_counter()
    if save not in ("ask", "save"):
        raise DocuhandError(
            f"save must be 'ask' or 'save' (got {save!r})",
            llm_hint="'ask' leaves the document dirty and reports the state; "
            "'save' writes the changes to disk after the last operation.",
        )

    app = _attach_running_app()
    doc, meta = _resolve_document(app, path_or_title)

    if meta.get("readonly"):
        raise DocuhandError(
            f"Document is open READ-ONLY in the user's session: {meta['full_name'] or meta['name']}",
            llm_hint="The user opened it read-only (or it is locked). Ask them to reopen it editable first.",
            details={"document": meta["full_name"] or meta["name"]},
        )

    op_results = _apply_operations(doc, operations)

    dirty = not bool(_g(lambda: doc.Saved, True))
    saved_now = False
    if save == "save":
        doc.Save()
        saved_now = True
        dirty = False

    return {
        "ok": True,
        "tool": "edit_open_document",
        "document": meta["full_name"] or meta["name"],
        "matched_by": "path" if meta["full_name"].lower() == str(Path(path_or_title).expanduser().resolve()).lower() else "title",
        "operations_applied": op_results,
        "saved": saved_now,
        "dirty_after": dirty,
        "engine": _g(lambda: app.Name, "unknown"),
        "engine_version": str(_g(lambda: app.Version, "") or ""),
        "duration_ms": round((time.perf_counter() - t0) * 1000),
    }
