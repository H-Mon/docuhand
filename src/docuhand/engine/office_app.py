"""Office engine: probe Word/WPS via COM, failover, and inspect_document.

Threading contract: every method that touches COM runs on the STA thread
owned by :class:`~docuhand.engine.com_thread.ComWorker` (via submit).
Only plain-attribute methods (``invalidate``) may run on other threads.

Engine selection (architecture decision #6): probe order is
``Word.Application`` → ``KWPS.Application`` by default, overridable with
``DOCUHAND_ENGINE=word|wps``. If an engine dies mid-call (RPC crash —
seen in production when Word drives legacy .doc on some machines) the
engine is discarded and the next one takes over transparently.
"""

from __future__ import annotations

import os
import re
import time
from pathlib import Path
from typing import Any, Callable

import win32com.client

from ..errors import (
    DocuhandError,
    EngineUnavailableError,
    FileNotFoundError,
    OpenFailedError,
    PasswordProtectedError,
)

# --- pure helpers (unit-testable, zero COM) -------------------------------

OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_MAGIC = b"PK\x03\x04"
FILE_ATTRIBUTE_READONLY = 0x1

FRIENDLY_FORMATS = {
    "ole2": "legacy .doc (OLE2 Compound File)",
    "zip": "zip / OOXML (.docx family)",
    "rtf": "RTF",
    "html": "HTML",
    "text": "plain text",
    "unknown": "unknown",
}

_EXT_TO_FORMAT = {
    ".doc": "ole2",
    ".dot": "ole2",
    ".docx": "zip",
    ".dotx": "zip",
    ".rtf": "rtf",
    ".txt": "text",
    ".htm": "html",
    ".html": "html",
}

# Word constants (locale-independent numbers, not string names)
WD_STAT_WORDS = 0
WD_STAT_PAGES = 2
WD_STAT_PARAGRAPHS = 4
WD_DO_NOT_SAVE_CHANGES = 0
WD_ALERTS_NONE = 0

_PROTECTION_MODES = {
    -1: "none",
    -2: "comments",
    -3: "form_fields",
    -4: "read_only",
    -5: "tracked_changes",
}


def sniff_format(head: bytes) -> str:
    """Detect the real container format from the first bytes."""
    if head[:8] == OLE2_MAGIC:
        return "ole2"
    if head[:4] == ZIP_MAGIC:
        return "zip"
    stripped = head.lstrip()[:512].lower()
    if stripped.startswith(b"{\\rtf"):
        return "rtf"
    if b"<html" in stripped or stripped.startswith(b"<!doctype html"):
        return "html"
    if b"\x00" not in head[:512]:
        try:
            head[:512].decode("utf-8")
            return "text"
        except UnicodeDecodeError:
            pass
    return "unknown"


def format_matches_extension(extension: str, detected: str) -> bool:
    return _EXT_TO_FORMAT.get(extension.lower()) == detected


def probe_writable(path: str) -> bool:
    """Best-effort lock probe: can another writer exist right now?

    A file opened in Word/WPS is held with share-write denied, so a write
    handle fails. Note: a READONLY *attribute* also denies write handles,
    so callers must exclude that case before calling this a "lock".
    """
    try:
        with open(path, "r+b"):
            pass
        return True
    except OSError:
        return False


def owner_lock_file(path: str) -> str | None:
    """Word/WPS leave a '~$...' owner file next to an open document."""
    p = Path(path)
    try:
        for cand in p.parent.iterdir():
            name = cand.name.lower()
            if not name.startswith("~$"):
                continue
            rest = name[2:]
            target = p.name.lower()
            if target[2:] == rest or target.startswith(rest):
                return cand.name
    except OSError:
        pass
    return None


def _g(fn: Callable[[], Any], default: Any = None) -> Any:
    """Guard a single metadata probe: one engine lacking a property must
    never fail the whole inspection (the seed of docs/wps-vs-word.md)."""
    try:
        return fn()
    except Exception:
        return default


def _short_exc(exc: BaseException) -> str:
    return f"{type(exc).__name__}: {str(exc)[:200]}"


def _looks_encrypted(exc: BaseException) -> bool:
    text = str(exc).lower()
    return "password" in text or "密码" in text


def _looks_like_arg_mismatch(exc: BaseException) -> bool:
    return isinstance(exc, TypeError) or "paramet" in str(exc).lower()


# --- the engine ------------------------------------------------------------


class OfficeEngine:
    """Holds the hidden Word/WPS Application object — STA-only."""

    def __init__(self) -> None:
        self._app: Any = None
        self._engine_name: str | None = None
        self._progid: str | None = None

    # -- lifecycle ---------------------------------------------------------

    def invalidate(self) -> None:
        """Drop all engine state after its STA was abandoned (no COM calls)."""
        self._app = None
        self._engine_name = None
        self._progid = None

    def shutdown(self) -> None:
        """Quit the app — must run on the engine's own STA (via submit)."""
        app = self._app
        self.invalidate()
        if app is not None:
            try:
                app.Quit(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass

    def _discard_app(self) -> None:
        """Quit and forget after a mid-call engine failure (STA-only)."""
        self.shutdown()

    def _engine_order(self) -> list[tuple[str, str]]:
        prefer = os.environ.get("DOCUHAND_ENGINE", "").strip().lower()
        word_first = [("word", "Word.Application"), ("wps", "KWPS.Application")]
        wps_first = [("wps", "KWPS.Application"), ("word", "Word.Application")]
        if prefer == "wps":
            return wps_first
        if prefer == "word":
            return word_first
        return word_first  # default: probe Word first, WPS second

    def _ensure_app(self, progid: str, engine_name: str) -> Any:
        if self._app is not None and self._engine_name == engine_name:
            return self._app
        self.shutdown()  # drop any previous engine's app (on this STA)
        app = win32com.client.DispatchEx(progid)  # new hidden instance (Launch)
        for attr, value in (("DisplayAlerts", WD_ALERTS_NONE), ("Visible", False)):
            try:
                setattr(app, attr, value)
            except Exception:
                pass  # engine differences must not block launch
        try:
            app.Options.ConfirmConversions = False
        except Exception:
            pass
        self._app = app
        self._engine_name = engine_name
        self._progid = progid
        return app

    # -- the tool ----------------------------------------------------------

    def inspect_document(self, path: str) -> dict[str, Any]:
        t0 = time.perf_counter()
        p = Path(path).expanduser()
        if not p.exists():
            raise FileNotFoundError(str(p))
        if p.is_dir():
            raise DocuhandError(
                f"Path is a directory, not a document: {p}",
                llm_hint="Pass the full path to a document file, not a folder.",
                details={"path": str(p)},
            )

        size = p.stat().st_size
        with open(p, "rb") as f:
            head = f.read(1024)
        detected = sniff_format(head)
        extension = p.suffix.lower()

        stat = p.stat()
        readonly_attr = bool(getattr(stat, "st_file_attributes", 0) & FILE_ATTRIBUTE_READONLY)
        writable = probe_writable(str(p))
        is_locked = (not writable) and (not readonly_attr)
        lockfile = owner_lock_file(str(p))

        order = self._engine_order()
        attempts: list[dict[str, str]] = []
        launch_failures = 0

        for engine_name, progid in order:
            try:
                app = self._ensure_app(progid, engine_name)
            except Exception as exc:
                launch_failures += 1
                attempts.append({"engine": engine_name, "stage": "launch", "error": _short_exc(exc)})
                self._discard_app()
                continue
            try:
                info = self._open_and_read(app, engine_name, str(p))
            except PasswordProtectedError:
                raise  # failover cannot help; tell the agent immediately
            except Exception as exc:
                attempts.append({"engine": engine_name, "stage": "read", "error": _short_exc(exc)})
                self._discard_app()  # engine died or refused — try the next
                continue

            info["ok"] = True
            info["tool"] = "inspect_document"
            info["path"] = str(p)
            info["engine"] = engine_name
            info["format"] = {
                "extension": extension or None,
                "detected": detected,
                "detected_friendly": FRIENDLY_FORMATS.get(detected, detected),
                "matches_extension": format_matches_extension(extension, detected),
                "size_bytes": size,
            }
            info["is_locked"] = is_locked
            info["readonly_attribute"] = readonly_attr
            info["owner_lock_file_present"] = lockfile is not None
            info["is_corrupted"] = False
            info["duration_ms"] = round((time.perf_counter() - t0) * 1000)
            return info

        if launch_failures == len(order):
            raise EngineUnavailableError(attempts)
        raise OpenFailedError(str(p), "word→wps failover exhausted", "; ".join(a["error"] for a in attempts))

    # -- per-engine worker (runs inside the try above) ----------------------

    def _open_and_read(self, app: Any, engine_name: str, path: str) -> dict[str, Any]:
        docs = app.Documents
        try:
            doc = docs.Open(
                FileName=path,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )
        except Exception as exc:
            if _looks_encrypted(exc):
                raise PasswordProtectedError(path, engine_name) from None
            if not _looks_like_arg_mismatch(exc):
                raise
            doc = docs.Open(path)  # minimal fallback for engines that reject named args

        try:
            pages = _g(lambda: doc.ComputeStatistics(WD_STAT_PAGES))
            words = _g(lambda: doc.ComputeStatistics(WD_STAT_WORDS))
            paragraphs = _g(lambda: doc.ComputeStatistics(WD_STAT_PARAGRAPHS))
            protection_raw = _g(lambda: doc.ProtectionType, -1)
            revisions = _g(lambda: doc.Revisions.Count, 0) or 0

            fields_count = _g(lambda: doc.Fields.Count, 0) or 0
            sample_codes: list[str] = []
            for i in range(min(int(fields_count), 10)):
                code = _g(lambda i=i: re.sub(r"\s+", " ", doc.Fields.Item(i + 1).Code.Text).strip())
                if code:
                    sample_codes.append(code)

            bookmarks_count = _g(lambda: doc.Bookmarks.Count, 0) or 0
            bookmark_names: list[str] = []
            for i in range(min(int(bookmarks_count), 50)):
                name = _g(lambda i=i: doc.Bookmarks.Item(i + 1).Name)
                if name:
                    bookmark_names.append(name)

            protection = protection_raw if isinstance(protection_raw, int) else -1
            return {
                "opened_read_only": True,
                "pages": pages,
                "words": words,
                "paragraphs": paragraphs,
                "protection": _PROTECTION_MODES.get(protection, str(protection_raw)),
                "has_unaccepted_revisions": bool(revisions),
                "bookmarks": {"count": bookmarks_count, "names": bookmark_names},
                "fields": {"count": fields_count, "sample_codes": sample_codes},
            }
        finally:
            try:
                doc.Close(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
