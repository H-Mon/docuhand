"""Engine package: dedicated COM STA worker + Office engine with Word/WPS failover."""

from __future__ import annotations

import atexit

from .com_thread import ComWorker
from .office_app import OfficeEngine, sniff_format

_worker = ComWorker()
_engine = OfficeEngine()

# A timed-out STA is abandoned wholesale; engine state that lived on that
# thread is void — reset it so the next call relaunches cleanly on a fresh
# STA thread. invalidate() only writes plain attributes (GIL-atomic), it
# never touches COM objects.
_worker.set_abandon_hook(_engine.invalidate)


def get_worker() -> ComWorker:
    return _worker


def get_engine() -> OfficeEngine:
    return _engine


def _shutdown() -> None:
    """Quit the hidden Word/WPS instance on its own STA thread at exit."""
    try:
        _worker.submit(_engine.shutdown, timeout=15.0)
    except Exception:
        pass  # process teardown must never hang on a sick COM app


atexit.register(_shutdown)

__all__ = ["ComWorker", "OfficeEngine", "get_engine", "get_worker", "sniff_format"]
