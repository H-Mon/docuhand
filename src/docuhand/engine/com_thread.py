"""Dedicated STA thread for all COM work — architecture decision #1.

Why this exists:
- The MCP server runs on an asyncio event loop; calling COM from it explodes
  with RPC_E_WRONG_THREAD.
- Naive "one thread per call" COM usage launches a fresh WINWORD.EXE per
  call — the instance explosion.
- One long-lived single-threaded-apartment thread with a serial work queue
  fixes both: every COM operation happens on the same STA, fully serialized.

Timeout policy: a stuck call (Word showing a modal first-run / activation /
recovery dialog) cannot be interrupted safely. On timeout we *abandon* that
thread and its queue entirely (daemon thread; dies with the process) and
lazily build a fresh STA thread on the next submit. Any engine state that
lived on the abandoned STA is void and is reset via the abandon hook.
"""

from __future__ import annotations

import os
import queue
import threading
from typing import Any, Callable

from ..errors import ComCallTimeoutError

_DEFAULT_TIMEOUT_S = float(os.environ.get("DOCUHAND_COM_TIMEOUT", "90"))


class ComWorker:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._queue: queue.SimpleQueue | None = None
        self._thread: threading.Thread | None = None
        self._on_abandon: Callable[[], None] | None = None

    def set_abandon_hook(self, fn: Callable[[], None]) -> None:
        self._on_abandon = fn

    def submit(self, fn: Callable[..., Any], /, *args: Any, timeout: float = _DEFAULT_TIMEOUT_S) -> Any:
        """Run fn(*args) on the STA thread; raise on timeout (and abandon)."""
        with self._lock:
            if self._queue is None or self._thread is None or not self._thread.is_alive():
                self._queue = queue.SimpleQueue()
                self._thread = threading.Thread(
                    target=self._run,
                    args=(self._queue,),
                    name="docuhand-com-sta",
                    daemon=True,
                )
                self._thread.start()
        work_queue = self._queue

        box: queue.SimpleQueue = queue.SimpleQueue()
        work_queue.put((fn, args, box))
        try:
            ok, payload = box.get(timeout=timeout)
        except queue.Empty:
            self._abandon()
            raise ComCallTimeoutError(timeout) from None
        if ok:
            return payload
        raise payload

    def _abandon(self) -> None:
        with self._lock:
            self._queue = None
            self._thread = None
        if self._on_abandon:
            try:
                self._on_abandon()
            except Exception:
                pass

    @staticmethod
    def _run(work_queue: queue.SimpleQueue) -> None:
        import pythoncom

        pythoncom.CoInitialize()  # STA — the whole point of this thread
        try:
            while True:
                fn, args, box = work_queue.get()
                try:
                    box.put((True, fn(*args)))
                except BaseException as exc:  # propagate to submit()'s caller
                    box.put((False, exc))
        finally:
            pythoncom.CoUninitialize()
