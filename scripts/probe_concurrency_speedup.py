"""PROBE (manual, never in CI): what is the real concurrency speedup ceiling?

Requires a real Word or WPS installation - this drives the actual engine.
Run it directly, not through pytest:

    uv run python scripts/probe_concurrency_speedup.py

Measures serial vs parallel wall time at 1, 2 and 4 workers, both cold
and pre-warmed, plus per-file cost.

Result on WPS (2026-11-17): 1.00x / 1.46x / 1.48x - concurrency saturates
at two workers. See docs/pitfalls.md #18.
"""

from __future__ import annotations

import os
import shutil
import sys
import threading
import time
from pathlib import Path

REPO = r"D:/Repositories/docuhand"
sys.path.insert(0, os.path.join(REPO, "src"))

from docuhand.engine.com_thread import ComWorker  # noqa: E402
from docuhand.engine.office_app import OfficeEngine  # noqa: E402

WORK = Path(os.environ.get("DOCUHAND_PROBE_DIR", "_p0_speed2"))
DEMO = Path(REPO) / "examples" / "demo-files" / "DocuHand_Demo.doc"
WD_DOCX = 16


def prep(n: int, tag: str) -> Path:
    d = WORK / tag
    d.mkdir(parents=True, exist_ok=True)
    for i in range(n):
        shutil.copy(DEMO, d / f"s_{i}.doc")
    return d


def build_pool(workers: int):
    engines = [OfficeEngine() for _ in range(workers)]
    cms = [ComWorker() for _ in range(workers)]
    for c, e in zip(cms, engines):
        c.set_abandon_hook(e.invalidate)
    return cms, engines


def prewarm(cms, engines) -> float:
    """Launch every engine so launch cost is not measured."""
    t0 = time.perf_counter()

    def launch(w, e):
        def work():
            app = e._ensure_app("Word.Application", "word")
            d = app.Documents.Open(str(DEMO.resolve()), ReadOnly=True,
                                   AddToRecentFiles=False)
            d.Close(0)
        w.submit(work)

    ts = [threading.Thread(target=launch, args=(cms[i], engines[i]))
          for i in range(len(cms))]
    for t in ts:
        t.start()
    for t in ts:
        t.join(timeout=120)
    return time.perf_counter() - t0


def timed_run(cms, engines, files, workers: int):
    idx = [0]
    lock = threading.Lock()
    done: list[str] = []
    failed: list[str] = []

    def loop(w, e):
        while True:
            with lock:
                i = idx[0]
                idx[0] += 1
            if i >= len(files):
                return
            f = files[i]
            dst = f.with_suffix(".out.docx")
            try:
                w.submit(e.convert_document, str(f), str(dst), WD_DOCX)
                done.append(f.name)
            except Exception as exc:
                failed.append(f"{f.name}: {type(exc).__name__}")
            finally:
                dst.unlink(missing_ok=True)

    ts = [threading.Thread(target=loop, args=(cms[i], engines[i]))
          for i in range(workers)]
    t0 = time.perf_counter()
    for t in ts:
        t.start()
    for t in ts:
        t.join(timeout=900)
    return time.perf_counter() - t0, len(done), failed


def shutdown(cms, engines):
    for c, e in zip(cms, engines):
        try:
            c.submit(e.shutdown, timeout=20.0)
        except Exception:
            pass
    time.sleep(1.0)


def main() -> int:
    shutil.rmtree(WORK, ignore_errors=True)
    N = 24

    print("=" * 76)
    print("PROBE - serial vs parallel wall time")
    print("=" * 76)

    print()
    print("### A. PRE-WARMED runs (launch cost excluded)")
    print()
    base = None
    for w in (1, 2, 4):
        d = prep(N, f"w{w}")
        files = sorted(d.glob("*.doc"))
        cms, engines = build_pool(w)
        pw = prewarm(cms, engines)
        wall, ndone, failed = timed_run(cms, engines, files, w)
        shutdown(cms, engines)
        if base is None:
            base = wall
        print(f"  workers={w}: prewarm {pw:.1f}s | convert {wall:.1f}s | "
              f"done={ndone} failed={len(failed)} | speedup vs serial = {base/wall:.2f}x")
        for f in failed[:2]:
            print(f"      !! {f}")

    print()
    print("### B. COLD runs (launch cost included)")
    print()
    base = None
    for w in (1, 2, 4):
        d = prep(N, f"c{w}")
        files = sorted(d.glob("*.doc"))
        cms, engines = build_pool(w)
        wall, ndone, failed = timed_run(cms, engines, files, w)
        shutdown(cms, engines)
        if base is None:
            base = wall
        print(f"  workers={w}: convert {wall:.1f}s (cold) | done={ndone} "
              f"failed={len(failed)} | speedup vs serial = {base/wall:.2f}x")

    print()
    print("### C. Per-file cost (warm)")
    print()
    print("  workers |  per-file  |  files/sec")
    print("  -------+------------+-----------")
    for w in (1, 2, 4):
        d = prep(N, f"p{w}")
        files = sorted(d.glob("*.doc"))
        cms, engines = build_pool(w)
        prewarm(cms, engines)
        wall, ndone, _ = timed_run(cms, engines, files, w)
        shutdown(cms, engines)
        print(f"    {w}    |  {wall/ndone*1000:6.0f} ms |  {ndone/wall:6.2f}")

    print()
    print("=" * 76)
    print("  A speedup well below the worker count means the engine serialises")
    print("  internally somewhere. Do not market it as N times faster.")
    print("=" * 76)
    return 0


if __name__ == "__main__":
    sys.exit(main())
