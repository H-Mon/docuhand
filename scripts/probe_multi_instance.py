"""PROBE (manual, never in CI): can WPS host several independent COM instances?

Requires a real Word or WPS installation - this drives the actual engine.
Run it directly, not through pytest:

    uv run python scripts/probe_multi_instance.py

Measures, on the local machine:
  - 4 independent ComWorker + OfficeEngine pairs opening documents at the
    same time, each seeing only its own
  - concurrent .doc -> .docx writes, output validity, source integrity
  - office process count during the run and after shutdown

Result on WPS (2026-11-17): 4/4 succeeded, no cross-talk, sources
untouched, processes back to 0. See docs/pitfalls.md #18.
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

WORK = Path(os.environ.get("DOCUHAND_PROBE_DIR", "_p0_strict"))
DEMO = Path(REPO) / "examples" / "demo-files" / "DocuHand_Demo.doc"
N = 4


def office_procs() -> dict[str, int]:
    counts: dict[str, int] = {}
    try:
        out = os.popen("tasklist /FO CSV /NH").read()
    except Exception:
        return counts
    for line in out.splitlines():
        p = [x.strip('"') for x in line.split('","')]
        if len(p) < 2:
            continue
        n = p[0]
        if n.lower() in ("winword.exe", "wps.exe"):
            counts[n] = counts.get(n, 0) + 1
    return counts


def main() -> int:
    shutil.rmtree(WORK, ignore_errors=True)
    src = WORK / "src"
    src.mkdir(parents=True, exist_ok=True)
    for i in range(N):
        shutil.copy(DEMO, src / f"doc_{i}.doc")

    print("=" * 74)
    print("PROBE - true simultaneity + instance identity")
    print("=" * 74)
    print(f"  office procs BEFORE: {office_procs() or '(none)'}")

    barrier = threading.Barrier(N + 1, timeout=180)
    state: dict[int, dict] = {}
    lock = threading.Lock()
    workers = []
    engines = []

    def run(idx: int):
        w = ComWorker()
        e = OfficeEngine()
        w.set_abandon_hook(e.invalidate)
        workers.append(w)
        engines.append(e)

        def open_and_hold():
            app = e._ensure_app("Word.Application", "word")
            doc = app.Documents.Open(str((src / f"doc_{idx}.doc").resolve()),
                                     ReadOnly=True, AddToRecentFiles=False)
            return {
                "app_name": str(app.Name),
                "app_version": str(app.Version),
                "open_docs": int(app.Documents.Count),
                "doc_name": str(doc.Name),
                "doc_chars": int(doc.Characters.Count),
                "doc_tables": int(doc.Tables.Count),
            }

        try:
            info = w.submit(open_and_hold)
            with lock:
                state[idx] = {"ok": True, **info}
        except Exception as exc:
            with lock:
                state[idx] = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
        barrier.wait()

    threads = [threading.Thread(target=run, args=(i,)) for i in range(N)]
    for t in threads:
        t.start()

    barrier.wait()
    time.sleep(1.0)

    print()
    print("--- SIMULTANEOUS STATE (all documents open at once) ---")
    for i in range(N):
        s = state.get(i, {})
        if s.get("ok"):
            print(f"  worker {i}: app={s['app_name']} v{s['app_version']}  "
                  f"open_docs={s['open_docs']}  doc={s['doc_name']}  "
                  f"chars={s['doc_chars']}  tables={s['doc_tables']}")
        else:
            print(f"  worker {i}: FAILED {s.get('error')}")

    during = office_procs()
    print()
    print(f"  office procs DURING: {during or '(none)'}")

    ok = sum(1 for i in range(N) if state.get(i, {}).get("ok"))
    names = {state.get(i, {}).get("doc_name") for i in range(N) if state.get(i, {}).get("ok")}
    print(f"  workers with a document open simultaneously: {ok}/{N}")
    print(f"  distinct document names: {len(names)} -> {sorted(n for n in names if n)}")

    print()
    print("--- cleanup ---")
    for w, e in zip(workers, engines):
        try:
            w.submit(e.shutdown, timeout=20.0)
        except Exception:
            pass
    time.sleep(1.5)
    print(f"  office procs AFTER: {office_procs() or '(none)'}")

    print()
    print("=" * 74)
    verdict = "PASS" if ok == N and len(names) == N else "PARTIAL" if ok else "FAIL"
    print(f"VERDICT: {verdict}")
    if ok == N and len(names) == N:
        print("  => N independent Application instances, each with its own document.")
    print("=" * 74)
    return 0 if verdict == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
