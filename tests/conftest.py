"""Pytest-wide environment hygiene — must run before test modules import mcp.

Dev boxes sometimes export a global PYTHONPATH (this one injects the Hermes
install). Two distinct poisons for this suite:

1. Parent process: PYTHONPATH entries land in sys.path *before site-packages*,
   so a foreign venv's pure-Python packages (mcp, pydantic) shadow our own
   venv's — until a compiled ``.pyd`` mismatches and the import dies.
   Popping the env var is not enough; the sys.path entries must go too.
2. Child processes: spawned MCP servers inherit our environment, so a
   hostile PYTHONPATH would poison them the same way.

Sanitize both, once, here.
"""

import os
import sys

_pypath = os.environ.pop("PYTHONPATH", "")
for _var in ("PYTHONHOME", "PYTHONSTARTUP"):
    os.environ.pop(_var, None)

if _pypath:
    _foreign = {
        os.path.normcase(os.path.abspath(p))
        for p in _pypath.split(os.pathsep)
        if p.strip()
    }
    sys.path = [
        p
        for p in sys.path
        if os.path.normcase(os.path.abspath(p)) not in _foreign
    ]
