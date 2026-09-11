"""Chinese-Windows stdio fix.

On Chinese Windows, Python's stdout/stderr default to the GBK codec. An MCP
stdio server writes JSON-RPC over stdout, so any non-GBK-encodable character
(a Chinese path, a curly quote in a document) raises UnicodeEncodeError and
kills the protocol stream. This is the invisible bomb behind most broken
MCP servers on CJK Windows machines.

We reconfigure both streams to UTF-8 *before* the MCP SDK writes anything.
Every entrypoint (``docuhand serve``, ``python -m docuhand``) calls this at
import time of :mod:`docuhand.server`.
"""

from __future__ import annotations

import sys


def fix_stdio() -> None:
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        if stream is None or not hasattr(stream, "reconfigure"):
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            # Never let a cosmetics fix break startup.
            pass
