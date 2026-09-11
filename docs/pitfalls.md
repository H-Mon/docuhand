# Battle-tested pitfalls

Every entry here cost real debugging time on real Chinese-Windows machines.
You get them pre-solved. (This file grows with every release — it is the
record of hard-won lessons, not marketing copy.)

## 1. Built-in style NAMES are locale-dependent; style CONSTANTS are not

`doc.Styles("Heading 1")` throws `com_error` on a Chinese-UI Word — the
style is called `标题 1` there. Use the locale-independent numeric
constants instead: `wdStyleHeading1 = -2`, `wdStyleNormal = -1`.
**Rule: drive Word by number, never by display string** (the same reason
`ComputeStatistics(2)` beats `"wdStatisticPages"`).

## 2. Quitting a driven Word instance can die with RPC_S_CALL_FAILED

A scripted `word.Quit()` may raise `(-2147023170, 'The remote procedure
call failed')` even when every prior call succeeded and the document is
already safely on disk. Do not let this kill your teardown: treat Quit as
best-effort (or swallow in `atexit`) and never retry document work against
an engine that RPC-dropped you — relaunch instead. This is the same crash
class that motivates DocuHand's per-call fault isolation and
Word→KWPS failover.

## 3. Drive COM from one dedicated STA thread

Calling COM objects from an asyncio event loop thread explodes with
`RPC_E_WRONG_THREAD`; calling from "a thread per operation" explodes with
a dozen orphaned WINWORD.EXE processes. One long-lived STA thread with a
serial work queue fixes both. On timeout, *abandon* the whole thread
(daemon) rather than trying to interrupt a stuck COM call — and reset any
engine state that lived on that thread.

## 4. Host MCP clients filter your subprocess environment

Hermes (and other MCP hosts) pass stdio servers a **filtered** environment:
only baseline vars (`PATH`, `SYSTEMROOT`, `APPDATA`, `TEMP`, ...) plus what
you declare in the server's `env` config. COM survives this fine — but if
your server needs anything else (e.g. an allowlist env var), it must be
declared in the MCP server config, not assumed.

## 5. GBK stdio is the invisible bomb on Chinese Windows

Python defaults stdout/stderr to the GBK codec there. One un-encodable
character in a JSON-RPC response (a Chinese path, a curly quote) raises
`UnicodeEncodeError` and kills the MCP transport. Reconfigure
`sys.stdout/stderr` to UTF-8 *before the first protocol write* — see
`docuhand/_compat.py`. Every MCP server on a CJK Windows machine needs
this; most don't know it yet.

## 6. pythonw.exe makes a cleaner stdio MCP server on Windows

Launching the server with `pythonw.exe` (windowless Python) avoids a
console window flashing per host restart, and stdio transport works
unchanged — verified with a real MCP client round-trip. Keep `python.exe`
for your own CLI debugging.

## 7. Lock detection needs three signals, not one

A file open in Word/WPS shows: write-handle denial + a `~$xxxx` owner
lockfile + (sometimes) a read-only attribute. But a read-only *attribute*
alone also denies write handles — so `is_locked` must exclude the
attribute case before calling it "open elsewhere", and the owner-lockfile
name is truncated to 8+3 legacy (`report.doc` → `~$port.doc`).
