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

## 8. A poisoned PYTHONPATH explodes only at the compiled layer — and popping
## the env var in conftest is not enough

Dev boxes with a global `PYTHONPATH` (this one injects the Hermes install)
look healthy until two venvs both ship the same package: the foreign
site-packages sits **ahead of the project venv in `sys.path`**, pure-Python
imports silently shadow (mcp, pydantic), and then `pydantic_core` dies with
`ModuleNotFoundError` on a binary `.pyd` version mismatch — far away from
the actual mistake.

Two traps that cost an afternoon:

- `os.environ.pop("PYTHONPATH")` **cannot save the parent pytest process**:
  the paths were already expanded into `sys.path` at interpreter startup.
  The env var and the `sys.path` entries must be scrubbed separately.
- `StdioServerParameters(env=...)` does not help either — the import dies
  in the *pytest parent*, before any child server is spawned.

Fix (`tests/conftest.py`): pop the env var (children then inherit a clean
environment) AND filter every PYTHONPATH entry out of `sys.path` (the
parent then imports its own venv). Both, once, before any test module
imports `mcp`.

## 9. A corrupt file can pop a MODAL dialog through COM — headless is a lie
### (and every in-COM defense fails somewhere)

Feeding a structurally broken file to `Documents.Open` via COM can raise a
**modal "file is corrupt" dialog** on the user's desktop — even though your
app is headless and `DisplayAlerts = 0`. Empirical scorecard on a fake-zip
`.docx`:

| defense | result |
|---|---|
| `DisplayAlerts = 0` | dialog still appears |
| `Application.FileValidation = 1` (wdFileValidationSkip) | **88 s hang** on WPS — the skip path itself stalls |
| `Open(..., PasswordDocument="wrong-pw")` | 3 s clean error on Word; on **WPS the dynamic dispatch explodes with `AttributeError` on the unknown kwarg and the app object dies** (next call: RPC_S_SERVER_UNAVAILABLE) |
| **validate the container in pure Python BEFORE COM** (`engine/container_guard.py`) | **modal is physically impossible — the engine never sees the trash** |

Only the last one is portable across Word and WPS, and it is the only one
that cannot regress with a new Office version. When a "password" error
does arrive, trust it only if the container could actually be encrypted
(check the OLE2 FIB fEncrypted bit at offset 0x0A); otherwise it is Word
misjudging corruption as encryption — report OPEN_FAILED, not
"ask the user for the password".

**Rule: `Documents.Open` is a window into the user's desktop. Anything
you can decide without the engine, decide without the engine.**

## 10. `Word.Application` may not be Word — ProgID hijacking on WPS machines

On a machine with WPS Office but no MS Word, the `Word.Application` ProgID
can be **registered by WPS** ( compat takeover). Every COM probe then
"succeeds with Word" while `'Kingsoft WPS'` sits in the `com_error`
source. Two consequences:

- Never trust `engine == "word"` for reporting unless you verified the
  binary, not just the ProgID.
- Machine-wide (`HKLM`) registry probing misses per-user WPS installs:
  check `HKCU\Software\Classes\<ProgID>` too, or your "is WPS available?"
  test says no while WPS works fine.

## 11. git-bash heredocs feeding `python -` can hang forever

`uv run python <<'EOF' ... EOF` from git-bash/MSYS on Windows may deliver
stdin in a way `python -` never sees EOF for — the process hangs past its
own logic with zero output (and stdout redirected to a file shows nothing
because of block buffering). Write the script to a real file and run it;
add `flush=True` to progress prints. Same box: PowerShell's `$_` is eaten
by git-bash too — put PowerShell logic in a `.ps1` file instead.

## 12. WPS's `Documents.Open` rejects RELATIVE paths with a bare failure

Hand WPS a relative path (`examples\\dir\\a.doc`) while your server's cwd
is perfectly correct and it fails with `文档打开失败` (error 3010) — the
same call with an absolute path succeeds. Word tolerates relative paths;
WPS does not. Rule: **`Path(...).expanduser().resolve()` every path at the
engine boundary**, before it can reach COM. The server's cwd is an
implementation detail the engine must never depend on.

## 13. Old Office compat: a missing COM method raises AttributeError, not
## TypeError

`SaveAs2` did not exist before Word 2010. On a Word 2007 machine, late
binding fails at **attribute access** with `AttributeError` — the call
never happens, so a `TypeError`-shaped fallback net never fires and every
conversion dies with OPEN_FAILED. Whenever you use a "modern" COM method
with an old-engine fallback (`SaveAs2→SaveAs`, `ContentControls`→nothing,
`OpenNoRepairDialog`→`Open`), the fallback must catch **both**
`AttributeError` (method absent) and the arg-mismatch family (method
present but rejects the signature). Many deployed machines still run
routinely still run Win7 + Office 2007/2010 — treat 2007 as a supported
floor until the field data says otherwise.

## 14. `Find.Execute(Replace:=wdReplaceAll)` returns True and replaces
## NOTHING on WPS-as-Word

On this engine family (WPS 12 masquerading as `Word.Application`,
reporting `Name = "Microsoft Word"`, `Version = 12.0`), the one-call
`Find.Execute(FindText=..., ReplaceWith=..., Replace=2)` completes
**successfully** (returns True) yet the text is unchanged. No error, no
dialog — the Replacement is silently dropped. The portable recipe (probe-
verified on both the hijacked Word ProgID and real `KWPS.Application`):

```python
# per hit: FRESH story range + fresh Find, find-only, then set range text
while True:
    f = doc.Content.Find          # fresh Range + Find every iteration
    f.ClearFormatting()
    f.Forward = True
    f.Wrap = 1                    # wdFindContinue
    if not f.Execute(token):      # find ONLY — never pass Replace here
        break
    f.Parent.Text = value         # Parent IS the matched range
```

Subtleties that each cost a probe cycle:
- Reusing one `Find` object across hits stops after the **first**
  replacement (its internal range state is consumed by the hit).
- A fresh `Find` from a *shared* Range object gets 2 of 3 — only a fresh
  `doc.Content` per iteration is reliable.
- `Wrap=0` (wdFindStop) under-counts; `Wrap=1` (continue) finds all.
- A value that contains its own token would loop forever — cap iterations
  and never interpolate data into `FindText`.
- Same family trait: an un-saved doc's `Close()` can kill the whole app
  instance (RPC_S_SERVER_UNAVAILABLE on the next property access) when the
  instance is headless with zero documents. Keep a sentinel doc open
  across multi-step scripts, `doc.Save()` before `Close(0)` when the doc
  was just created, and treat `Quit` as best-effort (pitfall #2).

## 15. `Cell.Text` does not exist on WPS-as-Word — read cells via
## `Cell.Range.Text`

Reading table cells with `table.Cell(r, c).Text` — the snippet every
tutorial shows — raises `AttributeError: Cell.Text` on this engine
family. The portable read is one indirection deeper:
`table.Cell(r, c).Range.Text`, which returns the cell content with a
`\r\x07` end-of-cell marker — strip `\x07` and `\r` (markdown-free cells
matter for JSON output). Verified against a round-tripped legacy `.doc`:
2×3 grid with CJK + decimal values reads back byte-perfect. Also: two
`DispatchEx` launches inside one process can hard-segfault the second
WPS instance (access violation in win32com) — probe scripts must reuse
one app instance with a sentinel document.

## 16. WPS 12 auto-uploads any file the engine OPENS to its cloud —
## including your test fixtures

On a default WPS 12 install with a logged-in account
(`HKCU\Software\Kingsoft\wpscloud :: uploadFile = true`), every document
the COM engine merely **opens** (even READ-ONLY, even from %TEMP%, no
SaveAs needed) gets silently pushed to the user's WPS cloud drive — it
then shows up in 传输助手 as "已上传". Verified experimentally:

- plain `SaveAs2` from a 10-line script → uploaded in ≤5 s
- COM `Documents.Open(ReadOnly=True)`, zero writes → uploaded
- pure file copy (WPS never touches it) → NOT uploaded

So the trigger is the OPEN, not the write, and it is not specific to
DocuHand — any script (or Word-if-installed equivalent workflow) driving
this engine family uploads. Consequences for an automation tool:

1. **Test fixtures must be synthetic** — ours are (see
   `scripts/make_demo_docs.py`). Never point DocuHand at documents you
   would not hand to WPS Cloud.
2. For privacy-sensitive environments, disable cloud sync machine-wide:
   WPS 配置工具 → 备份中心/云同步 off, or
   `HKCU\Software\Kingsoft\wpscloud: uploadFile=false`,
   `DesktopSync=false`, `SpecifyFolderSync=false`, then kill `wpscloudsvr.exe`.
3. THREAT-MODEL note: a local-execution promise ("documents never leave
   your machine") is bounded by the engine's own cloud behavior. DocuHand
   can and should SAY this: if your WPS is logged in with cloud sync on,
   WPS — not DocuHand — may copy opens to the cloud. Surface it in the
   security README section.

## 17. A stale ROT registration makes the user's OPEN document
## unreachable — `Dispatch` returns the wrong instance

`edit_open_document` attaches to the engine the user is actually running.
It used to do that with `win32com.client.Dispatch("Word.Application")`,
which resolves to **one** Running-Object-Table registration. Problem: the
ROT can hold **several** `Word.Application` registrations at once, and a
stale one — a headless `DispatchEx` instance whose `Quit()` died (pitfall
#2) — sits in front of the live one.

Observed on WPS 12 (2026-11-17), reproduced deliberately:

```
# GUI has save_probe.docx open
GetActiveObject("Word.Application").Documents.Count  ->  0   # the STALE one
ROT enumeration                                     ->  4 instances
```

With the stale instance first, `_resolve_document` raised
`Document not found among the 0 open document(s)` — a **misleading**
error, because the document *was* open. The user would reasonably think
their file had not been opened, and would re-open it, and it would still
fail.

The fix enumerates the ROT and prefers an instance that actually has
`Documents.Count > 0`. Two implementation notes, both required:

- `moniker.BindToObject(ctx, None, IID_IDispatch)` **fails** on the Word
  CLSID moniker with `E_INVALIDARG (-2147024809)`. The working shape is
  `rot.GetObject(moniker)` → `unk.QueryInterface(IID_IDispatch)`.
  (`PyIMoniker` has no `GetObject` method of its own.)
- WPS impersonates Word's ProgID, so `Application.Name` cannot be used to
  filter candidates. Filter on "exposes a `Documents` collection" instead.

Always fail *soft* here: if ROT enumeration raises for any reason, fall
back to the ProgID lookup rather than propagating.

