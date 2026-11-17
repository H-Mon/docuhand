# DocuHand 🖐️

**Give your AI agent hands to operate real Word & WPS documents on Windows.**

DocuHand is a Model Context Protocol (MCP) server that lets AI agents
(Claude Desktop, Cursor, Hermes, OpenClaw, or any other MCP client) actually
*drive* Microsoft Word and WPS Office via COM — not just read docx XML, but
automate the real applications: legacy `.doc` binaries, documents currently
open on screen, and PDF export through the word processor's own layout
engine.

**Who it is for:** you are running a smaller model — on a harness, locally,
or on a cheap API — against real files. It is smart enough to be useful and
not smart enough to be left alone. DocuHand gives it a stable tool surface
so it does not have to reason about *how* to touch a document.

**Who it is not for:** if you are on a frontier subscription and your agent
already writes the code you need, you will not get your money's worth here.

Works on **both** Microsoft Word and WPS Office. When Word is absent or
crashes mid-batch, the WPS engine takes over transparently. (WPS support is
believed to be unique among Office MCP servers.)

```bash
uvx docuhand serve
```

Security posture: **zero network code** (audited; see
[THREAT-MODEL.md](THREAT-MODEL.md)): no telemetry, no update checks, a
local JSONL audit log, and an explicit non-goals section.

---

## Using this from an AI agent

**If you are an AI agent**: read [examples/AGENT-QUICKSTART.md](examples/AGENT-QUICKSTART.md).
It is written for you — preconditions, the MCP config to paste, the exact
tool sequence to run, and what to do when a step fails.

**If you are a human wiring this up**: paste the following into your agent.
It is the prompt that was rehearsed end to end before it was published here.

> I have a legacy .doc file that contains a table. Please use the docuhand
> MCP server to:
> 1. inspect it first (`inspect_document`) and tell me the real format, page
>    count, table count, and whether it is locked or encrypted
> 2. extract its tables and show me the JSON grid (`extract_content`)
> 3. convert it to .docx (`convert_documents`)
> 4. fill the `DemoBookmark` in the converted file with a one-line note
>    (`fill_template`), saving to a new file
>
> Do not modify the original file.

No test file yet? One command downloads a synthetic `.doc` with a real table
in it and prints the path for you:

```bash
# PowerShell
irm https://raw.githubusercontent.com/H-Mon/docuhand/main/examples/quickstart.ps1 | iex

# Git Bash / WSL
curl -s https://raw.githubusercontent.com/H-Mon/docuhand/main/examples/quickstart.sh | bash
```

---

## Why DocuHand exists

Most Office MCP servers are built on `python-docx`: they manipulate XML in
idealized files. Real-world documents are not idealized:

- 20-year-old `.doc` binaries that XML libraries cannot touch
- Files currently **open and locked** in Word/WPS right now
- WPS Office — Kingsoft reported **676 million** WPS Office monthly active
  devices in June 2026, so the "default suite" problem is not a niche
- CJK fonts that silently reset when text is replaced (the `NameFarEast` trap)
- Password-protected and corrupt files that pop **modal dialogs** and hang
  headless servers

DocuHand drives the real applications, hardened by hands-on debugging on
real Chinese-Windows machines. Every pitfall in
[docs/pitfalls.md](docs/pitfalls.md) cost real debugging time, and you get
them pre-solved.

## Tools

| Tool | What it does |
|------|--------------|
| `inspect_document` | Pre-flight health check: real format vs extension, lock state, page/word counts, corruption & password detection, bookmarks, field codes |
| `convert_documents` | Batch `.doc ↔ .docx` for a whole directory in one call, Word→WPS failover, per-file fault isolation, sources never deleted |
| `fill_template` | Bookmarks / `{{placeholders}}` / content controls, CJK fonts protected through save |
| `export_pdf` | PDF via the real layout engine, optional page range |
| `extract_content` | Text / tables (as JSON grids) / heading outline from legacy `.doc` |
| `edit_open_document` | Edit a document the user **has open right now** — changes appear in their window |
| `merge_documents` | Mail-merge style batch generation: one template × N data rows |

Every tool returns structured JSON. A failure returns
`{"ok": false, "error": {error_code, human_message, llm_hint, details}}`:
errors are data the agent reads to self-correct, not tracebacks it drowns in.
`llm_hint` says how to fix the call — bad path, wrong argument, file outside
the allowlist.

## Quick start

Install from PyPI:

```bash
# with pip
pip install docuhand

# or with uv (also gives you the `uvx docuhand serve` one-liner)
uv tool install docuhand
```

Requirements: **Windows** with either Microsoft Word or WPS Office installed
(either one is enough; the engine probes `Word.Application` first, then
`KWPS.Application`, overridable with `DOCUHAND_ENGINE=wps|word`), and
**Python 3.12 or newer** (the package requires `>=3.12`; on older
interpreters pip reports `from versions: none`).

It runs with **your own user privileges**: no admin rights, no service
install, no daemon. It does not run on Linux or macOS: it drives the real
Office engine through COM.

**Test coverage:** the development machine has **WPS Office only, no
Microsoft Word**. Every end-to-end verification in this repository ran on
that machine; the Word-side code paths are the one thing not yet verified
against real Word. If you run on Word and hit a difference, please open an
issue, because that is exactly the gap worth closing.

Add to your MCP client (Claude Desktop example):

```json
{
  "mcpServers": {
    "docuhand": {
      "command": "docuhand",
      "args": ["serve"]
    }
  }
}
```

Then ask your agent: *"Convert every .doc in D:\\Reports to PDF"* — see
[Examples](#examples).

## Examples

**Examples of what to ask your agent**

Batch conversion — one call converts the whole folder:

> "Convert every .doc file in D:\Reports to .docx"

Template fill with CJK font safety — layout and East-Asian fonts survive:

> "Fill D:\Templates\notice.docx with name=Zhang San, amount=426, save it to my desktop"

Edit a document that's open on screen — the agent works in the user's
live Word/WPS window:

> "In the contract I have open right now, replace every 2023 with 2024 — don't save yet"

Mail-merge batch:

> "Generate one offer letter per row in data.json from the offer.docx template"

**Everything above runs against files already in this repo** — no setup,
no downloads:

| File | What it is |
|---|---|
| [examples/demo-files/DocuHand_Demo.doc](examples/demo-files/DocuHand_Demo.doc) | Legacy OLE2 `.doc` with a 4x4 table and a `DemoBookmark` — the file to try the prompt above on |
| [examples/demo-files/DocuHand_Demo.docx](examples/demo-files/DocuHand_Demo.docx) | The same content, modern format |
| [examples/convert-demo/](examples/convert-demo/) | 16 mixed `.doc`/`.docx` files for exercising `convert_documents` on a batch |
| [examples/hermes-config.json](examples/hermes-config.json) | MCP client config snippet |
| [examples/AGENT-QUICKSTART.md](examples/AGENT-QUICKSTART.md) | The step-by-step guide written for AI agents |
| [examples/quickstart.ps1](examples/quickstart.ps1) · [quickstart.sh](examples/quickstart.sh) | One-command demo fetch for PowerShell / bash |

All sample documents are synthetic (generated by
`scripts/make_demo_docs.py`) and contain no real or personal data.

## Security model (read before granting file access)

- **100% local execution** — your documents never leave your machine
- **No telemetry**, ever
- **Explicit path allowlist** — set `DOCUHAND_ALLOWLIST` to a
  `os.pathsep`-separated list of directories; when unset, the fact that you
  launched the server yourself is the consent
- **Read-only mode** — set `DOCUHAND_READONLY=1` to block every mutating
  tool; only `inspect_document` / `extract_content` / `export_pdf` remain
- **JSONL audit log** — every operation is appended to
  `%LOCALAPPDATA%\docuhand\audit.jsonl` (disable: `DOCUHAND_AUDIT=off`,
  relocate: `DOCUHAND_AUDIT_PATH`)
- **Pre-COM container gates** — broken/encrypted files are rejected in pure
  Python *before* the engine opens them, so a corrupt file can never pop a
  modal dialog on your desktop or hang the server (the #1 killer of headless
  Office automation)

Environment variables (`DOCUHAND_ENGINE`, `DOCUHAND_ALLOWLIST`,
`DOCUHAND_READONLY`, `DOCUHAND_AUDIT*`, `DOCUHAND_COM_TIMEOUT`) must be
declared in your MCP client's server config — hosts pass stdio servers a
filtered environment.

## Why not just use Copilot in Word / WPS AI?

Those are copilots: they serve the *person sitting in front of the
document*, through a closed, human-driven interface with no public API.
DocuHand is ground crew: it serves *agents working unattended*, from batch
jobs at 3 AM to pipelines, on any MCP client. Different buyer, different mode of
operation.

**"Why operate legacy .doc files at all?"**
Nobody *chooses* legacy formats. Archives of millions of `.doc` files are locked
by compliance (retain-original requirements), ecosystem inertia (counterparties
still sending `.doc`), and sheer scale. Converting them *is* one of our tools.
Copilot can't touch them; python-docx can't either. Manual clerical work or
DocuHand: those are the options.

## DocuHand vs. XML-level Office tools

|  | python-docx based | DocuHand |
|---|---|---|
| Legacy `.doc` binaries | ❌ | ✅ (real app required) |
| Open/locked files | ❌ | ✅ `edit_open_document` |
| WPS Office | ❌ | ✅ dedicated engine |
| CJK font safety | ⚠️ n/a or silent corruption | ✅ hardened (`NameFarEast` re-assertion) |
| PDF fidelity | ⚠️ re-layout | ✅ print-engine export |
| Corrupt/password files | ⚠️ crash or hang | ✅ pre-COM gates, structured errors |

## Architecture notes

The short version of the four decisions that make this survive real machines
(full rationale in [docs/pitfalls.md](docs/pitfalls.md)):

- **One dedicated STA thread** owns all COM: the asyncio event loop must
  never touch COM directly (`RPC_E_WRONG_THREAD`), and one serial queue
  means no WINWORD.EXE instance explosion. A stuck call times out and the
  thread is abandoned wholesale.
- **Word → WPS failover**: an engine that dies mid-call is discarded and
  the next one takes over; probe order is configurable.
- **Errors are structured data** for the calling agent, not stack traces.
- **Safety is registration-time middleware**, not "remember to check".

## Roadmap

- v0.2: `track_changes`, `compare_documents`, protect/unprotect, Excel family
- Pro (separate, later): batch queue/concurrency, service mode, priority support

The free core stays free (MIT).

## Development

```bash
git clone <this repo> && cd docuhand
uv sync

# CI-safe unit tests (no Office needed)
uv run pytest tests/unit

# COM integration tests (local machine with Word/WPS only, never CI)
uv run pytest tests -m com
```

## License

MIT. Free core, forever. Pro is a separate commercial license, later.
