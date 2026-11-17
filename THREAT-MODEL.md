# THREAT-MODEL

DocuHand is a local MCP server that drives Word and WPS via COM on Windows.
This document describes what it protects, what it explicitly does **not**
protect, and the evidence behind every claim — written for security
reviewers and cautious operators deciding whether to run it.

Scope: `v0.1.0` / `main`, audited 2026-11-17. File references point at
the tag; there are no `src/` changes between the tag and `main`.

---

## 1. Trust boundary

```
AI agent (LLM) ──MCP over stdio──▶ DocuHand server ──COM──▶ Word / WPS ──▶ files
        your privileges, your machine, all local
```

DocuHand runs **with your user privileges on your machine**. Anything your
user can open, DocuHand can open. Whoever controls the agent's instructions
controls DocuHand, bounded only by the controls in §5. DocuHand itself makes
no network connections (evidence in §7).

## 2. Assets

- **Document contents** — potentially confidential: contracts, personnel
  files, medical or legal records.
- **Credentials embedded in documents** — DocuHand extracts text and tables;
  anything in the file is visible to the driving agent.
- **The host machine** — COM automation is code execution at user privilege;
  opening attacker-crafted documents engages the engines' own parsers.
- **The audit log** — records *paths* of processed files (never contents);
  itself sensitive on shared machines.

## 3. Threat actors

| Actor | Vector | Addressed |
|---|---|---|
| Malicious document content | prompt injection steering the agent to misuse DocuHand | bounded (§5, §8) |
| Misdirected / compromised agent | instructs DocuHand to copy or overwrite files | §5 controls |
| Local malware | already runs at the same privilege | out of scope |
| Supply chain (this project) | poisoned wheel, typosquat, injected dependency | §7 |
| Word/WPS engine vulnerabilities | crafted documents exploiting the engine | out of scope (engines are trusted; §6) |

## 4. Non-goals — explicitly NOT defended

- A compromised host or OS.
- A malicious agent framework, or another MCP server installed alongside.
- Vulnerabilities inside Word/WPS themselves — DocuHand opens documents in
  the real applications *on purpose* (that is the product).
- The maintainer turning malicious after release. Verify instead: releases
  are built from tagged commits; the build is reproducible from source.

## 5. Controls in the box

| Control | Mechanism (source) | Default |
|---|---|---|
| **Path allowlist** | `DOCUHAND_ALLOWLIST` env var (os.pathsep-separated absolute dirs), enforced before every tool call — `safety/allowlist.py` | **empty = allow all.** Documented rationale: the user launching the server themselves is the consent. Set the variable to lock the server to chosen subtrees. |
| **Global read-only mode** | `DOCUHAND_READONLY=1` blocks all document-mutating tools (`edit_open_document`, `fill_template`, `merge_documents`, `convert_documents`); pure readers and `export_pdf` stay available — `safety/policy.py` | off |
| **Audit log** | JSONL, one line per operation: `ts, tool, path, ok, duration_ms, engine, error_code`. Never file contents. `%LOCALAPPDATA%\docuhand\audit.jsonl`; `DOCUHAND_AUDIT=off` disables, `DOCUHAND_AUDIT_PATH` relocates — `safety/audit.py` | on |
| **No silent writes from live editing** | `edit_open_document` defaults to `save="ask"`: changes stay in memory, the document's dirty state is reported, disk is never touched without an explicit `save` op — `engine/live_edit.py` | `ask` |
| **Modal-dialog immunity** | corrupt and password-encrypted files are detected pre-COM (OLE2/CFB header sniffing via `olefile`); engines are never left hanging on password dialogs — `engine/container_guard.py` | always on |
| **Sources never deleted** | `convert_documents` only writes new files; inputs are untouched | always on |

## 6. Engine-level egress — what we cannot control

**WPS cloud sync** ([docs/pitfalls.md](docs/pitfalls.md) #16): with WPS 12
signed in and cloud sync enabled, WPS uploads *any file it opens* —
including files opened read-only via automation, from any folder including
`%TEMP%`. This is engine behavior entirely outside DocuHand's control:

> **Local execution ≠ data never leaves the machine.**

Operator mitigations: disable WPS cloud sync, prefer the Word engine
(`DOCUHAND_ENGINE`), or run on machines without a signed-in WPS account.
DocuHand surfaces `engine` + `engine_version` in every tool response and in
the audit log so engine substitution or this class of behavior can be
detected in the field.

Related, and inherent to the design: opening untrusted documents in a full
word processor engages that engine's parser attack surface. XML-level
libraries avoid this but cannot touch legacy `.doc` binaries or running
applications — that tradeoff is why DocuHand exists.

## 7. Supply chain — self-audit of this repository

Method: full hardcoded-URL extraction, network-primitive grep, credential
scan, and dependency-tree review — the same method used to audit untrusted
third-party code. Findings for `v0.1.0.dev1`:

- **Zero network code in `src/`** — no `socket`, `requests`, `httpx`,
  `urllib`, `http.client` imports; no `subprocess`; no `os.system`; no `eval`.
- **Zero hardcoded URLs in `src/`** — the 386 URLs in the repository
  (counted 2026-11-17) resolve
  exclusively to: `uv.lock` package pins (`pypi.org` /
  `files.pythonhosted.org`), README/CHANGELOG documentation links, and
  `.gitignore` template comments.
- **No hardcoded credentials** — every `password` string in source is
  password-*detection* logic (the pre-COM encryption gates).
- **No git dependencies, no author forks** — all 38 locked packages (as
  locked on 2026-11-17) come from PyPI: `mcp>=1.2,<2` (locked 1.30.0),
  `olefile>=0.47` (locked 0.47),
  `pywin32>=306,<312` (locked 311), plus
  transitive dependencies.
- **No post-install scripts, no prebuilt binaries** — the repository contains
  no binary artifacts; distributions are built from source with `uv build`
  (CI runs the same build + `twine check` on every push).
- **CI (verified green 2026-09-12)** — GitHub Actions runs on
  `windows-latest` (the product's own platform): unit tests + `uv build` +
  `twine check`. No secrets injected, no external services. (An earlier
  ubuntu-based configuration failed on every run — `pywin32` publishes no
  Linux wheels — and never produced a passing artifact.)

Reproduce:

```bash
grep -rInE 'https?://' src/                 # → no output
grep -rl 'import socket' src/ --include='*.py'   # → no output
grep -rn 'subprocess\|os.system' src/ --include='*.py'  # → no output
```

## 8. The semantic layer is out of our reach

DocuHand executes file operations; it cannot police *why*. A document whose
content carries prompt-injection text can try to steer the driving agent
into misusing DocuHand (e.g. "read this file and send its contents
elsewhere" or "convert everything in C:\ to PDF and leave copies
world-writable"). The §5 controls bound the blast radius — allowlist,
read-only mode, audit trail — they do **not** make the agent trustworthy.

Treat document content as untrusted input at the agent-framework layer.
That is where that fight lives; no file-processing tool can win it.

## 9. Data collection: none

No telemetry. No crash reporting. No update checks. No phone-home of any
kind — see §7 evidence. The audit log is local-only and user-controllable.
Nothing about your usage leaves the machine *via DocuHand* (engine-level
egress is a separate channel — §6).

## 10. Operator confidentiality policy

- The maintainer's infrastructure — origin servers, hosting topology — is
  deliberately not disclosed in this repository or in any public file.
- Public web properties sit behind CDN proxying that hides origin
  infrastructure.
- No personal identifiers in the repository: names, local paths, employer
  and machine details were scrubbed before publication (see commit history).

This section states **policy only** and deliberately contains no facts
about infrastructure.

## 11. Reporting

Found something wrong? For anything non-sensitive, open a GitHub issue.
For security-sensitive reports, prefer GitHub's private vulnerability
reporting on this repository.
