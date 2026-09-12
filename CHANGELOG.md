# Changelog

All notable changes to this project are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/) (loosely);
versions follow [Semantic Versioning](https://semver.org/) (pre-1.0: the
minor slot is the API surface).

## [0.1.0.dev1] - 2026-09-12

First public pre-release. The complete v0.1 tool set:

### Added
- `inspect_document` — pre-flight health check: true format vs extension,
  lock state, page/word/paragraph counts, corruption & password detection,
  protection mode, bookmarks, field codes.
- `convert_documents` — batch `.doc ↔ .docx` per directory; Word→WPS
  automatic failover; per-file fault isolation; sources never deleted;
  overwrite policy (`skip`/`overwrite`); optional output directory;
  pre-COM container & encryption gates (no modal-dialog hangs).
- `fill_template` — bookmarks / `{{placeholders}}` / content controls in
  `auto` mode; CJK font protection (`NameFarEast` re-assertion before and
  after `SaveAs2`); leftover-placeholder scan; per-key fill report.
- `export_pdf` — print-fidelity PDF via `ExportAsFixedFormat` with a
  `SaveAs(pdf)` fallback for older engines; page-range grammar
  (`all` | `7` | `3-9`); `%PDF-` header verification.
- `extract_content` — `text` / `tables` (row-major JSON grids from legacy
  `.doc`, merged cells keep the grid rectangular) / `outline` via
  locale-independent outline levels.
- `edit_open_document` — attaches to the user's live Word/WPS instance,
  resolves documents by path or window title, validates the whole operation
  batch before any COM call, `save="ask"` default never writes disk.
- `merge_documents` — one template × N data rows with per-row fault
  isolation; sanitized `{i}`/`{field}` output naming; collision-free targets.
- Safety layer: `DOCUHAND_ALLOWLIST` path allowlist, `DOCUHAND_READONLY`
  global read-only mode, JSONL audit log (`DOCUHAND_AUDIT`,
  `DOCUHAND_AUDIT_PATH`).
- Architecture: dedicated COM STA worker thread with serial queue and
  timeout-abandon semantics; Word→WPS engine probe/failover; structured
  errors (`error_code` / `human_message` / `llm_hint`); Chinese-Windows
  UTF-8 stdio fix; pure-logic planning modules unit-tested in CI.
- Documentation: `docs/pitfalls.md` — 15 battle-tested pitfalls from
  hands-on debugging on real machines.
