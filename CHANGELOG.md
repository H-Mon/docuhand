# Changelog

All notable changes to this project are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/) (loosely);
versions follow [Semantic Versioning](https://semver.org/) (pre-1.0: the
minor slot is the API surface).

## [0.1.1] - 2026-11-17

### Fixed
- `docuhand.__version__` reported the wrong value. The version literal
  lived in two places - `pyproject.toml` and `src/docuhand/__init__.py` -
  and only the first was updated at release time, so the `0.1.0` wheel on
  PyPI reported `0.1.0.dev1` to `import docuhand`. The version is now
  read from the installed distribution's metadata, with the literal kept
  out of the source entirely, and `tests/unit/test_version.py` locks the
  two together.

## [Unreleased]

### Documentation
- `docs/pitfalls.md`: three new entries from migrating a real Chinese
  templates of the same form — #19 (WPS resolves `Font.Name`
  to a *different* East-Asian face; the fix is to leave `NameFarEast`
  empty, not to set it), #20 (merged cells throw on column indexing,
  vertical headers return one paragraph per character, source and target
  forms can have different shapes and different field semantics, template
  example values carry colour that you inherit, photos are `InlineShape`s
  governed by line-spacing rather than row height), #21 (template example
  content lives in table cells, body paragraphs *and* shapes — and two
  logical fields can share one physical paragraph, so paragraph-level
  replacement deletes a sibling).

## [0.1.0] - 2026-11-17

First public release. The complete v0.1 tool set:

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
- `export_pdf` — PDF via `ExportAsFixedFormat` with a
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
- Documentation: `docs/pitfalls.md` — 17 battle-tested pitfalls from
  hands-on debugging on real Chinese-Windows machines.

### Fixed
- `edit_open_document` no longer attaches to the wrong engine instance.
  A stale zero-document registration left in the Running Object Table
  (e.g. after a crashed engine) used to shadow the user's live instance,
  producing a misleading `Document not found among the 0 open document(s)`.
  The attach path now enumerates the ROT and prefers an instance that
  actually holds open documents. See `docs/pitfalls.md` #17.
