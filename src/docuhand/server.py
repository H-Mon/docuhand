"""MCP stdio entry point — deliberately a thin layer.

Protocol concerns live here; all real logic lives in engine/ and tools/.
If MCP is ever replaced, only this file and
the tool registration move — the engine carries zero MCP dependencies.
"""

from __future__ import annotations

import argparse

from mcp.server.fastmcp import FastMCP

from ._compat import fix_stdio
from .tools.convert import run as convert_documents_run
from .tools.edit_open import run as edit_open_document_run
from .tools.export_pdf import run as export_pdf_run
from .tools.extract import run as extract_content_run
from .tools.fill_template import run as fill_template_run
from .tools.inspect import run as inspect_document_run
from .tools.merge import run as merge_documents_run

fix_stdio()

mcp = FastMCP(
    "docuhand",
    instructions=(
        "DocuHand drives REAL Microsoft Word / WPS Office via COM on Windows. "
        "Use inspect_document as the pre-flight health check for any Word "
        "document: it reports the true format (legacy .doc vs .docx vs a "
        "renamed file), lock state (open in Word right now?), page/word/paragraph "
        "counts, corruption suspicion, password protection, editor restriction "
        "mode, bookmarks and field codes. Use convert_documents to batch-convert "
        "a whole directory of legacy .doc files to modern .docx (or the reverse) "
        "in one call — Microsoft Word is used when present and WPS Office is the "
        "automatic fallback. Use fill_template to produce a filled copy of a "
        "template via bookmarks, {{placeholders}} or content controls with CJK "
        "font protection; export_pdf for print-fidelity PDFs (optionally a page "
        "range); extract_content to pull text, tables (as JSON grids) or the "
        "heading outline out of legacy .doc files; merge_documents for "
        "mail-merge-style batch generation from data rows; and "
        "edit_open_document to edit a document the user has open RIGHT NOW "
        "(changes stay visible in their window; it never saves unless asked). "
        "Results are structured JSON; every error carries an llm_hint "
        "describing the next corrective step."
    ),
)


@mcp.tool()
def convert_documents(
    directory: str,
    to_format: str = "docx",
    recursive: bool = False,
    overwrite: str = "skip",
    output_dir: str | None = None,
) -> dict:
    """Batch-convert every .doc/.docx file in a directory in ONE call.

    Drives the real Word/WPS engine per file (Word first, WPS fallback when
    Word is absent or crashes). Files already in the target format are
    reported as already_target and left untouched; other extensions are
    ignored. Sources are NEVER deleted. One unreadable file fails alone
    with a structured error instead of killing the batch; only a total
    engine-launch failure aborts the run.

    Args:
        directory: Absolute Windows path to the folder holding the documents.
        to_format: Target format — "docx" (recommended) or "doc".
        recursive: Also descend into subdirectories.
        overwrite: "skip" (default) never touches an existing target file;
            "overwrite" replaces existing targets deliberately.
        output_dir: Optional target folder. Omit to convert in place (the
            converted file lands next to its source).
    """
    return convert_documents_run(
        directory,
        to_format=to_format,
        recursive=recursive,
        overwrite=overwrite,
        output_dir=output_dir,
    )


@mcp.tool()
def inspect_document(path: str) -> dict:
    """Health-check a Word document before touching it.

    Opens the document READ-ONLY in a hidden Word/WPS instance and reports:
    true format vs extension (catches renamed or pseudo-.doc files), page /
    word / paragraph counts, lock state (file currently open elsewhere),
    corruption suspicion, password protection, protection mode, bookmarks
    and field codes. Works on legacy .doc binaries that XML libraries cannot
    read, and on files currently open in Word/WPS.

    Args:
        path: Absolute Windows path to the document (e.g. D:\\\\Reports\\\\a.doc).
    """
    return inspect_document_run(path)


@mcp.tool()
def fill_template(
    template_path: str,
    data: dict,
    output_path: str | None = None,
    mode: str = "auto",
    bookmarks: list[str] | None = None,
) -> dict:
    """Fill a Word template with data and save a NEW document.

    Produces template_filled.<ext> next to the template by default (the
    template itself is never modified). Fill mechanisms, tried in 'auto':
    content controls (Tag/Title = key), bookmarks (name = key), and literal
    {{key}} placeholder tokens. Layout and East-Asian fonts are protected
    through the fill (NameFarEast re-assertion before AND after save).

    Args:
        template_path: Absolute Windows path to the .doc/.docx template.
        data: JSON object mapping keys to replacement text, e.g.
            {"name": "张三", "amount": "1234.56"} fills {{name}}/{{amount}},
            a bookmark named 'name', or a content control tagged 'name'.
        output_path: Where to write the filled document (.doc/.docx).
        mode: "auto" (default), "placeholders", "bookmarks" or "content_controls".
        bookmarks: Optional explicit bookmark name list (mode auto/bookmarks).
    """
    return fill_template_run(template_path, data, output_path, mode, bookmarks)


@mcp.tool()
def export_pdf(path: str, output_path: str | None = None, page_range: str | None = None) -> dict:
    """Export a print-fidelity PDF using the real Word/WPS layout engine.

    The PDF matches what would come out of the printer — not a re-layout by
    a converter library. Works on legacy .doc files and preserves CJK
    pagination. The source document is opened read-only and never modified.

    Args:
        path: Absolute Windows path to the document.
        output_path: Target .pdf path; defaults to the source name with .pdf.
        page_range: "all" (default), a single page "7", or a contiguous
            range "3-9" (1-based). Split comma lists into separate calls.
    """
    return export_pdf_run(path, output_path, page_range)


@mcp.tool()
def extract_content(path: str, what: str = "text") -> dict:
    """Extract text, tables or the heading outline from a Word document.

    Reads legacy .doc binaries that XML libraries cannot open, in a hidden
    READ-ONLY Word/WPS instance.

    Args:
        path: Absolute Windows path to the document.
        what: "text" — full plain text; "tables" — every table as a
            row-major JSON grid (merged cells repeat their origin text so
            the grid stays rectangular); "outline" — headings as
            [{level, text}] via locale-independent outline levels.
    """
    return extract_content_run(path, what)


@mcp.tool()
def edit_open_document(path_or_title: str, operations: list, save: str = "ask") -> dict:
    """Edit a document the user has OPEN in Word/WPS right now.

    Attaches to the user's running instance (not a hidden copy), finds the
    document by absolute path or window title, and applies small operations
    in order. The user SEES the changes happen in their window — this is
    the only Office MCP capability that works on live, locked files.

    Args:
        path_or_title: Absolute path, or the document's window title / file name.
        operations: Ordered list, each one of:
            {"op": "replace_all", "find": "2023", "replace": "2024"}
            {"op": "insert_text", "text": "Appendix", "where": "end"|"start"}
            {"op": "save"}
        save: "ask" (default) — do NOT save, report the dirty state and let
            the user decide; "save" — write changes to disk when done.
    """
    return edit_open_document_run(path_or_title, operations, save)


@mcp.tool()
def merge_documents(
    template_path: str,
    data_rows: list,
    output_dir: str | None = None,
    name_pattern: str = "{i}_{name}",
) -> dict:
    """Mail-merge style batch generation: one template × N data rows.

    Generates one filled document per row into output_dir (default:
    merged_<template-stem>/ next to the template). Rows fail in isolation —
    one bad row never stops the batch. Name outputs with {i} (1-based row,
    zero-padded) and any {field} tokens; names are sanitized.

    Args:
        template_path: Absolute Windows path to the .doc/.docx template with
            {{field}} placeholders and/or bookmarks.
        data_rows: Array of row objects, e.g. [{"name": "张三", "dept": "财务"},
            {"name": "李四", "dept": "法务"}] → one document each.
        output_dir: Target folder; created if missing.
        name_pattern: Output file name pattern (default "{i}_{name}").
    """
    return merge_documents_run(template_path, data_rows, output_dir, name_pattern)


def main() -> None:
    parser = argparse.ArgumentParser(prog="docuhand")
    sub = parser.add_subparsers(dest="command")
    sub.add_parser("serve", help="run the MCP stdio server (default)")
    args = parser.parse_args()
    if args.command not in (None, "serve"):
        parser.error(f"unknown command: {args.command}")
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
