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
from .tools.inspect import run as inspect_document_run

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
        "automatic fallback; source files are never deleted. Results are "
        "structured JSON; every error carries an llm_hint describing the next "
        "corrective step."
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
        path: Absolute Windows path to the document (e.g. D:\\Reports\\a.doc).
    """
    return inspect_document_run(path)


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
