"""MCP stdio entry point — deliberately a thin layer.

Protocol concerns live here; all real logic lives in engine/ and tools/.
If MCP is ever replaced, only this file and
the tool registration move — the engine carries zero MCP dependencies.
"""

from __future__ import annotations

import argparse

from mcp.server.fastmcp import FastMCP

from ._compat import fix_stdio
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
        "mode, bookmarks and field codes. Results are structured JSON; every "
        "error carries an llm_hint describing the next corrective step."
    ),
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
