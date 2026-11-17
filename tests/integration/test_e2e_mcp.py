"""End-to-end smoke test through a REAL MCP stdio client + REAL Word/WPS COM.

This is exactly what a real MCP client does (Claude Desktop, Cursor, ...):
spawn the server as a subprocess, connect an MCP client over stdio, list
tools, call inspect_document on real files. Marked `com` (local runs, not CI).
"""

import asyncio
import json
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.com

REPO = Path(__file__).resolve().parents[2]
DEMO_DIR = REPO / "examples" / "demo-files"
PY = sys.executable


@pytest.fixture(scope="module")
def demo_files():
    doc = DEMO_DIR / "DocuHand_Demo.doc"
    docx = DEMO_DIR / "DocuHand_Demo.docx"
    assert doc.exists() and docx.exists(), "run scripts/make_demo_docs.py first"
    return doc, docx


def _call_tool(path: Path) -> dict:
    """Full MCP round trip: spawn server subprocess, initialize, call tool.

    The spawned server inherits this process's env — already sanitized by
    tests/conftest.py (no PYTHONPATH poison reaches the child).
    """
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client

    async def _run() -> dict:
        params = StdioServerParameters(
            command=PY,
            args=["-m", "docuhand"],
            cwd=str(REPO),
        )
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                result = await session.call_tool("inspect_document", {"path": str(path)})
                return json.loads(result.content[0].text)

    return asyncio.run(asyncio.wait_for(_run(), timeout=120))


def test_tools_listed(demo_files):
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client

    async def _run() -> list[str]:
        params = StdioServerParameters(command=PY, args=["-m", "docuhand"], cwd=str(REPO))
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                tools = await session.list_tools()
                return [t.name for t in tools.tools]

    names = asyncio.run(asyncio.wait_for(_run(), timeout=60))
    assert "inspect_document" in names


def test_inspect_doc(demo_files):
    doc, _ = demo_files
    r = _call_tool(doc)
    assert r["ok"] is True, r
    assert r["format"]["detected"] == "ole2"
    assert r["format"]["extension"] == ".doc"
    assert r["format"]["matches_extension"] is True
    assert isinstance(r["pages"], int) and r["pages"] >= 1
    assert r["is_locked"] in (True, False)


def test_inspect_docx(demo_files):
    _, docx = demo_files
    r = _call_tool(docx)
    assert r["ok"] is True, r
    assert r["format"]["detected"] == "zip"
    assert r["format"]["extension"] == ".docx"


def test_missing_file_is_structured_error(demo_files):
    doc, _ = demo_files
    r = _call_tool(doc.with_name("definitely_missing.doc"))
    assert r["ok"] is False
    assert r["error"]["error_code"] == "FILE_NOT_FOUND"
    assert r["error"]["llm_hint"]
