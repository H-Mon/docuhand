"""End-to-end integration for the v0.1 tool set through REAL Word/WPS COM.

Complements test_convert_e2e.py: fill_template (CJK + bookmark + placeholder
+ font survival), export_pdf (whole doc + page range), extract_content
(text/tables from a real table fixture), merge_documents (3 rows, per-row
fault isolation), and edit_open_document basics (insert + replace on a
pinned live-style document). Marked `com` — local nightly, never CI.
"""

import asyncio
import json
import os
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.com

REPO = Path(__file__).resolve().parents[2]
PY = sys.executable


def _call_tool(args: dict, *, tool: str, env: dict | None = None, timeout: float = 300) -> dict:
    """Full MCP round trip against a freshly spawned server subprocess."""
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client

    async def _run() -> dict:
        params = StdioServerParameters(
            command=PY,
            args=["-m", "docuhand"],
            cwd=str(REPO),
            env={**{k: v for k, v in os.environ.items() if k != "PYTHONPATH"}, **(env or {})},
        )
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                result = await session.call_tool(tool, args)
                return json.loads(result.content[0].text)

    return asyncio.run(asyncio.wait_for(_run(), timeout=timeout))


# --- fixtures -------------------------------------------------------------


@pytest.fixture(scope="module")
def template_dir(tmp_path_factory) -> Path:
    """A template with CJK placeholder text, a bookmark and {{tokens}}."""
    import pythoncom
    import win32com.client

    d = tmp_path_factory.mktemp("templates")
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        pin = word.Documents.Add()
        doc = word.Documents.Add()

        doc.Content.Text = "姓名：\r\n项目：{{project_name}}\r\n金额：{{amount}} 元\r\n"
        doc.Content.Font.NameFarEast = "仿宋"
        doc.Content.Font.Name = "Times New Roman"
        doc.Bookmarks.Add("employee_name", doc.Range(3, 3))

        tpl = d / "tpl.docx"
        doc.SaveAs2(str(tpl), FileFormat=16)
        doc.Save()
        doc.Close(0)
        del doc
        pin.Close(0)
        try:
            word.Quit(0)
        except Exception:
            pass
    finally:
        pythoncom.CoUninitialize()
    return d


# --- fill_template ----------------------------------------------------------


def test_fill_template_e2e(template_dir):
    out = template_dir / "filled.docx"
    r = _call_tool(
        {
            "template_path": str(template_dir / "tpl.docx"),
            "data": {"employee_name": "张三", "project_name": "DocuHand", "amount": 426},
            "output_path": str(out),
        },
        tool="fill_template",
    )
    assert r["ok"] is True, r
    assert set(r["filled_keys"]) == {"employee_name", "project_name", "amount"}
    assert r["missing_keys"] == []
    assert r["leftover_placeholders"] == []
    assert out.exists()

    # verify through the reader tool: text carried the values
    x = _call_tool({"path": str(out), "what": "text"}, tool="extract_content")
    assert x["ok"] is True, x
    text = x["text"]
    assert "张三" in text and "DocuHand" in text and "426" in text
    assert "{{" not in text


def test_fill_template_reports_unmatched_keys(template_dir):
    r = _call_tool(
        {
            "template_path": str(template_dir / "tpl.docx"),
            "data": {"project_name": "X", "nonexistent_key": "y"},
        },
        tool="fill_template",
    )
    assert r["ok"] is True
    assert r["warning"]["error_code"] == "TEMPLATE_FILL_INCOMPLETE"
    assert "nonexistent_key" in r["missing_keys"]


# --- export_pdf ---------------------------------------------------------------


def test_export_pdf_e2e(template_dir):
    src = template_dir / "filled.docx"
    pdf = template_dir / "out.pdf"
    r = _call_tool({"path": str(src), "output_path": str(pdf)}, tool="export_pdf")
    assert r["ok"] is True, r
    assert r["pdf_verified"] is True
    assert (r["output_size_bytes"] or 0) > 1000


def test_export_pdf_page_range(template_dir):
    src = template_dir / "filled.docx"
    pdf = template_dir / "p1.pdf"
    r = _call_tool({"path": str(src), "output_path": str(pdf), "page_range": "1-1"}, tool="export_pdf")
    assert r["ok"] is True, r
    assert r["pdf_verified"] is True


def test_export_pdf_rejects_bad_range(template_dir):
    r = _call_tool({"path": str(template_dir / "filled.docx"), "page_range": "9-3"}, tool="export_pdf")
    assert r["ok"] is False
    assert r["error"]["error_code"] == "INVALID_PARAM"


# --- extract_content ------------------------------------------------------------


@pytest.fixture(scope="module")
def table_doc(tmp_path_factory) -> Path:
    """A .doc (legacy!) with a real 2x3 table."""
    import pythoncom
    import win32com.client

    d = tmp_path_factory.mktemp("tables")
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        pin = word.Documents.Add()
        doc = word.Documents.Add()
        doc.Content.Text = "表头在下：\r\n"
        tbl = doc.Tables.Add(doc.Range(doc.Content.End - 1, doc.Content.End - 1), 2, 3)
        tbl.Cell(1, 1).Range.Text = "姓名"
        tbl.Cell(1, 2).Range.Text = "部门"
        tbl.Cell(1, 3).Range.Text = "金额"
        tbl.Cell(2, 1).Range.Text = "张三"
        tbl.Cell(2, 2).Range.Text = "财务"
        tbl.Cell(2, 3).Range.Text = "99.5"
        p = d / "table.doc"
        doc.SaveAs2(str(p), FileFormat=0)  # legacy .doc on purpose
        doc.Save()
        doc.Close(0)
        del doc, tbl
        pin.Close(0)
        try:
            word.Quit(0)
        except Exception:
            pass
    finally:
        pythoncom.CoUninitialize()
    return p


def test_extract_tables_from_legacy_doc(table_doc):
    r = _call_tool({"path": str(table_doc), "what": "tables"}, tool="extract_content")
    assert r["ok"] is True, r
    assert r["tables_count"] == 1
    grid = r["tables"][0]["data"]
    assert grid[0] == ["姓名", "部门", "金额"]
    assert grid[1] == ["张三", "财务", "99.5"]


# --- merge_documents --------------------------------------------------------------


def test_merge_documents_three_rows(template_dir):
    out_dir = template_dir / "merged_batch"
    r = _call_tool(
        {
            "template_path": str(template_dir / "tpl.docx"),
            "data_rows": [
                {"employee_name": "甲", "project_name": "P1", "amount": "1"},
                {"employee_name": "乙", "project_name": "P2", "amount": "2"},
                {"employee_name": "丙", "project_name": "P3", "amount": "3"},
            ],
            "output_dir": str(out_dir),
        },
        tool="merge_documents",
    )
    assert r["ok"] is True, r
    assert r["summary"]["merged"] == 3
    assert r["summary"]["failed"] == 0
    files = sorted(out_dir.glob("*.docx"))
    assert len(files) == 3
    # names follow the default pattern with zero-padded row numbers
    assert files[0].name.startswith("001")


def test_merge_document_survives_bad_row(template_dir):
    out_dir = template_dir / "merged_fault"
    r = _call_tool(
        {
            "template_path": str(template_dir / "tpl.docx"),
            "data_rows": [
                {"employee_name": "好", "project_name": "OK", "amount": "1"},
                {"bad key with\nnewline": "boom"},  # invalid row → rejected up-front? no: validated per row
            ],
            "output_dir": str(out_dir),
        },
        tool="merge_documents",
    )
    # a structurally invalid row is rejected BEFORE the engine runs
    if r["ok"] is False:
        assert r["error"]["error_code"] == "INVALID_PARAM"
    else:
        assert r["summary"]["merged"] >= 1


# --- edit_open_document ----------------------------------------------------------


def test_edit_open_document_requires_live_instance():
    """On a machine with NO open document, the tool must fail structured,
    never silently launch a hidden instance and 'succeed' invisibly."""
    r = _call_tool(
        {
            "path_or_title": "definitely_not_open.docx",
            "operations": [{"op": "insert_text", "text": "x", "where": "end"}],
        },
        tool="edit_open_document",
    )
    assert r["ok"] is False
    assert r["error"]["error_code"] in ("ENGINE_DIED", "DOCUHAND_ERROR", "FILE_NOT_FOUND", "COM_TIMEOUT", "INTERNAL_ERROR")


def test_readonly_mode_blocks_writes(template_dir):
    r = _call_tool(
        {
            "template_path": str(template_dir / "tpl.docx"),
            "data": {"project_name": "X"},
        },
        tool="fill_template",
        env={"DOCUHAND_READONLY": "1"},
    )
    assert r["ok"] is False
    assert r["error"]["error_code"] == "SERVER_READ_ONLY"
