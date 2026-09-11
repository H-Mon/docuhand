"""End-to-end convert_documents through a REAL MCP stdio client + REAL Word/WPS.

The completion criterion for D3–D5, as an executable proof: a directory with
10 mixed .doc/.docx files is converted with ONE tool call; a re-run is
idempotent (nothing re-converted, existing targets skipped); one corrupt
file fails alone without killing the batch; and DOCUHAND_ENGINE=wps proves
the WPS engine path end-to-end. Marked `com` (local runs, never CI).
"""

import asyncio
import json
import os
import sys
import winreg
from pathlib import Path

import pytest

pytestmark = pytest.mark.com

REPO = Path(__file__).resolve().parents[2]
PY = sys.executable


# --- helpers ---------------------------------------------------------------


def _wps_registered() -> bool:
    # WPS personal installs register ProgIDs under HKCU\Software\Classes;
    # machine-wide installs land in HKLM. Check both (32/64 views included).
    for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for view in (0, winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY):
            try:
                with winreg.OpenKey(root, r"SOFTWARE\Classes\KWPS.Application", 0, winreg.KEY_READ | view):
                    return True
            except OSError:
                continue
    return False


def _make_office_dir(parent: Path, variants: int, start_index: int = 1) -> Path:
    """Synthetic mixed .doc/.docx fixtures via real Word COM (no real data)."""
    import pythoncom
    import win32com.client

    parent.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            doc = word.Documents.Add()
            for i in range(start_index, start_index + variants):
                ext = "doc" if i % 2 == 1 else "docx"
                doc.Content.Text = (
                    f"DocuHand convert fixture variant {i:02d}. "
                    f"Unique marker VAR{i:02d}. 中文内容也必须完整存活。"
                )
                doc.SaveAs2(
                    str(parent / f"fixture_{i:02d}.{ext}"),
                    FileFormat=(0 if ext == "doc" else 16),
                )
            doc.Close(0)
        finally:
            word.Quit(0)
    finally:
        pythoncom.CoUninitialize()
    return parent


def _call_tool(args: dict, *, tool: str = "convert_documents", env: dict | None = None, timeout: float = 300, cwd: Path | None = None) -> dict:
    """Full MCP round trip against a freshly spawned server subprocess."""
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client

    async def _run() -> dict:
        params = StdioServerParameters(
            command=PY,
            args=["-m", "docuhand"],
            cwd=str(cwd or REPO),
            env={**os.environ, **(env or {})},
        )
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                result = await session.call_tool(tool, args)
                return json.loads(result.content[0].text)

    return asyncio.run(asyncio.wait_for(_run(), timeout=timeout))


# --- fixtures ---------------------------------------------------------------


@pytest.fixture(scope="module")
def mixed_dir(tmp_path_factory) -> Path:
    """10 mixed files: odd indexes .doc (5), even indexes .docx (5)."""
    return _make_office_dir(tmp_path_factory.mktemp("mixed10"), variants=10)


# --- the completion criterion -----------------------------------------------


def test_convert_mixed_ten_one_call(mixed_dir):
    r = _call_tool({"directory": str(mixed_dir)})
    assert r["ok"] is True, r
    s = r["summary"]
    assert s["doc_files_found"] == 10
    assert s["to_convert"] == 5
    assert s["already_target"] == 5
    assert s["converted"] == 5
    assert s["failed"] == 0
    assert r["sources_deleted"] is False

    converted = [e for e in r["results"] if e["status"] == "converted"]
    assert len(converted) == 5
    for e in converted:
        assert e["source"].lower().endswith(".doc")
        assert e["target"].lower().endswith(".docx")
        assert Path(e["target"]).exists()
        assert e["target_verified_format"] == "zip"  # a real .docx container
        assert (e["target_size_bytes"] or 0) > 0
        assert e["engine"] in ("word", "wps")

    already = [e for e in r["results"] if e["status"] == "already_target"]
    assert len(already) == 5
    assert all(e["source"].lower().endswith(".docx") for e in already)


def test_convert_rerun_is_idempotent(mixed_dir):
    # after the flagship run the dir holds 15 files (5 .doc + 10 .docx)
    r = _call_tool({"directory": str(mixed_dir)})
    assert r["ok"] is True, r
    s = r["summary"]
    assert s["doc_files_found"] == 15
    assert s["to_convert"] == 5
    assert s["already_target"] == 10
    assert s["converted"] == 0
    assert s["skipped_existing"] == 5  # overwrite='skip' protects the artifacts


def test_convert_isolates_failing_file(tmp_path):
    """One un-openable file must fail ALONE with a structured error while
    every other file still converts (the batch survival guarantee).

    Fixture candidates verified empirically against real Word:
    - truncated OLE2 garbage: Word silently REPAIRS and converts it (bad)
    - 0-byte .doc: Word opens it as an empty document (bad)
    - fake ZIP magic .docx: open fails instantly, every time (good)
    """
    d = _make_office_dir(tmp_path / "with_bad", variants=2)
    (d / "corrupt.docx").write_bytes(b"PK\x03\x04" + b"\x00" * 64 + b"GARBAGE-TAIL")

    # convert TOWARDS .doc so corrupt.docx enters the conversion queue
    # (classification is extension-based: in a to_format=docx run it would
    # be already_target and never touched — planner working as designed)
    r = _call_tool({"directory": str(d), "to_format": "doc"})
    assert r["ok"] is True, r  # one bad file must not kill the batch
    s = r["summary"]
    assert s["doc_files_found"] == 3  # fixture_01.doc + fixture_02.docx + corrupt.docx
    assert s["converted"] == 1  # fixture_02.docx → .doc
    assert s["already_target"] == 1  # fixture_01.doc untouched
    assert s["failed"] == 1

    failed = [e for e in r["results"] if e["status"] == "failed"]
    assert len(failed) == 1
    assert failed[0]["source"].endswith("corrupt.docx")
    assert failed[0]["error"]["error_code"] == "OPEN_FAILED"
    assert failed[0]["error"]["llm_hint"]


def test_convert_output_dir(tmp_path):
    d = _make_office_dir(tmp_path / "src", variants=2)
    out = tmp_path / "converted"

    r = _call_tool({"directory": str(d), "output_dir": str(out)})
    assert r["ok"] is True, r
    # variants=2 → one .doc (converted) + one .docx (already_target)
    assert r["summary"]["converted"] == 1
    assert r["summary"]["already_target"] == 1
    assert out.is_dir()
    # sources untouched, targets relocated
    assert len(list(d.glob("*.doc"))) == 1 and len(list(d.glob("*.docx"))) == 1
    assert len(list(out.glob("*.docx"))) == 1


@pytest.mark.skipif(not _wps_registered(), reason="WPS Office (KWPS) not installed")
def test_convert_wps_engine_end_to_end(tmp_path):
    """Word-absence simulated at the API boundary: force the WPS engine and
    prove the KWPS path converts a real legacy .doc to a valid .docx."""
    d = _make_office_dir(tmp_path / "wps_run", variants=1)  # one .doc

    r = _call_tool({"directory": str(d)}, env={"DOCUHAND_ENGINE": "wps"})
    assert r["ok"] is True, r
    assert r["engine_order"] == ["wps", "word"]
    converted = [e for e in r["results"] if e["status"] == "converted"]
    assert len(converted) == 1
    assert converted[0]["engine"] == "wps"
    assert converted[0]["target_verified_format"] == "zip"


def test_convert_accepts_relative_directory(mixed_dir):
    """Regression: WPS's Documents.Open dies on relative paths even when the
    server cwd is correct ('文档打开失败'). The engine must resolve paths
    to absolute server-side before any COM call (docs/pitfalls.md #12).
    Runs against the module fixture dir by bare name, relative to cwd."""
    r = _call_tool({"directory": mixed_dir.name}, cwd=mixed_dir.parent)
    assert r["ok"] is True, r
    # the payload reports the resolved absolute path
    assert Path(r["directory"]).is_absolute()
    # all planned conversions succeeded (targets exist from earlier tests
    # in this module, so the assert here is 'no relative-path failures')
    assert r["summary"]["failed"] == 0
