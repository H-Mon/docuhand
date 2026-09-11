"""Generate examples/convert-demo/ — a 10-file mixed .doc/.docx directory.

Reproduces the D3–D5 completion-criterion scenario locally (synthetic data
only, no real/personal information). Then point convert_documents at the
directory for the one-call conversion demo:

    uv run python scripts/make_mixed_dir.py
    uv run python -c "import json; from docuhand.tools.convert import run; \
print(json.dumps(run(r'examples/convert-demo')['summary'], indent=2))"
"""

from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client

DEMO_DIR = Path(__file__).resolve().parents[1] / "examples" / "convert-demo"
WD_FORMAT_DOC = 0
WD_FORMAT_DOCX = 16
WD_DO_NOT_SAVE = 0

TEXT = (
    "DocuHand convert demo file {i:02d}. "
    "This synthetic document exists to prove one-call batch conversion of a "
    "mixed .doc/.docx directory. 中文内容必须完整存活。Marker DEMO{i:02d}."
)


def main() -> int:
    DEMO_DIR.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Add()
        for i in range(1, 11):
            ext = "doc" if i % 2 == 1 else "docx"
            doc.Content.Text = TEXT.format(i=i)
            target = DEMO_DIR / f"mixed_demo_{i:02d}.{ext}"
            doc.SaveAs2(str(target), FileFormat=(WD_FORMAT_DOC if ext == "doc" else WD_FORMAT_DOCX))
            print(f"[OK] {target.name}")
        doc.Close(WD_DO_NOT_SAVE)
    finally:
        word.Quit(WD_DO_NOT_SAVE)
        pythoncom.CoUninitialize()
    print(f"\n10 mixed files ready in {DEMO_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
