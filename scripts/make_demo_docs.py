"""Generate synthetic demo documents (no real/personal data) via real Word COM.

Creates examples/demo-files/DocuHand_Demo.{doc,docx} — the files the
end-to-end demo and integration tests run against.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client

DEMO_DIR = Path(__file__).resolve().parents[1] / "examples" / "demo-files"
WD_FORMAT_DOC = 0
WD_FORMAT_DOCX = 16
WD_DO_NOT_SAVE = 0
# Built-in style CONSTANTS are locale-independent; style *names* are not
# ("Heading 1" breaks on Chinese-UI Word — it's "标题 1" there).
WD_STYLE_NORMAL = -1
WD_STYLE_HEADING1 = -2
WD_STYLE_HEADING2 = -3

TEXT = [
    (WD_STYLE_HEADING1, "DocuHand Demo Document"),
    (WD_STYLE_NORMAL, "This file is synthetic demo data generated for the DocuHand end-to-end test. It contains no real or personal information."),
    (WD_STYLE_HEADING2, "Section 1: Format stress"),
    (WD_STYLE_NORMAL, "A legacy .doc binary like this one is unreadable to XML-based libraries. DocuHand drives the real application instead."),
    (WD_STYLE_NORMAL, "Chinese text sample: 文档体检 / 格式检测 / 锁定状态 — CJK content must survive every round trip."),
    (WD_STYLE_HEADING2, "Section 2: Structure"),
    (WD_STYLE_NORMAL, "The document carries a bookmark, a field code and multiple paragraphs so inspect_document has something real to report."),
]


def main() -> int:
    DEMO_DIR.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Add()
        sel = word.Selection
        for style, text in TEXT:
            sel.Style = doc.Styles(style)
            sel.TypeText(text)
            sel.TypeParagraph()
        doc.Bookmarks.Add("DemoBookmark", doc.Paragraphs(2).Range)
        doc.Fields.Add(doc.Paragraphs(1).Range, 33)  # 33 = wdFieldFileName

        docx_path = DEMO_DIR / "DocuHand_Demo.docx"
        doc.SaveAs2(str(docx_path), FileFormat=WD_FORMAT_DOCX)
        print(f"[OK] {docx_path}")

        doc.SaveAs2(str(DEMO_DIR / "DocuHand_Demo.doc"), FileFormat=WD_FORMAT_DOC)
        print(f"[OK] {DEMO_DIR / 'DocuHand_Demo.doc'}")

        doc.Close(WD_DO_NOT_SAVE)
    finally:
        word.Quit(WD_DO_NOT_SAVE)
        pythoncom.CoUninitialize()
    return 0


if __name__ == "__main__":
    sys.exit(main())
