"""Shared Word constant numbers — locale-independent (docs/pitfalls.md #1).

Kept in a leaf module so engine submodules can share them without
import cycles. Drive Word by number, never by display string.
"""

# document-close / quit disposition
WD_DO_NOT_SAVE_CHANGES = 0
WD_ALERTS_NONE = 0

# ComputeStatistics selectors
WD_STAT_WORDS = 0
WD_STAT_PAGES = 2
WD_STAT_PARAGRAPHS = 4

# ExportAsFixedFormat / SaveAs PDF
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_ALL_DOCUMENT = 0
WD_EXPORT_FROM_TO = 3

# SaveFormat
WD_FORMAT_DOC = 0
WD_FORMAT_DOCX = 16
