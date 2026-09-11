"""Pre-COM container validation — the deterministic corrupt-file gate.

Why this exists (found the hard way, 2026-09-11): handing a structurally
broken file to Word/WPS via COM can pop a MODAL "file is corrupt" dialog.
DisplayAlerts=0 does not suppress it; FileValidation=1 (the documented
"skip validation" switch) made it WORSE on WPS: 88s hang vs 3s clean error
without it. The only airtight fix is to never hand a broken container to
the engine: validate the container structure in pure Python first and
reject trash with a structured OPEN_FAILED before any COM call. A modal
dialog cannot appear for a file the engine never sees.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import olefile

# Container magic bytes — defined here (not imported from office_app) to
# keep the dependency direction one-way: office_app → container_guard.
OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_MAGIC = b"PK\x03\x04"

# A valid legacy .doc MUST carry a WordDocument stream (or the 6.x-era
# WordDocument alias); the WPS writer may additionally expose bindings,
# but this stream is the constant across Word 97-2019 and WPS.
_WORD_STREAMS = ("worddocument", "1table", "0table")


def validate_container(path: str) -> tuple[bool, str]:
    """Return (ok, reason). ok=True → safe to hand to the COM engine.

    Extension-agnostic: the container decides, not the filename.
    Files that are NOT office containers at all (rtf/txt/html disguised as
    .doc) pass through: the engines open those natively without dialogs.
    """
    p = Path(path)
    try:
        with open(p, "rb") as f:
            head = f.read(8)
    except OSError as exc:
        return False, f"unreadable: {exc}"

    if head[:4] == ZIP_MAGIC:
        # OOXML family: a real .docx is a structurally valid zip with
        # [Content_Types].xml. Anything else is corrupt — and Word/WPS may
        # react to it with the modal we must avoid.
        try:
            size = p.stat().st_size
        except OSError as exc:
            return False, f"unreadable: {exc}"
        if size < 60:  # a zip EOCD alone is 22B; no real .docx is this small
            return False, "zip container truncated"
        try:
            with zipfile.ZipFile(p) as z:
                bad = z.testzip()
                if bad is not None:
                    return False, f"corrupt zip member: {bad}"
                names = set(z.namelist())
                if "[Content_Types].xml".lower() not in {n.lower() for n in names}:
                    return False, "zip lacks [Content_Types].xml (not an OOXML document)"
        except NotImplementedError:
            # exotic compression method — unusual but a real, openable docx
            return True, ""
        except zipfile.BadZipFile as exc:
            return False, f"not a valid zip container: {exc}"
        except OSError as exc:
            return False, f"unreadable: {exc}"
        return True, ""

    if head[:8] == OLE2_MAGIC:
        # Legacy .doc family: must be a parseable compound file carrying a
        # WordDocument stream. 8 bytes of OLE magic + garbage otherwise.
        try:
            size = p.stat().st_size
        except OSError as exc:
            return False, f"unreadable: {exc}"
        if size < 512:  # OLE header alone is 512B; no real .doc is smaller
            return False, "OLE2 container truncated"
        try:
            if not olefile.isOleFile(p):
                return False, "OLE2 header present but compound file is malformed"
            with olefile.OleFileIO(p) as ole:
                # listdir() entries are [storage, ..., stream] path tuples
                streams = {
                    entry[-1].lower()
                    for entry in ole.listdir()
                    if entry and isinstance(entry[-1], str)
                }
                if not ({"worddocument"} & streams):
                    # Not a Word doc — could be Excel .xls binary or legacy
                    # WPS .wps. The engines still open those natively without
                    # modals, so pass them through rather than mislabel.
                    return True, ""
                # sanity: WordDocument stream must be readable/non-empty
                data = ole.openstream("WordDocument").read(64)
                if not data:
                    return False, "WordDocument stream is empty"
        except OSError as exc:
            return False, f"malformed OLE2 compound file: {exc}"
        except Exception as exc:  # defensive: olefile raises varied parse errors
            return False, f"OLE2 parse error: {exc}"
        return True, ""

    # RTF / plain text / HTML / unknown disguised as .doc: the engines
    # convert these through their native converters, no modals observed.
    # (Verified: Word opens truncated-OLE2 as text-repair silently, and
    # plain garbage as an encoded-text recovery — neither prompts.)
    return True, ""
