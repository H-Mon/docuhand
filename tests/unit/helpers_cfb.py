"""Minimal CFB (OLE2) writer for synthetic test fixtures.

Test-only helper: builds a structurally valid 512-byte-sector compound
file with top-level streams (all stored as regular sectors — no mini
stream), readable by olefile. Derived from the MS-CFB spec ([MS-CFB]).
"""

from __future__ import annotations

import struct

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

SECTOR = 512

# Directory entry: name(64) namelen(H) type(B) color(B) left(I) right(I)
# child(I) clsid(16s) state(I) ctime(8s) mtime(8s) start(I) size(Q) = 128
_DIRENT = struct.Struct("<64sHBBIII16sI8s8sIQ")
assert _DIRENT.size == 128, _DIRENT.size


def _dirent(name: str, etype: int, color: int, left: int, right: int,
            child: int, start: int, size: int) -> bytes:
    raw = name.encode("utf-16-le") + b"\x00\x00"
    return _DIRENT.pack(
        raw.ljust(64, b"\x00"), len(raw), etype, color,
        left, right, child,
        b"\x00" * 16,           # CLSID
        0,                      # state bits
        b"\x00" * 8,           # creation time
        b"\x00" * 8,           # modified time
        start, size,
    )


def write_cfb(path: str, streams: dict[str, bytes]) -> None:
    """Write a minimal but olefile-readable CFB with top-level streams."""
    names = list(streams)
    n = len(names)

    # Sector plan: [0] directory (1 sector = 4 entries, enough for root+3;
    # extend if needed) ... then stream data ... then the single FAT sector.
    n_dir_entries = n + 1
    dir_sectors = max(1, -(-n_dir_entries * 128 // SECTOR))

    data_chunks = []
    data_sectors_per_stream = []
    for name in names:
        blob = streams[name]
        k = -(-len(blob) // SECTOR) if blob else 0
        data_sectors_per_stream.append(k)
        data_chunks.append(blob + b"\x00" * (k * SECTOR - len(blob)))
    n_data = sum(data_sectors_per_stream)

    total = dir_sectors + n_data + 1  # + FAT sector
    assert total <= 128, "fixture too large for the 1-sector FAT writer"

    # FAT: sector 0..dir_sectors-1 = directory chain; then per-stream
    # chains; last sector = FATSECT.
    fat = [FREESECT] * 128
    cur = 0
    for i in range(dir_sectors):
        fat[cur] = cur + 1 if i < dir_sectors - 1 else ENDOFCHAIN
        cur += 1
    starts = []
    for k in data_sectors_per_stream:
        starts.append(cur if k else ENDOFCHAIN)
        for i in range(k):
            fat[cur + i] = cur + i + 1 if i < k - 1 else ENDOFCHAIN
        cur += k
    fat_start = cur
    fat[fat_start] = FATSECT

    # Directory entries: root's CHILD starts the stream red-black tree;
    # siblings link via right (my fixtures are linear right-chains).
    entries = [_dirent("Root Entry", 5, 1, NOSTREAM,
                       NOSTREAM, 1 if n else NOSTREAM, ENDOFCHAIN, 0)]
    for i, name in enumerate(names):
        entries.append(_dirent(
            name, 2, 1, NOSTREAM,
            i + 2 if i + 1 < n else NOSTREAM,
            NOSTREAM, starts[i], len(streams[name]),
        ))
    dir_data = b"".join(entries)
    dir_data += b"\x00" * (dir_sectors * SECTOR - len(dir_data))

    # Header (MS-CFB 2.2)
    header = bytearray(b"\xff" * 512)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"          # signature
    struct.pack_into("<H", header, 24, 0x003E)                  # minor ver
    struct.pack_into("<H", header, 26, 3)                       # major ver 3
    struct.pack_into("<H", header, 28, 0xFFFE)                  # byte order
    struct.pack_into("<H", header, 30, 9)                       # sector shift
    struct.pack_into("<H", header, 32, 6)                       # mini shift
    struct.pack_into("<I", header, 40, 0)                       # num dir sectors (v3: 0)
    struct.pack_into("<I", header, 44, 1)                       # num FAT sectors
    struct.pack_into("<I", header, 48, 0)                       # first dir sector
    struct.pack_into("<I", header, 52, 0)                       # transaction
    struct.pack_into("<I", header, 56, 0x1000)                  # mini cutoff
    struct.pack_into("<I", header, 60, ENDOFCHAIN)              # mini FAT start
    struct.pack_into("<I", header, 64, 0)                       # num mini FAT
    struct.pack_into("<I", header, 68, ENDOFCHAIN)              # DIFAT start
    struct.pack_into("<I", header, 72, 0)                       # num DIFAT
    # DIFAT: sector NUMBERS of the FAT sectors (not FAT entry values!)
    struct.pack_into("<I", header, 76, fat_start)
    for i in range(1, 109):
        struct.pack_into("<I", header, 76 + 4 * i, FREESECT)

    # Root entry rebuild must match: child = first stream, not right.
    entries[0] = _dirent("Root Entry", 5, 1, NOSTREAM,
                         NOSTREAM, 1 if n else NOSTREAM, ENDOFCHAIN, 0)
    dir_data = b"".join(entries)
    dir_data += b"\x00" * (dir_sectors * SECTOR - len(dir_data))

    with open(path, "wb") as f:
        f.write(bytes(header))
        f.write(dir_data)
        for chunk in data_chunks:
            f.write(chunk)
        f.write(b"".join(struct.pack("<I", v) for v in fat))
