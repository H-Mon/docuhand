"""Shared COM guard helpers — leaf module, no engine imports (no cycles).

``_g``: guard a single metadata probe — one engine lacking a property must
never fail the whole operation (the seed of docs/wps-vs-word.md).
"""

from __future__ import annotations

from typing import Any, Callable

OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _g(fn: Callable[[], Any], default: Any = None) -> Any:
    try:
        return fn()
    except Exception:
        return default


def _short_exc(exc: BaseException) -> str:
    return f"{type(exc).__name__}: {str(exc)[:200]}"


def _looks_encrypted(exc: BaseException) -> bool:
    text = str(exc).lower()
    return "password" in text or "密码" in text


def _looks_like_arg_mismatch(exc: BaseException) -> bool:
    return isinstance(exc, TypeError) or "paramet" in str(exc).lower()


def file_claims_encrypted(path: str) -> bool | None:
    """Header sniff: can this file really be COM-password-encrypted?

    True  → OLE2 with the FIB fEncrypted bit set: genuinely a password .doc.
    False → container rules it out (plain zip can't be COM-encrypted): a
            "password" error on it means the file is CORRUPT, not encrypted.
    None  → OLE2 without the bit — could be a normal .doc or an encrypted
            OOXML wrapper; treat encryption reports as genuine (safe default).
    """
    from pathlib import Path

    try:
        with open(Path(path), "rb") as f:
            head = f.read(16)
    except OSError:
        return None
    if head[:8] != OLE2_MAGIC:
        return False
    try:
        flags = int.from_bytes(head[0x0A:0x0C], "little")
        return bool(flags & 0x0001)
    except Exception:
        return None
