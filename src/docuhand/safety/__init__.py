"""Safety layer: allowlist enforcement + JSONL audit log.

Per the original design decision this is structural (registered around every
tool),
not a "remember to call it" convention.
"""

from .allowlist import check_allowed, get_allowlist
from .audit import record

__all__ = ["check_allowed", "get_allowlist", "record"]
