"""Pure planning for edit_open_document — operation validation, zero COM.

An edit batch is a list of small operations executed in order against one
open document. Everything is validated BEFORE any COM call happens so a
bad batch never reaches the engine mid-way.

Operations (v0.1):
- replace_all   {find, replace}          literal find→set two-step (pitfall #14)
- insert_text   {text, where}            where: start|end
- save          {}                       explicit save checkpoint
v0.2 reserved: set_bookmark, apply_style, delete_range.
"""

from __future__ import annotations

from typing import Any

MAX_OPS = 50
MAX_FIND_LEN = 500
MAX_TEXT_LEN = 100_000

KNOWN_OPS = {"replace_all", "insert_text", "save"}


def validate_operations(operations: Any) -> list[dict[str, Any]]:
    """Validate the op batch up-front; raises ValueError on any problem."""
    if not isinstance(operations, list) or not operations:
        raise ValueError("operations must be a non-empty list of edit operations")
    if len(operations) > MAX_OPS:
        raise ValueError(f"at most {MAX_OPS} operations per call (got {len(operations)})")
    clean: list[dict[str, Any]] = []
    for i, op in enumerate(operations):
        if not isinstance(op, dict):
            raise ValueError(f"operations[{i}] must be an object")
        kind = op.get("op")
        if kind not in KNOWN_OPS:
            raise ValueError(
                f"operations[{i}].op must be one of {sorted(KNOWN_OPS)} (got {kind!r})"
            )
        if kind == "replace_all":
            find = op.get("find")
            replace = op.get("replace", "")
            if not isinstance(find, str) or not find:
                raise ValueError(f"operations[{i}].find must be a non-empty string")
            if len(find) > MAX_FIND_LEN:
                raise ValueError(f"operations[{i}].find exceeds {MAX_FIND_LEN} chars")
            if not isinstance(replace, str):
                raise ValueError(f"operations[{i}].replace must be a string")
            if len(replace) > MAX_TEXT_LEN:
                raise ValueError(f"operations[{i}].replace exceeds {MAX_TEXT_LEN} chars")
            clean.append({"op": "replace_all", "find": find, "replace": replace})
        elif kind == "insert_text":
            text = op.get("text")
            where = op.get("where", "end")
            if not isinstance(text, str) or not text:
                raise ValueError(f"operations[{i}].text must be a non-empty string")
            if len(text) > MAX_TEXT_LEN:
                raise ValueError(f"operations[{i}].text exceeds {MAX_TEXT_LEN} chars")
            if where not in ("start", "end"):
                raise ValueError(f"operations[{i}].where must be 'start' or 'end'")
            clean.append({"op": "insert_text", "text": text, "where": where})
        else:  # save
            clean.append({"op": "save"})
    return clean


def summarize_ops(operations: list[dict[str, Any]]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for op in operations:
        counts[op["op"]] = counts.get(op["op"], 0) + 1
    return counts
