"""Templating engine: fill templates via content controls / bookmarks /
``{{placeholder}}`` tokens — STA-only, same failover contract as the rest
of the engine (docs/pitfalls.md battle scars applied throughout):

- **NameFarEast restore** (production scar #1): replacing ``Range.Text``
  resets East-Asian font names to the theme default. Any value spliced
  into a range whose text contains CJK gets its NameFarEast re-set from a
  pre-read font snapshot, and the whole document gets a final sweep before
  SaveAs2 — because SaveAs2 itself can reset it again.
- **Literal Find, never wildcards**, for ``{{key}}`` tokens: data values
  are user text; a wildcard pass would let ``[a-z]*`` in a value eat the
  document. Keys are validated to exclude braces/newlines anyway.
- Old-engine nets catch AttributeError (pitfall #13): ContentControls and
  Find.Wildcards may not exist on Word 2007 / WPS oddities — degrade, don't die.
"""

from __future__ import annotations

import re
from typing import Any

from ..errors import (
    DocuhandError,
    EngineUnavailableError,
    FileNotFoundError,
    OpenFailedError,
    PasswordProtectedError,
    TemplateFillError,
)
from .container_guard import detect_encryption, validate_container
from .com_utils import (
    _g,
    _looks_encrypted,
    _looks_like_arg_mismatch,
    _short_exc,
    file_claims_encrypted,
)
from .wd_constants import WD_DO_NOT_SAVE_CHANGES
from pathlib import Path

# SaveFormat numbers (pitfall #1 — drive Word by number, never by name)
_CJK_RE = re.compile(r"[\u3000-\u9fff\uf900-\ufaff\uff00-\uffef]")


def _has_cjk(text: str) -> bool:
    return bool(_CJK_RE.search(text or ""))


def _snapshot_font(font: Any) -> dict[str, Any]:
    """Capture the fonts of a range BEFORE any text replacement."""
    return {
        "name": _g(lambda: font.Name),
        "far_east": _g(lambda: font.NameFarEast),
        "size": _g(lambda: font.Size),
        "bold": _g(lambda: font.Bold),
    }


def _restore_font(font: Any, snap: dict[str, Any], value: str) -> None:
    """Re-apply what the replacement may have clobbered.

    Only touch NameFarEast when the spliced value actually contains CJK
    (a pure-Latin value has no far-east glyphs, and re-asserting a CJK font
    name over a Latin-only run is harmless but pointless).
    """
    if snap.get("far_east") and _has_cjk(value):
        try:
            font.NameFarEast = snap["far_east"]
        except Exception:
            pass
    if snap.get("name"):
        try:
            font.Name = snap["name"]
        except Exception:
            pass


def _replace_bookmarks(doc: Any, data: dict[str, str], wanted: list[str]) -> tuple[list[str], list[str]]:
    """Fill bookmarks; returns (filled, missing) key lists."""
    filled: list[str] = []
    missing: list[str] = []
    try:
        count = int(doc.Bookmarks.Count)
    except Exception:
        count = 0
    present: dict[str, Any] = {}
    for i in range(count):
        bm = _g(lambda i=i: doc.Bookmarks.Item(i + 1))
        name = _g(lambda bm=bm: bm.Name)
        if name:
            present[name] = bm
    targets = wanted if wanted else list(present.keys())
    for key in targets:
        if key not in data:
            continue  # a bookmark without data is not an error — leave it
        bm = present.get(key)
        if bm is None:
            missing.append(key)
            continue
        rng = _g(lambda bm=bm: bm.Range)
        if rng is None:
            missing.append(key)
            continue
        value = data[key]
        snap = _snapshot_font(_g(lambda rng=rng: rng.Font, {}) or {})
        try:
            rng.Text = value
        except Exception:
            missing.append(key)
            continue
        _restore_font(_g(lambda rng=rng: rng.Font, {}) or {}, snap, value)
        filled.append(key)
    return filled, missing


def _replace_placeholders(doc: Any, data: dict[str, str]) -> tuple[list[str], list[str]]:
    """Literal {{key}} → value. Returns (filled, unmatched).

    PORTABLE RECIPE (probe-verified on Word-hijack 12.0 and real KWPS):
    ``Find.Execute(Replace:=wdReplaceAll)`` reports True but silently
    replaces NOTHING on this engine family — the word-processor accepts
    the call and drops the Replacement. The two-step works everywhere:
    per iteration take a FRESH full-story Range + Find (ClearFormatting,
    Forward, Wrap=wdFindContinue), Execute(token) with no Replace arg,
    then set ``f.Parent.Text = value`` (Parent IS the matched range).
    Loop until Execute returns False. Never let the data value leak into
    FindText (injection would loop forever on self-matching values).
    """
    filled: list[str] = []
    unmatched: list[str] = []
    for key, value in data.items():
        token = "{{" + key + "}}"
        n = 0
        while n < 500:  # hard cap: a self-replicating value cannot hang us
            find = _g(lambda: doc.Content.Find, None)
            if find is None:
                break
            try:
                find.ClearFormatting()
                find.Forward = True
                find.Wrap = 1  # wdFindContinue — whole story from the top
            except Exception:
                pass
            try:
                hit = find.Execute(token)  # find only — NEVER Replace here
            except Exception:
                break
            if not hit:
                break
            try:
                find.Parent.Text = value  # replace the matched range
            except Exception:
                unmatched.append(key)
                break
            n += 1
        if n:
            filled.append(key)
        else:
            unmatched.append(key)
    return filled, unmatched


def _count_leftover_placeholders(doc: Any) -> list[str]:
    """Scan for {{...}} remnants so the agent learns what did not fill.

    Two-step find (same portable recipe as replacement): fresh Content
    Range per iteration; Parent.Text reads the matched token.
    """
    leftovers: list[str] = []
    for _ in range(50):  # report at most 50 — enough to act on
        find = _g(lambda: doc.Content.Find, None)
        if find is None:
            break
        try:
            find.ClearFormatting()
            find.Forward = True
            find.Wrap = 1
            find.MatchWildcards = True
            hit = find.Execute(r"\{\{*\}\}")
        except Exception:
            # wildcard mode unsupported on this engine — fall back to
            # literal scan of the two most common remnants
            try:
                find2 = doc.Content.Find
                find2.ClearFormatting()
                find2.MatchWildcards = False
                find2.Wrap = 1
                hit = find2.Execute("{{")
            except Exception:
                break
            if not hit:
                break
            token = _g(lambda f2=find2: f2.Parent.Text, "")
            if token:
                leftovers.append(str(token)[:120])
            try:
                find2.Parent.Text = ""  # strip so the loop advances
            except Exception:
                break
            continue
        if not hit:
            break
        token = _g(lambda: find.Parent.Text, "")
        if token:
            leftovers.append(str(token)[:120])
        try:
            find.Parent.Text = ""  # strip so the loop advances
        except Exception:
            break
    return leftovers


def _replace_content_controls(doc: Any, data: dict[str, str]) -> tuple[list[str], list[str]]:
    """Map ContentControl Tag/Title → data key. Optional path (2007 lacks it)."""
    filled: list[str] = []
    unmatched: list[str] = []
    try:
        controls = doc.ContentControls
        count = int(controls.Count)
    except AttributeError:
        return filled, unmatched  # old engine — degrade silently
    except Exception:
        return filled, unmatched
    for i in range(count):
        cc = _g(lambda i=i: controls.Item(i + 1))
        if cc is None:
            continue
        tag = _g(lambda cc=cc: cc.Tag, "") or ""
        title = _g(lambda cc=cc: cc.Title, "") or ""
        key = tag if tag in data else (title if title in data else None)
        if key is None:
            continue
        try:
            # Setting Range.Text keeps the control (and its styling) alive;
            # plain text mapping only in v0.1.
            cc.LockContents = False
            cc.Range.Text = data[key]
            filled.append(key)
        except Exception:
            unmatched.append(key)
    return filled, unmatched


def _final_far_east_sweep(doc: Any) -> None:
    """Whole-document NameFarEast re-assertion before/after SaveAs2.

    The scar that taught us this: replacing
    ``Range.Text`` resets East-Asian font names document-wide, and SaveAs2
    can reset them AGAIN. This sweep re-pins NameFarEast on the story
    ranges that carry text — idempotent, and cheap compared to a
    re-typeset.
    """
    for story_getter in (
        lambda: doc.Content,
        lambda: doc.StoryRanges.Item(1),  # wdMainTextStory
    ):
        rng = _g(story_getter, None)
        if rng is None:
            continue
        font = _g(lambda rng=rng: rng.Font, None)
        if font is None:
            continue
        current = _g(lambda font=font: font.NameFarEast, None)
        if not current:
            # engine lost the name entirely — pin the CJK workhorse default
            for candidate in ("宋体", "SimSun"):
                try:
                    font.NameFarEast = candidate
                    break
                except Exception:
                    continue


class _FillOutcome:
    __slots__ = ("filled", "missing", "controls_ok", "controls_bad", "leftovers", "bookmarks_seen")

    def __init__(self) -> None:
        self.filled: list[str] = []
        self.missing: list[str] = []
        self.controls_ok: list[str] = []
        self.controls_bad: list[str] = []
        self.leftovers: list[str] = []
        self.bookmarks_seen = 0


def _fill_open_document(
    doc: Any,
    data: dict[str, str],
    mode: str,
    bookmark_names: list[str],
) -> _FillOutcome:
    out = _FillOutcome()
    try:
        out.bookmarks_seen = int(_g(lambda: doc.Bookmarks.Count, 0) or 0)
    except Exception:
        pass

    if mode in ("auto", "content_controls"):
        ok, bad = _replace_content_controls(doc, data)
        out.controls_ok, out.controls_bad = ok, bad

    if mode in ("auto", "bookmarks"):
        # only bookmark keys actually present in data participate
        wanted = [n for n in bookmark_names if n in data] if bookmark_names else []
        filled, missing = _replace_bookmarks(doc, data, wanted)
        out.filled.extend(filled)
        out.missing.extend(missing)

    if mode in ("auto", "placeholders"):
        remaining = {k: v for k, v in data.items() if k not in set(out.filled)}
        filled, unmatched = _replace_placeholders(doc, remaining)
        out.filled.extend(filled)
        out.missing.extend(unmatched)

    out.leftovers = _count_leftover_placeholders(doc)
    return out
