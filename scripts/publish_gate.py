#!/usr/bin/env python
"""Pre-publication gate for an outward-facing article.

An article, once posted, cannot be withdrawn cleanly - comments have been
written, people have linked to it, and the platform's own edit history
keeps the original.  This gate runs BEFORE publishing and blocks the
release on anything that would have to be fixed afterwards.

It checks four things:

1. **Identity / privacy** - real names, real user directories, client
   document or sector names, internal planning documents.
2. **Scenario wording** - phrasing that reveals the client's sector or
   the person's job (this is the one that code scanning never catches).
3. **Technical wording must survive** - reported, so that a later edit
   does not "fix" the wrong thing.
4. **Internal consistency** - the file's own numbers must not contradict
   each other, and it must not carry the private pricing.

Usage:

    python publish_gate.py <article.md> [--strict]

Exit code 0 = clear to publish.  Non-zero = do not publish yet.

`--strict` additionally fails on any advisory warning, not only on
blockers.  Use it for the final run before posting.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

# --------------------------------------------------------------------------
# 1. Hard identity / privacy
# --------------------------------------------------------------------------

BLOCKERS: list[tuple[str, str]] = [
    # A real Windows user directory.  Generic placeholders are allowed.
    (r"\bC:[\\/]Users[\\/](?!user[\\/]|users[\\/]|x[\\/]|username[\\/]|yourname[\\/]|your_name[\\/]|YOUR_NAME[\\/]|<[^>]+>[\\/])\w+",
     "absolute path with a real user name"),
    (r"AppData[\\/]Local[\\/]hermes", "development-environment path"),
    (r"\bD:[\\/]人工智能\b", "private directory path"),
    (r"\bD:[\\/]股市投资\b", "private directory path"),
    (r"\bD:[\\/]摄影天地\b", "private directory path"),

    # Real document / organisation names.
    ("干部任免", "client's real document name"),
    ("干部人事档案", "client's real document name"),
    ("苏州历史文化名城", "client's organisation"),
    ("华润置地", "client's organisation"),
    ("韩君奕", "a real person's name"),
    ("韩少石", "a real person's name"),

    # Internal planning documents.
    ("开工包", "internal planning document"),
    ("路线图-市场方向", "internal planning document"),
    ("Pro功能规格", "internal planning document"),
    ("定价方案", "internal planning document"),
    ("life-blueprint", "internal planning document"),
]

# --------------------------------------------------------------------------
# 2. Scenario wording - reveals the client sector or the person's job
# --------------------------------------------------------------------------

SCENARIO: list[tuple[str, str]] = [
    (r"government form", "names the client's sector"),
    (r"chinese government", "names the client's sector"),
    (r"government/education", "names the client's sector"),
    (r"appointment-and-dismissal", "names the client's document type"),
    (r"\bpublic sector\b", "names the client's sector"),
    (r"\bcivil servant\b", "names the client's sector"),
    (r"人事档案", "names the client's sector"),
    (r"\b政务\b", "names the client's sector"),

    # Job-role wording.  These do not name a sector directly, but combined
    # with the rest of the story they reconstruct it.
    (r"\bHer job is\b", "describes the person's job role"),
    (r"\bshe inherits\b", "describes the person's job role"),
    (r"\binherits a batch\b", "describes the person's job role"),
    (r"\bher employer\b", "describes the person's job role"),
]

# --------------------------------------------------------------------------
# 3. Technical wording that must stay (reported, never blocks)
# --------------------------------------------------------------------------

KEEP = [
    "Chinese-Windows", "CJK", "GBK", "East-Asian", "NameFarEast",
    "SimSun", "宋体", "方正宋三简体", "Hangul Word Processor",
    "OLE2", "python-docx", "COM",
]

# --------------------------------------------------------------------------
# 4. Private pricing must not be public
# --------------------------------------------------------------------------

PRIVATE_PRICING = ["$39", "$99", "$799", "$29", "$79", "$499",
                   "39/month", "99/month", "799/year"]


def scan(text: str) -> tuple[list[str], list[str]]:
    blockers: list[str] = []
    warnings: list[str] = []

    for pattern, why in BLOCKERS:
        for m in re.finditer(pattern, text, re.I):
            blockers.append(f"{why}: {m.group(0)[:70]}")

    for pattern, why in SCENARIO:
        for m in re.finditer(pattern, text, re.I):
            blockers.append(f"{why}: {m.group(0)[:70]}")

    for price in PRIVATE_PRICING:
        if price in text:
            blockers.append(f"private pricing leaked into a public article: {price}")

    # Consistency: collect every quantity word and look for a self-contradiction
    # pattern (a specific count that appears with two different totals).
    # Only numeric counts can contradict each other.  "fifteen files" and
    # "Fifteen files" are the same sentence; a word-count regex must not
    # treat them as a conflict.
    file_counts = set()
    for m in re.finditer(r"\b(\d+|one|two|three|four|five|six|seven|eight|nine|ten|"
                         r"eleven|twelve|thirteen|fourteen|fifteen|twenty|thirty|"
                         r"forty|fifty)\s+files\b", text, re.I):
        file_counts.add(m.group(1).lower())
    if len(file_counts) > 2:
        warnings.append(
            "several different '<n> files' counts appear - check they agree: "
            + ", ".join(sorted(file_counts)))

    # A relative time expression will be false on the day of publication.
    # Only the ones that anchor a claim to a moment.  "the tools I use today"
    # is plain present tense and stays true forever - flagging it would train
    # the author to ignore the gate.
    TIME_ANCHORS = [
        (r"\bI ran this today\b", "anchors a claim to today"),
        (r"\bas of today\b", "anchors a claim to today"),
        (r"\btoday I (?:ran|tried|built|tested)\b", "anchors a claim to today"),
        (r"\bwe (?:ran|shipped|released) it today\b", "anchors a claim to today"),
        (r"\bthis (?:morning|afternoon|evening)\b", "anchors a claim to a moment"),
        (r"\blast night\b", "anchors a claim to a moment"),
        (r"\bone afternoon\b", "anchors a claim to a moment"),
        (r"\byesterday\b", "anchors a claim to a moment"),
    ]
    for pat, why in TIME_ANCHORS:
        for m in re.finditer(pat, text, re.I):
            warnings.append(f"{why}: {m.group(0)!r}")

    return blockers, warnings


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("article")
    ap.add_argument("--strict", action="store_true",
                    help="fail on warnings too (use for the final run)")
    args = ap.parse_args()

    path = Path(args.article)
    if not path.exists():
        print(f"no such file: {path}", file=sys.stderr)
        return 2

    text = path.read_text(encoding="utf-8")
    flat = " ".join(text.split())

    blockers, warnings = scan(flat)

    print("=" * 70)
    print(f"PUBLISH GATE: {path.name}")
    print("=" * 70)

    present = [k for k in KEEP if k in flat]
    print(f"\n  technical wording present (keep): {', '.join(present) if present else '(none)'}")

    words = len(re.findall(r"[A-Za-z0-9'\-]+", text))
    em = text.count("\u2014") / max(words, 1) * 1000
    print(f"  length: {words} words, {len(text.splitlines())} lines, em-dash {em:.1f}/1k")

    if blockers:
        print("\n  BLOCKERS:")
        for b in blockers:
            print(f"    - {b}")
    else:
        print("\n  no blockers.")

    if warnings:
        print("\n  WARNINGS:")
        for w in warnings:
            print(f"    - {w}")

    print()
    if blockers:
        print("RESULT: DO NOT PUBLISH. Fix the blockers above.")
        return 1
    if warnings and args.strict:
        print("RESULT: DO NOT PUBLISH (strict mode - warnings are errors).")
        return 1
    print("RESULT: clear to publish.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
