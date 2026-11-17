#!/usr/bin/env python
"""Pre-commit gate for a public repository.

Runs on every commit and blocks it when the staged changes would publish
something that cannot be taken back: a real identity, a machine-specific
path, a client detail, a lockfile pointing at a regional mirror.

Design rules, all of them learned the hard way:

* **Only inspect what is staged.**  Scanning the whole worktree is slow and
  produces false alarms from files nobody is about to publish.
* **Separate technical from scenario wording.**  "Chinese-Windows" describes
  the platform the tool supports - it must stay.  "government form" describes
  who the client is - it must go.  A gate that cannot tell them apart trains
  people to disable the gate.
* **Fail loudly, exit non-zero.**  A gate that always passes is decoration.
* **Bypass is possible but visible.**  `--no-verify` still exists, but the
  hook prints what it would have blocked.

Usage (from the repository root):

    python <this file>            # check the staged diff
    python <this file> --all      # also check the whole worktree
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

# --------------------------------------------------------------------------
# Rules
# --------------------------------------------------------------------------

# Pattern -> why it matters.  Order matters only for readability.
BLOCKERS: list[tuple[str, str]] = [
    # A real Windows user directory.  The generic placeholders below are
    # deliberately NOT listed: they are how a reader is told to substitute
    # their own path.
    (r"\bC:[\\/]Users[\\/](?!user[\\/]|users[\\/]|x[\\/]|username[\\/]|yourname[\\/]|your_name[\\/]|YOUR_NAME[\\/]|USERNAME[\\/]|<[^>]+>[\\/])\w+",
     "absolute path containing a real Windows user name"),

    # Hermes' own scratch/venv paths - a development-environment trace.
    (r"AppData[\\/]Local[\\/]hermes",
     "development-environment path (hermes scratch/venv)"),

    # Internal planning documents that must never be public.
    # NOTE: no \b here - CJK has no word boundaries, so \b never matches
    # next to a Chinese character.  Substring matching is the correct tool.
    ("开工包", "internal planning document name"),
    ("干部任免", "client's real document name"),
    ("干部人事档案", "client's real document name"),

    # A regional mirror written into a committed lockfile.  The canonical
    # hosts are pypi.org (index) and files.pythonhosted.org (artefacts) -
    # anything else on those domains is a mirror.
    (r"pypi\.(?!org\b)[a-z0-9.-]+", "non-canonical PyPI mirror in a published file"),
    (r"files\.pythonhosted\.(?!org\b)[a-z0-9.-]+", "non-canonical package host"),
    (r"https?://(?!files\.pythonhosted\.org/)[a-z0-9.-]*pythonhosted\.[a-z]+/",
     "non-canonical package host"),

    # Live credentials.
    (r"\b[pP]yPI[_ -]?(?:token|password)\s*[:=]\s*['\"]?[A-Za-z0-9_\-]{16,}",
     "possible PyPI credential"),
    (r"\bpypi-[A-Za-z0-9_\-]{16,}", "possible PyPI API token"),
    (r"\bsk-[A-Za-z0-9]{20,}", "possible API key"),
    (r"\bghp_[A-Za-z0-9]{20,}", "possible GitHub token"),
]

# Terms that describe the CLIENT SECTOR or the person's job.  These are the
# ones that leak.  Phrases are matched case-insensitively.
SCENARIO_TERMS: list[tuple[str, str]] = [
    (r"government form", "names the client's sector"),
    (r"chinese government", "names the client's sector"),
    (r"government/education", "names the client's sector"),
    (r"appointment-and-dismissal", "names the client's document type"),
    (r"\bpublic sector\b", "names the client's sector"),
    (r"\bcivil servant", "names the client's sector"),
    # NOTE: no \b here - CJK has no word boundaries, so \b never matches
    # next to a Chinese character.  Substring matching is the correct tool.
    ("人事档案", "names the client's sector"),
    ("政务", "names the client's sector"),

    # Job-role wording.  These do not name a sector directly, but combined
    # with the rest of the story they reconstruct it.
    (r"\bHer job is\b", "describes the person's job role"),
    (r"\bshe inherits\b", "describes the person's job role"),
    (r"\binherits a batch\b", "describes the person's job role"),
    (r"\bher employer\b", "describes the person's job role"),
]

# Terms that MUST survive.  Used only to report, never to block: if a future
# rule starts matching these, the gate needs adjusting.
TECHNICAL_TERMS = [
    "Chinese-Windows", "CJK", "GBK", "East-Asian", "NameFarEast",
    "SimSun", "Hangul Word Processor",
]

# Files that legitimately contain absolute paths (they are the tools that
# operate on them) or nothing worth scanning.
SKIP_SUFFIXES = {".pyc", ".pyo", ".png", ".jpg", ".jpeg", ".gif", ".ico",
                  ".whl", ".tar.gz", ".zip", ".pdf"}

# The gate itself contains every string it blocks, because a rule has to
# name the thing it forbids.  Without this exemption the gate blocks its own
# commit and the first person to install it concludes the gate is broken.
SELF_EXEMPT = {"scripts/precommit_gate.py", "scripts/install_precommit_hook.sh"}

# Any file whose job is to define or test the rules necessarily contains the
# strings the rules forbid - a rule has to spell out the thing it blocks.
# Exempt by pattern rather than by an explicit list, so that adding a new
# gate script does not require editing this one (and does not block its own
# first commit, which is how a gate gets declared "broken").
SELF_EXEMPT_PATTERNS = (
    "precommit_gate.py",
    "publish_gate.py",
    "install_precommit_hook.sh",
)


def is_exempt(rel: str) -> bool:
    norm = rel.replace("\\", "/")
    name = Path(norm).name
    return (norm in SELF_EXEMPT
            or name in SELF_EXEMPT
            or any(name == pat for pat in SELF_EXEMPT_PATTERNS))


def run(cmd: list[str], cwd: Path) -> tuple[int, str]:
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                       errors="replace")
    return p.returncode, (p.stdout or "") + (p.stderr or "")


def staged_files(repo: Path) -> list[str]:
    rc, out = run(["git", "diff", "--cached", "--name-only", "--diff-filter=ACMR"], repo)
    if rc != 0:
        return []
    return [ln for ln in out.splitlines() if ln.strip()]


def staged_patch_by_file(repo: Path) -> dict[str, str]:
    """The ADDED lines of the staged diff, grouped per file.

    Per-file grouping matters because the gate exempts itself: a rule has
    to name the string it forbids, so the rule file necessarily contains it.
    """
    rc, out = run(["git", "diff", "--cached", "--unified=0"], repo)
    if rc != 0:
        return {}
    per_file: dict[str, list[str]] = {}
    current = None
    for line in out.splitlines():
        if line.startswith("+++ b/"):
            current = line[6:].strip()
            per_file.setdefault(current, [])
        elif line.startswith("+++ "):
            current = None
        elif line.startswith("+") and not line.startswith("+++"):
            if current is not None:
                per_file[current].append(line[1:])
    return {k: "\n".join(v) for k, v in per_file.items() if v}


def scan_text(text: str, label: str) -> list[str]:
    findings: list[str] = []
    for pattern, why in BLOCKERS:
        for m in re.finditer(pattern, text):
            snippet = m.group(0)
            if len(snippet) > 70:
                snippet = snippet[:70] + "..."
            findings.append(f"{label}: {why}\n    {snippet}")
    for pattern, why in SCENARIO_TERMS:
        for m in re.finditer(pattern, text, re.I):
            findings.append(f"{label}: {why}\n    {m.group(0)}")
    return findings


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", default=".")
    ap.add_argument("--all", action="store_true",
                    help="also scan every tracked file in the worktree")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    repo = Path(args.repo).resolve()
    if not (repo / ".git").exists():
        print(f"not a git repository: {repo}", file=sys.stderr)
        return 2

    findings: list[str] = []

    # --- 1. the staged diff (what this commit would publish) ---
    exempt_names = {Path(s).name for s in SELF_EXEMPT}
    staged = staged_patch_by_file(repo)
    if not staged:
        print("no staged changes - nothing to check.")
    for rel, added in staged.items():
        if is_exempt(rel):
            continue
        findings.extend(scan_text(added, f"staged:{rel}"))

    # --- 2. the whole worktree, optionally ---
    if args.all:
        for rel in run(["git", "ls-files"], repo)[1].splitlines():
            rel = rel.strip()
            if not rel:
                continue
            if Path(rel).suffix.lower() in SKIP_SUFFIXES:
                continue
            if is_exempt(rel):
                continue
            fp = repo / rel
            if not fp.is_file():
                continue
            try:
                content = fp.read_text(encoding="utf-8")
            except (UnicodeDecodeError, OSError):
                continue
            findings.extend(scan_text(content, rel))

    if not findings:
        if not args.quiet:
            print("pre-commit gate: clean.")
        return 0

    print()
    print("=" * 70)
    print("PRE-COMMIT GATE: BLOCKED")
    print("=" * 70)
    for f in findings:
        print("  " + f)
    print()
    print("These would be published and cannot be withdrawn cleanly later.")
    print("Fix them, or - if a hit is a genuine false positive - add an")
    print("explicit exception to the rules in the gate script.")
    print()
    print("  Note: technical wording (Chinese-Windows, CJK, GBK, East-Asian,")
    print("  NameFarEast) is intentionally allowed.  What is blocked is a real")
    print("  identity, a machine path, a client detail, or a regional mirror.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
