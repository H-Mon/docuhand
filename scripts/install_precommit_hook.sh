#!/bin/sh
# Install the pre-commit privacy gate into .git/hooks/.
#
#   bash scripts/install_precommit_hook.sh
#
# The hook file itself is NOT version-controlled (git ignores .git/), so a
# fresh clone has to install it once.  This script is the version-controlled
# part.
set -e

REPO_ROOT="$(git rev-parse --show-toplevel)"
HOOK_DIR="$REPO_ROOT/.git/hooks"
HOOK="$HOOK_DIR/pre-commit"

mkdir -p "$HOOK_DIR"

cat > "$HOOK" <<'HOOK_EOF'
#!/bin/sh
set -e
REPO_ROOT="$(git rev-parse --show-toplevel)"
GATE="$REPO_ROOT/scripts/precommit_gate.py"
[ -f "$GATE" ] || exit 0
PY=""
for cand in python3 python py; do
    command -v "$cand" >/dev/null 2>&1 && { PY="$cand"; break; }
done
[ -n "$PY" ] || exit 0
"$PY" "$GATE" --repo "$REPO_ROOT" "$@"
HOOK_EOF

chmod +x "$HOOK"
echo "Installed: $HOOK"
