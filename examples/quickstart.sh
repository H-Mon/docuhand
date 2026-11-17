#!/usr/bin/env bash
# DocuHand - one-command first run.
#
# Downloads a legacy .doc with a real table in it and prints the prompt you
# should paste into your AI agent. Nothing is installed, nothing registered.
# Safe to re-run: the demo folder is rebuilt each time.
set -euo pipefail

BASE="https://raw.githubusercontent.com/H-Mon/docuhand/main/examples/demo-files"

# Pick the home directory. $USERPROFILE on Windows is a native path with
# backslashes - mixing it with a forward-slash suffix produces
# "C:\Users\x/docuhand-demo", which still works for mkdir but is ugly in the
# prompt we print. Prefer USERPROFILE only if we are on Windows.
case "$(uname -s)" in
    MINGW*|MSYS*|CYGWIN*)
        # cygpath -m gives a native Windows path; cygpath -u gives an MSYS
        # one. curl is a native Windows binary and does NOT understand MSYS
        # paths - it fails with "client returned ERROR on write". So the
        # download path must stay in native form.
        DEMO_DIR_NATIVE="$(cygpath -m "${USERPROFILE:-}")/docuhand-demo"
        DEMO_DIR="$(cygpath -u "${USERPROFILE:-}")/docuhand-demo"
        ;;
    *)
        DEMO_DIR_NATIVE="${HOME:-$(pwd)}/docuhand-demo"
        DEMO_DIR="$DEMO_DIR_NATIVE"
        ;;
esac

DOC="$DEMO_DIR/DocuHand_Demo.doc"
DOC_NATIVE="$DEMO_DIR_NATIVE/DocuHand_Demo.doc"

echo "==> Creating demo folder: $DEMO_DIR"
mkdir -p "$DEMO_DIR"

echo "==> Downloading the demo .doc (a legacy WPS-created binary)"
curl -fsSL "$BASE/DocuHand_Demo.doc" -o "$DOC_NATIVE"

echo "==> Done."
echo
echo "Your demo file is here:"
echo "    $DOC_NATIVE"
echo
echo "Now point your AI agent at it - copy this into your agent chat:"
echo
cat <<EOF
I have a legacy .doc file at $DOC_NATIVE
It contains a table.

Please use the docuhand MCP server to:
1. inspect it first (inspect_document) and tell me the real format, page
   count, table count, and whether it is locked or encrypted
2. extract its tables and show me the JSON grid (extract_content)
3. convert it to .docx (convert_documents)
4. fill the DemoBookmark in the converted file with a one-line note
   (fill_template), saving to a new file

Do not modify the original .doc.
EOF
echo
echo "If your agent says the docuhand MCP server is not connected, add this to"
echo "your MCP client config and restart it:"
echo
cat <<EOF
{
  "mcpServers": {
    "docuhand": {
      "command": "uvx",
      "args": ["docuhand", "serve"]
    }
  }
}
EOF
