# DocuHand — one-command first run (PowerShell).
#
# Downloads a legacy .doc with a real table in it and prints the prompt you
# should paste into your AI agent. Nothing is installed, nothing registered.
# Safe to re-run: the demo folder is rebuilt each time.

$ErrorActionPreference = 'Stop'

$Base = 'https://raw.githubusercontent.com/H-Mon/docuhand/main/examples/demo-files'
$DemoDir = Join-Path $env:USERPROFILE 'docuhand-demo'
$Doc = Join-Path $DemoDir 'DocuHand_Demo.doc'

Write-Host "==> Creating demo folder: $DemoDir"
New-Item -ItemType Directory -Force -Path $DemoDir | Out-Null

Write-Host "==> Downloading the demo .doc (a legacy WPS-created binary)"
Invoke-WebRequest -Uri "$Base/DocuHand_Demo.doc" -OutFile $Doc

Write-Host "==> Done."
Write-Host ""
Write-Host "Your demo file is here:"
Write-Host "    $Doc"
Write-Host ""
Write-Host "Now point your AI agent at it - copy this into your agent chat:"
Write-Host ""

$Prompt = @"
I have a legacy .doc file at $Doc
It contains a table.

Please use the docuhand MCP server to:
1. inspect it first (inspect_document) and tell me the real format, page
   count, table count, and whether it is locked or encrypted
2. extract its tables and show me the JSON grid (extract_content)
3. convert it to .docx (convert_documents)
4. fill the DemoBookmark in the converted file with a one-line note
   (fill_template), saving to a new file

Do not modify the original .doc.
"@

Write-Host $Prompt
Write-Host ""
Write-Host "If your agent says the docuhand MCP server is not connected, add this to"
Write-Host "your MCP client config and restart it:"
Write-Host ""
Write-Host '{'
Write-Host '  "mcpServers": {'
Write-Host '    "docuhand": {'
Write-Host '      "command": "uvx",'
Write-Host '      "args": ["docuhand", "serve"]'
Write-Host '    }'
Write-Host '  }'
Write-Host '}'
