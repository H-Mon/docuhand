"""Mimic the client's exact MCP subprocess spawn and capture what the child does."""
import subprocess
import sys
import time

HERMES_ROOT = r"C:\Users\user\AppData\Local\hermes\hermes-agent"
sys.path.insert(0, HERMES_ROOT)

from tools.mcp_tool import _build_safe_env  # noqa: E402

EXE = sys.argv[1] if len(sys.argv) > 1 else r"D:\Repositories\docuhand\.venv\Scripts\pythonw.exe"

env = _build_safe_env(None)
print(f"spawn: {EXE} -m docuhand  (env keys: {len(env)})")
p = subprocess.Popen(
    [EXE, "-m", "docuhand"],
    env=env,
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
)
time.sleep(6)
rc = p.poll()
print(f"poll after 6s: rc={rc}")
if rc is not None:
    err = p.stderr.read().decode("utf-8", "replace")
    print("STDERR:")
    print(err[:3000])
p.kill()
