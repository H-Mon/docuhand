"""MCP-facing tools: pydantic-checked args + LLM-facing behaviour.

v0.1 scope: inspect_document only. Each tool module pairs the safety
wrappers (allowlist/audit) with the engine call and returns structured
dicts — errors included (architecture decision #3: errors are data the
agent reads to self-correct, not tracebacks it drowns in).
"""
