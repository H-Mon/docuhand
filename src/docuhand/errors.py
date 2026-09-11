"""Structured error model — architecture decision #3.

Errors are read by *LLMs*, not just humans. Every DocuHand error carries:
- ``error_code``: stable machine token the agent can branch on
- ``human_message``: what a person would want to see
- ``llm_hint``: what the agent should *do next* to self-correct

Tool implementations raise :class:`DocuhandError` subclasses; the server
layer converts them into structured result dicts so the agent can read and
recover from them instead of staring at a raw COM traceback.
"""

from __future__ import annotations

from typing import Any


class DocuhandError(Exception):
    error_code = "DOCUHAND_ERROR"  # never used directly; subclasses override

    def __init__(
        self,
        human_message: str,
        *,
        llm_hint: str = "",
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(human_message)
        self.human_message = human_message
        self.llm_hint = llm_hint
        self.details = details or {}

    def to_dict(self) -> dict[str, Any]:
        return {
            "ok": False,
            "error": {
                "error_code": self.error_code,
                "human_message": self.human_message,
                "llm_hint": self.llm_hint,
                "details": self.details,
            },
        }


class PathNotAllowedError(DocuhandError):
    error_code = "PATH_NOT_ALLOWED"

    def __init__(self, path: str, allowlist: list[str]) -> None:
        super().__init__(
            f"Path is outside the configured allowlist: {path}",
            llm_hint=(
                "Ask the user to add the directory to the DOCUHAND_ALLOWLIST "
                "environment variable (os.pathsep-separated absolute dirs), "
                "or to pass a path inside an already-allowed directory."
            ),
            details={"path": path, "allowlist": allowlist},
        )


class FileNotFoundError(DocuhandError):
    error_code = "FILE_NOT_FOUND"

    def __init__(self, path: str) -> None:
        super().__init__(
            f"File not found: {path}",
            llm_hint=(
                "Check the spelling of the path and that the drive letter is "
                "correct. Use an absolute Windows path (e.g. D:\\Reports\\a.doc)."
            ),
            details={"path": path},
        )


class FileLockedError(DocuhandError):
    error_code = "FILE_LOCKED"

    def __init__(self, path: str, detail: str = "") -> None:
        super().__init__(
            f"File is locked by another process: {path}" + (f" ({detail})" if detail else ""),
            llm_hint=(
                "The file is probably open in Word/WPS right now. For a "
                "read-only inspection retry once; to *modify* it, ask the "
                "user to close it first (edit_open_document arrives in v0.1)."
            ),
            details={"path": path},
        )


class EngineUnavailableError(DocuhandError):
    error_code = "ENGINE_UNAVAILABLE"

    def __init__(self, attempts: list[dict[str, str]]) -> None:
        super().__init__(
            "No Office engine available (tried Word.Application and KWPS.Application)",
            llm_hint=(
                "Neither Microsoft Word nor WPS Office could be launched via "
                "COM. Tell the user to install either one, or check that the "
                "app is not stuck in a first-run dialog."
            ),
            details={"attempts": attempts},
        )


class OpenFailedError(DocuhandError):
    error_code = "OPEN_FAILED"

    def __init__(self, path: str, engine: str, com_error: str) -> None:
        super().__init__(
            f"{engine} failed to open: {path}",
            llm_hint=(
                "The file may be corrupted, password-protected, or in a "
                "format this engine cannot open. Report is_format_recognized "
                "from the inspection payload to the user and suggest trying "
                "the other engine (DOCUHAND_ENGINE=wps or =word)."
            ),
            details={"path": path, "engine": engine, "com_error": com_error},
        )


class PasswordProtectedError(DocuhandError):
    error_code = "FILE_PASSWORD_PROTECTED"

    def __init__(self, path: str, engine: str) -> None:
        super().__init__(
            f"Document is password-protected, cannot be opened read-only: {path}",
            llm_hint=(
                "Opening it would require the password. Ask the user for the "
                "password — never guess one. No other engine will help; do "
                "not retry."
            ),
            details={"path": path, "engine": engine},
        )


class ComCallTimeoutError(DocuhandError):
    error_code = "COM_TIMEOUT"

    def __init__(self, timeout_s: float) -> None:
        super().__init__(
            f"COM operation timed out after {timeout_s:.0f}s",
            llm_hint=(
                "Word/WPS may be showing a modal dialog (first-run wizard, "
                "activation, recovery prompt). Ask the user to check the "
                "screen, then retry once."
            ),
            details={"timeout_s": timeout_s},
        )


class InternalError(DocuhandError):
    error_code = "INTERNAL_ERROR"
