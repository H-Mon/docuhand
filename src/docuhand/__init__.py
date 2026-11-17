"""DocuHand - read and edit legacy Word documents through the real engine."""

from __future__ import annotations

__all__ = ["__version__"]


def _resolve_version() -> str:
    """Read the installed distribution's version.

    The literal below is only a fallback for running from a source tree
    that was never installed. Hardcoding the version in two places is
    how a release ships `0.1.0.dev1` inside a wheel called `0.1.0`.
    """
    try:
        from importlib.metadata import PackageNotFoundError, version
    except ImportError:  # pragma: no cover - Python < 3.8
        return "0.0.0.dev0"
    try:
        return version("docuhand")
    except PackageNotFoundError:
        return "0.0.0.dev0"


__version__ = _resolve_version()
