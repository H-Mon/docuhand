"""The reported version must match the packaged version.

A release once shipped `0.1.0` on PyPI whose wheel contained
`__version__ = "0.1.0.dev1"`, because the version literal was duplicated
in pyproject.toml and in the package source and only one of them was
updated. `pip show` reported the right version (it reads the metadata);
`import docuhand; docuhand.__version__` reported the wrong one.

These tests make the duplication impossible to reintroduce silently.
"""

from __future__ import annotations

import re
import tomllib
from pathlib import Path

import docuhand

ROOT = Path(__file__).resolve().parents[2]


def _declared_version() -> str:
    pyproject = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    return pyproject["project"]["version"]


def test_reported_version_matches_pyproject() -> None:
    assert docuhand.__version__ == _declared_version()


def test_source_does_not_hardcode_a_stale_version() -> None:
    """No literal version in the package source.

    The version comes from the installed distribution metadata. A literal
    here is exactly what let the two drift apart.
    """
    source = (ROOT / "src" / "docuhand" / "__init__.py").read_text(encoding="utf-8")
    literals = re.findall(r'__version__\s*=\s*["\']([^"\']+)["\']', source)
    assert not literals, f"hardcoded version literal(s) found: {literals}"


def test_reported_version_looks_like_a_version() -> None:
    assert re.fullmatch(r"\d+\.\d+\.\d+([.-]?(dev|rc|post)\d+)?", docuhand.__version__), (
        f"unexpected version string: {docuhand.__version__!r}"
    )
