"""CI-safe unit tests for convert_documents planning logic (zero COM)."""

from pathlib import Path

import pytest

from docuhand.engine import convert_plan
from docuhand.engine.office_app import engine_order_names


class TestNormalizeToFormat:
    def test_plain(self):
        assert convert_plan.normalize_to_format("docx") == "docx"
        assert convert_plan.normalize_to_format("doc") == "doc"

    def test_case_and_dot_tolerant(self):
        assert convert_plan.normalize_to_format(".DOCX") == "docx"
        assert convert_plan.normalize_to_format("  Doc ") == "doc"

    @pytest.mark.parametrize("bad", ["pdf", "", "rtf", "docxx", ".txt"])
    def test_invalid_rejected(self, bad):
        with pytest.raises(ValueError):
            convert_plan.normalize_to_format(bad)


class TestResolveTarget:
    def test_in_place_swaps_extension(self, tmp_path):
        src = tmp_path / "report.doc"
        assert convert_plan.resolve_target(src, "docx") == tmp_path / "report.docx"

    def test_dotted_stem_keeps_its_dots(self, tmp_path):
        src = tmp_path / "2024.05.report.doc"
        assert convert_plan.resolve_target(src, "docx") == tmp_path / "2024.05.report.docx"

    def test_output_dir_relocates(self, tmp_path):
        src = tmp_path / "a" / "report.doc"
        out = tmp_path / "out"
        assert convert_plan.resolve_target(src, "docx", out) == out / "report.docx"

    def test_target_never_equals_source_for_real_conversions(self, tmp_path):
        # invariant that protects every source file from being clobbered.
        # classify_source routes same-format files to already_target, so
        # resolve_target is only ever invoked on genuine extension changes;
        # the engine additionally refuses src==dst as a backstop.
        for fmt in ("doc", "docx"):
            other = "docx" if fmt == "doc" else "doc"
            src = tmp_path / f"x.{other}"
            assert convert_plan.resolve_target(src, fmt) != src


class TestClassifySource:
    @pytest.mark.parametrize(
        ("ext", "fmt", "expected"),
        [
            (".doc", "docx", "convert"),
            (".docx", "doc", "convert"),
            (".docx", "docx", "already_target"),
            (".doc", "doc", "already_target"),
            (".DOC", "docx", "convert"),  # case-insensitive extension
            (".rtf", "docx", "unsupported_source"),
            (".txt", "docx", "unsupported_source"),
            ("", "docx", "unsupported_source"),  # e.g. a file literally named ".doc"
        ],
    )
    def test_matrix(self, ext, fmt, expected):
        assert convert_plan.classify_source(ext, fmt) == expected


class TestEngineOrderNames:
    def test_default_word_first(self, monkeypatch):
        monkeypatch.delenv("DOCUHAND_ENGINE", raising=False)
        assert engine_order_names() == ["word", "wps"]

    def test_env_prefers_wps(self, monkeypatch):
        monkeypatch.setenv("DOCUHAND_ENGINE", "wps")
        assert engine_order_names() == ["wps", "word"]

    def test_env_prefers_word(self, monkeypatch):
        monkeypatch.setenv("DOCUHAND_ENGINE", "word")
        assert engine_order_names() == ["word", "wps"]

    def test_env_garbage_falls_back_to_default(self, monkeypatch):
        monkeypatch.setenv("DOCUHAND_ENGINE", "libreoffice")
        assert engine_order_names() == ["word", "wps"]
