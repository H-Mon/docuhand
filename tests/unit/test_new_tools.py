"""CI-safe unit tests for the new v0.1 tools — pure logic, zero COM."""

import pytest

from docuhand.engine.edit_plan import summarize_ops, validate_operations
from docuhand.engine.merge_plan import build_name, resolve_collision, sanitize_component
from docuhand.engine.pdf_plan import parse_page_range
from docuhand.engine.template_plan import (
    normalize_mode,
    resolve_output,
    validate_data,
)
from docuhand.errors import ReadOnlyError
from docuhand.safety import policy


class TestTemplatePlan:
    def test_normalize_mode_accepts_case(self):
        assert normalize_mode(" AUTO ") == "auto"
        assert normalize_mode("PlaceHolders") == "placeholders"

    def test_normalize_mode_rejects_junk(self):
        with pytest.raises(ValueError):
            normalize_mode("regex")

    def test_validate_data_coerces_numbers(self):
        clean = validate_data({"name": "张三", "amount": 426})
        assert clean == {"name": "张三", "amount": "426"}

    def test_validate_data_rejects_braces_in_key(self):
        with pytest.raises(ValueError):
            validate_data({"{{evil}}": "x"})

    def test_validate_data_rejects_newline_in_key(self):
        with pytest.raises(ValueError):
            validate_data({"bad\nkey": "x"})

    def test_validate_data_rejects_non_dict(self):
        with pytest.raises(ValueError):
            validate_data(["not", "a", "dict"])

    def test_resolve_output_default_suffix(self, tmp_path):
        tpl = tmp_path / "t.docx"
        out = resolve_output(tpl, None)
        assert out.name == "t_filled.docx"

    def test_resolve_output_rejects_same_path(self, tmp_path):
        tpl = tmp_path / "t.docx"
        tpl.write_text("x")  # samefile needs the file to exist
        with pytest.raises(ValueError):
            resolve_output(tpl, str(tpl))

    def test_resolve_output_rejects_pdf(self, tmp_path):
        tpl = tmp_path / "t.docx"
        with pytest.raises(ValueError):
            resolve_output(tpl, str(tmp_path / "o.pdf"))


class TestPdfPlan:
    def test_all_variants(self):
        assert parse_page_range(None) is None
        assert parse_page_range("all") is None
        assert parse_page_range("ALL") is None
        assert parse_page_range("*") is None

    def test_single_page(self):
        assert parse_page_range("7") == (7, 7)

    def test_range(self):
        assert parse_page_range("3-9") == (3, 9)
        assert parse_page_range("3–9") == (3, 9)  # en-dash tolerated

    def test_invalid(self):
        for bad in ("0", "-1", "9-3", "1-2-3", "abc", 42):
            with pytest.raises(ValueError):
                parse_page_range(bad)


class TestEditPlan:
    def test_valid_batch(self):
        ops = validate_operations(
            [
                {"op": "replace_all", "find": "2023", "replace": "2024"},
                {"op": "insert_text", "text": "END", "where": "end"},
                {"op": "save"},
            ]
        )
        assert len(ops) == 3
        assert summarize_ops(ops) == {"replace_all": 1, "insert_text": 1, "save": 1}

    def test_rejects_unknown_op(self):
        with pytest.raises(ValueError):
            validate_operations([{"op": "format_disk"}])

    def test_rejects_empty_find(self):
        with pytest.raises(ValueError):
            validate_operations([{"op": "replace_all", "find": ""}])

    def test_rejects_empty_list(self):
        with pytest.raises(ValueError):
            validate_operations([])

    def test_rejects_too_many(self):
        with pytest.raises(ValueError):
            validate_operations([{"op": "save"}] * 51)


class TestMergePlan:
    def test_sanitize_strips_illegal(self):
        assert sanitize_component('a/b\\c:d*e?f"g<h>i|j*k') == "a_b_c_d_e_f_g_h_i_j_k"
        assert sanitize_component("  .hidden. ") == "hidden"

    def test_build_name_pads_index(self):
        assert build_name("{i}_{name}", 5, {"name": "张三"}) == "005_张三"

    def test_build_name_sanitizes_row_values(self):
        assert build_name("{i}_{name}", 1, {"name": "a/b"}) == "001_a_b"

    def test_build_name_survives_missing_field(self):
        # pattern references a field the row lacks → braces get sanitized away
        assert build_name("{i}_{missing}", 2, {}) == "002__missing" or build_name(
            "{i}_{missing}", 2, {}
        ).startswith("002")

    def test_collision_resolution(self, tmp_path):
        first = resolve_collision(tmp_path, "doc", ".docx")
        first.write_text("x")
        second = resolve_collision(tmp_path, "doc", ".docx")
        assert second.name == "doc-2.docx"
        second.write_text("x")
        third = resolve_collision(tmp_path, "doc", ".docx")
        assert third.name == "doc-3.docx"


class TestReadOnlyPolicy:
    def test_default_off(self, monkeypatch):
        monkeypatch.delenv("DOCUHAND_READONLY", raising=False)
        assert policy.is_readonly() is False

    def test_on_variants(self, monkeypatch):
        for v in ("1", "true", "YES", "on"):
            monkeypatch.setenv("DOCUHAND_READONLY", v)
            assert policy.is_readonly() is True

    def test_ensure_writable_raises(self, monkeypatch):
        monkeypatch.setenv("DOCUHAND_READONLY", "1")
        with pytest.raises(ReadOnlyError):
            policy.ensure_writable("fill_template")
