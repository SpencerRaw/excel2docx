"""Tests for excel2docx transformer."""

import pytest
from excel2docx.transformer import transform, _resolve_source, _build_prompt


class TestResolveSource:
    def test_simple_key(self):
        data = {"a": 1, "b": {"c": 2}}
        assert _resolve_source(data, "a") == 1
        assert _resolve_source(data, "b.c") == 2

    def test_nested_list(self):
        data = {"sheets": {"main": {"rows": [{"name": "Alice"}, {"name": "Bob"}]}}}
        assert _resolve_source(data, "sheets.main.rows.0.name") == "Alice"
        assert _resolve_source(data, "sheets.main.rows.1.name") == "Bob"

    def test_missing_key(self):
        assert _resolve_source({"a": 1}, "b") is None
        assert _resolve_source({"a": 1}, "a.b.c") is None

    def test_index_out_of_range(self):
        assert _resolve_source({"items": [1]}, "items.5") is None


class TestBuildPrompt:
    def test_substitutes_data_json(self):
        template = "Data: {{DATA}}"
        data = {"sheets": {"main": {"rows": [{"x": 1}]}}}
        result = _build_prompt(template, data)
        assert '"x": 1' in result
        assert "{{DATA}}" not in result

    def test_substitutes_sheet_rows(self):
        template = "Rows: {{main.rows}}"
        data = {"sheets": {"main": {"rows": [{"x": 1}]}}}
        result = _build_prompt(template, data)
        assert '"x": 1' in result
        assert "{{main.rows}}" not in result


class TestTransformRules:
    def test_simple_field_mapping(self):
        parsed = {"metadata": {"total_rows": 42}}
        config = {
            "mode": "rules",
            "rules": [
                {"field": "overview.count", "source": "metadata.total_rows"},
            ],
        }
        result = transform(parsed, config)
        assert result["overview"]["count"] == 42

    def test_nested_fields(self):
        parsed = {"data": {"name": "test"}}
        config = {
            "mode": "rules",
            "rules": [
                {"field": "report.header.title", "source": "data.name"},
            ],
        }
        result = transform(parsed, config)
        assert result["report"]["header"]["title"] == "test"

    def test_default_value(self):
        parsed = {"data": {}}
        config = {
            "mode": "rules",
            "rules": [
                {"field": "field", "source": "data.missing", "default": "N/A"},
            ],
        }
        result = transform(parsed, config)
        assert result["field"] == "N/A"

    def test_format_template(self):
        parsed = {"val": 123}
        config = {
            "mode": "rules",
            "rules": [
                {"field": "display", "source": "val", "format": "${value:,}"},
            ],
        }
        result = transform(parsed, config)
        assert result["display"] == "$123"

    def test_empty_rules(self):
        result = transform({"a": 1}, {"mode": "rules", "rules": []})
        assert result == {}


class TestTransformLLMErrors:
    def test_llm_mode_requires_client(self):
        parsed = {"data": {}}
        config = {"mode": "llm", "prompt_template": "test"}
        with pytest.raises(ValueError, match="llm_client"):
            transform(parsed, config, llm_client=None)

    def test_unknown_mode(self):
        with pytest.raises(ValueError, match="Unknown transform mode"):
            transform({}, {"mode": "invalid"})
