"""Tests for excel2docx generator."""

import tempfile
from pathlib import Path

import pytest

from excel2docx.generator import generate, _resolve, _color


class TestResolve:
    def test_nested_dict(self):
        data = {"a": {"b": {"c": 42}}}
        assert _resolve(data, "a.b.c") == 42

    def test_list_index(self):
        data = {"items": [10, 20, 30]}
        assert _resolve(data, "items.1") == 20

    def test_none_on_missing(self):
        assert _resolve({"a": 1}, "b") is None
        assert _resolve({"a": 1}, "a.b.c.d") is None


class TestColor:
    def test_hex_to_rgb(self):
        from docx.shared import RGBColor
        c = _color("#2F5496")
        assert isinstance(c, RGBColor)
        assert c == RGBColor(0x2F, 0x54, 0x96)

    def test_no_hash(self):
        from docx.shared import RGBColor
        c = _color("003366")
        assert c == RGBColor(0, 51, 102)


class TestGenerate:
    def test_minimal_document(self):
        """Generate a minimal docx and verify it exists."""
        report = {}
        template = {
            "page": {"size": "A4", "font": "Calibri", "font_size": 10},
            "title": {"text": "Test Report", "size": 16},
        }

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            out = f.name

        result = generate(report, template, out)
        assert result.exists()
        assert result.stat().st_size > 0
        Path(out).unlink()

    def test_paragraph_with_source(self):
        """Paragraph from data source."""
        report = {"summary": "Hello World"}
        template = {
            "page": {"size": "A4"},
            "title": {"text": "Report"},
            "sections": [{
                "heading": "Overview",
                "elements": [
                    {"type": "paragraph", "text": "{value}", "source": "summary"},
                ],
            }],
        }

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            out = f.name

        result = generate(report, template, out)
        assert result.exists()
        assert result.stat().st_size > 1000  # Should have content
        Path(out).unlink()

    def test_key_value_section(self):
        """Key-value rendering."""
        report = {"sheets": {"logs": {"aggregates": {"Security": 5, "Ops": 3}}}}
        template = {
            "page": {"size": "A4"},
            "title": {"text": "Report"},
            "sections": [{
                "heading": "Breakdown",
                "elements": [
                    {"type": "key_value", "source": "sheets.logs.aggregates"},
                ],
            }],
        }

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            out = f.name

        result = generate(report, template, out)
        assert result.exists()
        Path(out).unlink()

    def test_table_rendering(self):
        """Table from list of dicts."""
        report = {
            "items": [
                {"id": "1", "name": "Alpha", "value": 100},
                {"id": "2", "name": "Beta", "value": 200},
            ],
        }
        template = {
            "page": {"size": "A4"},
            "title": {"text": "Report"},
            "sections": [{
                "heading": "Data",
                "elements": [{
                    "type": "table",
                    "source": "items",
                    "columns": [
                        {"field": "id", "header": "ID"},
                        {"field": "name", "header": "Name"},
                        {"field": "value", "header": "Value"},
                    ],
                }],
            }],
        }

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            out = f.name

        result = generate(report, template, out)
        assert result.exists()
        Path(out).unlink()

    def test_list_rendering(self):
        """List from array."""
        report = {"notes": ["Item 1", "Item 2", "Item 3"]}
        template = {
            "page": {"size": "A4"},
            "title": {"text": "Report"},
            "sections": [{
                "heading": "Notes",
                "elements": [
                    {"type": "list", "source": "notes", "item_format": "- {value}"},
                ],
            }],
        }

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            out = f.name

        result = generate(report, template, out)
        assert result.exists()
        Path(out).unlink()

    def test_metadata_injection(self):
        """Metadata fields appear in document."""
        report = {}
        template = {
            "page": {"size": "A4"},
            "title": {"text": "Report"},
            "meta_fields": [
                {"label": "Generated:", "source": "missing", "default": "N/A"},
            ],
        }

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            out = f.name

        result = generate(report, template, out)
        assert result.exists()
        Path(out).unlink()

    def test_footer(self):
        """Footer with timestamp."""
        report = {}
        template = {
            "page": {"size": "A4"},
            "title": {"text": "Report"},
            "footer": {"text": "Confidential", "include_timestamp": True},
        }

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            out = f.name

        result = generate(report, template, out)
        assert result.exists()
        Path(out).unlink()

    def test_comments_element(self):
        """Comments section rendering."""
        report = {"data": {"comments": ["Note A", "Note B"]}}
        template = {
            "page": {"size": "A4"},
            "title": {"text": "Report"},
            "sections": [{
                "heading": "Comments",
                "elements": [{
                    "type": "comments",
                    "source": "data.comments",
                    "label": "Notes",
                }],
            }],
        }

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            out = f.name

        result = generate(report, template, out)
        assert result.exists()
        Path(out).unlink()

    def test_empty_comments(self):
        """Empty comments show fallback text."""
        report = {"data": {"comments": []}}
        template = {
            "page": {"size": "A4"},
            "title": {"text": "Report"},
            "sections": [{
                "heading": "Comments",
                "elements": [{
                    "type": "comments",
                    "source": "data.comments",
                    "empty_text": "Nothing to report",
                }],
            }],
        }

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            out = f.name

        result = generate(report, template, out)
        assert result.exists()
        Path(out).unlink()
