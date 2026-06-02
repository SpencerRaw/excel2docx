"""End-to-end integration test for excel2docx pipeline.

Creates a temp Excel workbook with two sheets, a temp config YAML,
runs the full pipeline, and verifies the output .docx.
"""

import tempfile
from pathlib import Path

import pytest
import yaml

from excel2docx.pipeline import run


@pytest.fixture
def temp_excel():
    """Create a temp Excel with two sheets: Daily Log and Summary."""
    import openpyxl

    wb = openpyxl.Workbook()

    # Sheet 1: Daily Log
    ws1 = wb.active
    ws1.title = "Daily Log"
    ws1.append(["Ref ID", "Category", "Department", "Description"])
    ws1.append(["R001", "Incident", "Security", "Door alarm triggered"])
    ws1.append(["R002", "Report", "Operations", "Shift change complete"])
    ws1.append(["R003", "Incident", "Security", "Visitor badge issued"])
    ws1.append(["R004", "Report", "Operations", "Equipment inspection"])
    ws1.append(["R005", "Incident", "Security", "Unauthorized access attempt"])

    # Sheet 2: Summary
    ws2 = wb.create_sheet("Summary")
    ws2.append(["Metric", "Value"])
    ws2.append(["Total Staff", 42])
    ws2.append(["Open Issues", 3])
    ws2.append(["Compliance Rate", 98.5])

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        wb.save(f.name)
        yield f.name

    Path(f.name).unlink(missing_ok=True)


@pytest.fixture
def temp_config():
    """Create a temp config YAML for the integration test."""
    config = {
        "parser": {
            "sheets": [
                {
                    "sheet_name": "Daily Log",
                    "output_name": "logs",
                    "header_row": 1,
                    "data_start_row": 2,
                    "columns": [
                        {"col": 0, "name": "ref_id", "type": "str"},
                        {"col": 1, "name": "category", "type": "str"},
                        {"col": 2, "name": "department", "type": "str"},
                        {"col": 3, "name": "description", "type": "str"},
                    ],
                    "aggregates": {
                        "group_by": "department",
                        "metric": "count",
                    },
                },
                {
                    "sheet_name": "Summary",
                    "output_name": "summary",
                    "header_row": 1,
                    "data_start_row": 2,
                    "columns": [
                        {"col": 0, "name": "metric", "type": "str"},
                        {"col": 1, "name": "value", "type": "float"},
                    ],
                },
            ]
        },
        "transform": {
            "mode": "rules",
            "rules": [
                {
                    "field": "header.title",
                    "source": "metadata.total_rows",
                    "format": "{value} total log entries",
                },
                {
                    "field": "overview.total_entries",
                    "source": "metadata.total_rows",
                },
                {
                    "field": "overview.summary_text",
                    "source": "sheets.summary.rows.0.metric",
                    "default": "No summary available",
                },
            ],
        },
        "template": {
            "page": {
                "size": "A4",
                "margins": {"top": 2, "bottom": 2, "left": 2.5, "right": 2},
                "font": "Calibri",
                "font_size": 10,
            },
            "title": {
                "text": "Integration Test Report",
                "size": 18,
                "color": "#003366",
            },
            "subtitle": {
                "text": "Automated Summary",
                "size": 13,
                "color": "#003366",
            },
            "meta_fields": [
                {"label": "Date:", "source": "header.date", "default": "N/A"},
                {"label": "Total Entries:", "source": "overview.total_entries"},
            ],
            "sections": [
                {
                    "heading": "Overview",
                    "elements": [
                        {
                            "type": "paragraph",
                            "text": "{value}",
                            "source": "overview.summary_text",
                        },
                    ],
                },
                {
                    "heading": "Department Breakdown",
                    "elements": [
                        {
                            "type": "key_value",
                            "source": "sheets.logs.aggregates",
                        },
                    ],
                },
                {
                    "heading": "Detailed Log",
                    "elements": [
                        {
                            "type": "table",
                            "source": "sheets.logs.rows",
                            "columns": [
                                {"field": "ref_id", "header": "Ref ID"},
                                {"field": "category", "header": "Category"},
                                {"field": "department", "header": "Department"},
                                {"field": "description", "header": "Description"},
                            ],
                        },
                    ],
                },
            ],
            "footer": {
                "text": "Confidential — For Internal Use Only",
                "include_timestamp": True,
            },
        },
        "metadata": {
            "generated_at": "auto",
            "generator": "excel2docx",
        },
    }

    with tempfile.NamedTemporaryFile(
        suffix=".yaml", mode="w", delete=False, encoding="utf-8"
    ) as f:
        yaml.dump(config, f, default_flow_style=False, allow_unicode=True)
        config_path = f.name

    yield config_path

    Path(config_path).unlink(missing_ok=True)


def test_full_pipeline_end_to_end(temp_excel, temp_config):
    """Run the full pipeline and verify the output .docx."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        output_path = f.name

    try:
        result = run(temp_excel, temp_config, output_path)

        # Check result structure
        assert "output" in result
        assert "stats" in result
        assert "mode" in result
        assert result["mode"] == "rules"

        # Check stats
        stats = result["stats"]
        assert stats["total_rows"] == 8  # 5 log rows + 3 summary rows
        assert "logs" in stats["sheets_parsed"]
        assert "summary" in stats["sheets_parsed"]

        # Check output file
        output_file = Path(result["output"])
        assert output_file.exists()
        assert output_file.suffix == ".docx"

        file_size = output_file.stat().st_size
        # A DOCX with two sheets of data, a table, key_value, title, subtitle,
        # footer, etc. should be at least ~5 KB
        assert file_size > 3000, f"DOCX too small: {file_size} bytes"

    finally:
        Path(output_path).unlink(missing_ok=True)


def test_pipeline_with_debug_mode(temp_excel, temp_config):
    """Run pipeline with debug=True and verify intermediate JSON files."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        output_path = f.name

    try:
        result = run(temp_excel, temp_config, output_path, debug=True)

        # Verify intermediate files
        parsed_json = Path(output_path).with_suffix(".parsed.json")
        transformed_json = Path(output_path).with_suffix(".transformed.json")

        assert parsed_json.exists(), f"Debug parsed JSON missing: {parsed_json}"
        assert transformed_json.exists(), f"Debug transformed JSON missing: {transformed_json}"

        # Verify parsed JSON has expected content
        import json
        parsed = json.loads(parsed_json.read_text())
        assert parsed["metadata"]["total_rows"] == 8

        # Cleanup debug files
        parsed_json.unlink(missing_ok=True)
        transformed_json.unlink(missing_ok=True)

    finally:
        Path(output_path).unlink(missing_ok=True)
