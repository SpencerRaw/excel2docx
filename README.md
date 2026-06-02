# excel2docx — General-Purpose Excel-to-DOCX Pipeline

[![Tests](https://github.com/SpencerRaw/excel2docx/actions/workflows/ci.yml/badge.svg)](https://github.com/SpencerRaw/excel2docx/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9+-blue)](https://pypi.org)

> **Turn any structured Excel workbook into a formatted Word report.**
> Config-driven. Domain-agnostic. LLM optional.

---

## What This Is

A three-stage pipeline:

```
┌──────────┐      ┌──────────────┐      ┌────────────────┐
│  EXCEL   │ ──→  │  TRANSFORM   │ ──→  │     DOCX       │
│ (parser) │      │ (rules/LLM)  │      │ (generator)    │
└──────────┘      └──────────────┘      └────────────────┘
     ↑                   ↑                    ↑
  Schema config     Rule mappings        Template config
  (YAML/JSON)       or LLM prompt        (sections, tables)
```

No hardcoded domain logic. You define:
1. **Parser config**: which sheets, which columns, what types
2. **Transform config**: rule-based field mapping OR LLM-powered narrative
3. **Template config**: document structure (headings, tables, paragraphs)

---

## Quick Start

```bash
pip install excel2docx

# Run with a config
excel2docx --excel data.xlsx --config report.yaml --output report.docx
```

```python
from excel2docx import pipeline

pipeline.run("data.xlsx", "report.yaml", "output.docx")
```

---

## Pipeline Config Example

```yaml
parser:
  sheets:
    - sheet_name: "Daily Log"
      header_row: 2
      data_start_row: 3
      columns:
        - {col: 0, name: ref_id, type: str}
        - {col: 2, name: department, type: str}
      aggregates:
        group_by: department
        metric: count

transform:
  mode: rules  # or "llm" for AI-powered narrative
  rules:
    - field: overview.total
      source: metadata.total_rows

template:
  title: {text: "Daily Report", size: 18}
  sections:
    - heading: "Department Breakdown"
      elements:
        - {type: key_value, source: "sheets.Daily Log.aggregates"}
```

See `examples/business_daily_report.yaml` for a complete example.

---

## Two Modes

### Rule-Based (`mode: rules`)
- Mechanical field mapping
- No API calls, no cost
- Fast, deterministic
- Best for: structured data → structured report

### LLM-Powered (`mode: llm`)
- Send parsed data to LLM for narrative generation
- Works with any OpenAI-compatible API (OpenAI, DeepSeek, local LLMs)
- Context-aware summarization and commentary
- Best for: data → human-readable narrative

```bash
export OPENAI_API_KEY="sk-..."
excel2docx --excel data.xlsx --config llm_report.yaml --output report.docx
```

---

## Use Cases

- **Business**: daily operations reports, sales summaries, KPI dashboards
- **Academic**: experiment logs → formatted lab reports
- **Finance**: transaction logs → audit-ready summaries
- **Healthcare**: patient records → clinical summaries
- **Operations**: shift logs, incident reports, inventory audits

---

## Development

```bash
git clone https://github.com/SpencerRaw/excel2docx.git
cd excel2docx
pip install -e ".[dev]"
pytest
```

---

## License

MIT — SpencerRaw 2026
