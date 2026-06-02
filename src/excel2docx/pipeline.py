"""End-to-end pipeline: Excel → parse → transform → generate DOCX.

Usage:
    # CLI
    excel2docx --excel data.xlsx --config config.yaml --output report.docx

    # Python
    from excel2docx import pipeline
    pipeline.run("data.xlsx", "config.yaml", "report.docx")
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Callable

from .parser import parse_excel, load_config
from .transformer import transform
from .generator import generate


def run(
    excel_path: str | Path,
    config_path: str | Path,
    output_path: str | Path,
    *,
    llm_client: Callable | None = None,
    debug: bool = False,
) -> dict[str, Any]:
    """Run the full pipeline.

    Args:
        excel_path: Path to Excel workbook
        config_path: Path to pipeline config (YAML/JSON)
        output_path: Path for output .docx
        llm_client: Optional callable for LLM mode transforms.
            Signature: llm_client(prompt: str, response_schema: dict, **kwargs) -> str|dict
        debug: If True, save intermediate JSON files

    Returns:
        Pipeline result with stats and output path
    """
    config = load_config(config_path)
    excel_path = Path(excel_path)
    output_path = Path(output_path)

    # 1. Parse
    print(f"[1/3] Parsing Excel: {excel_path}")
    parsed = parse_excel(excel_path, config.get("parser", {}))
    print(f"      {parsed['metadata']['total_rows']} rows across {len(parsed['metadata']['sheets_parsed'])} sheet(s)")

    if debug:
        jp = output_path.with_suffix(".parsed.json")
        jp.write_text(json.dumps(parsed, indent=2, ensure_ascii=False, default=str))
        print(f"      Debug: {jp}")

    # 2. Transform
    print("[2/3] Transforming data...")
    transform_cfg = config.get("transform", {})

    if transform_cfg.get("mode") == "llm":
        if llm_client is None:
            raise ValueError("LLM mode requires an llm_client. Pass it via run(..., llm_client=...)")

    report = transform(parsed, transform_cfg, llm_client=llm_client)

    if debug:
        jp = output_path.with_suffix(".transformed.json")
        jp.write_text(json.dumps(report, indent=2, ensure_ascii=False, default=str))
        print(f"      Debug: {jp}")

    # 3. Generate DOCX
    print("[3/3] Generating Word document...")
    meta = config.get("metadata", {})
    meta.setdefault("generator", "excel2docx")
    result_path = generate(report, config.get("template", {}), output_path, metadata=meta)
    print(f"      Done! → {result_path}")

    return {
        "output": str(result_path),
        "stats": parsed["metadata"],
        "mode": transform_cfg.get("mode", "rules"),
    }


SAMPLE_CONFIG_YAML = r"""# =============================================================================
# excel2docx Pipeline Configuration
# =============================================================================
# This is a well-commented template. Edit it to match your Excel workbook.
#
# Usage:
#   excel2docx --excel data.xlsx --config this_config.yaml --output report.docx
#   excel2docx --excel data.xlsx --config this_config.yaml --output report.docx --debug
#
# Pipeline stages:
#   1. Parser   — reads Excel sheets according to schema below
#   2. Transform — rule-based field mapping (or LLM-powered narrative generation)
#   3. Generator — produces formatted .docx from structured data + template
# =============================================================================

# ── Parser: define which sheets/columns to read ─────────────────────────────
parser:
  sheets:
    # Each entry defines one sheet to read.
    - sheet_name: "Sheet1"          # Exact or partial sheet name (fuzzy match)
      output_name: "data"           # Name to use in transform references
      header_row: 1                 # Row number (1-indexed) containing headers
      data_start_row: 2             # First data row
      # data_end_row: 100           # Optional: last data row (omit to scan all)
      # skip_empty_rows: true       # Skip rows where all cells are empty
      # skip_pattern: "^TOTAL"      # Optional: regex to skip rows
      columns:
        # Column mapping — three ways to reference columns:
        # 1. By zero-based index
        - col: 0
          name: "id"                # Output field name
          type: str                 # str, int, float, date
          # required: true          # Optional: flag required fields
          # aliases: ["Identifier", "ID"]  # Optional: alternate header names
        - col: 1
          name: "category"
          type: str
        - col: 2
          name: "amount"
          type: float
        # 2. By column letter
        # - col: "D"
        #   name: "notes"
        #   type: str
        # 3. By header name matching (uses aliases)
        # - name: "description"
        #   type: str
        #   aliases: ["Description", "Details", "Notes"]
      aggregates:                   # Optional: compute aggregations
        group_by: "category"        # Field to group by
        metric: count               # count | sum(amount) | avg(amount)

    # Example second sheet:
    # - sheet_name: "Summary"
    #   output_name: "summary"
    #   header_row: 1
    #   data_start_row: 2
    #   columns:
    #     - col: 0
    #       name: "metric"
    #       type: str
    #     - col: 1
    #       name: "value"
    #       type: float

# ── Transform: map parsed data to report fields ─────────────────────────────
transform:
  mode: rules                      # "rules" (mechanical) or "llm" (LLM-powered)

  # --- Rules mode: mechanical field mapping (no LLM needed) ---
  rules:
    - field: "header.title"        # Dot-path output field
      source: "metadata.total_rows" # Dot-path to parsed data
      # format: "{value} entries"  # Optional: Python format string
      # default: "0"               # Optional: fallback if source is missing

    - field: "overview.count"
      source: "metadata.total_rows"

    # - field: "overview.summary"
    #   source: "sheets.summary.rows.0.metric"
    #   default: "No summary available"

  # --- LLM mode: send data to an LLM for narrative generation ---
  # mode: llm
  # prompt_template: |
  #   Generate a business report from the following data:
  #   {{DATA}}
  #
  #   Format the output as:
  #   - An executive summary paragraph
  #   - Key findings as a list
  #   - Recommendations
  # llm_config:
  #   model: gpt-4o
  #   temperature: 0.3
  #   max_tokens: 4000
  #   base_url: https://api.openai.com/v1   # For OpenAI-compatible APIs
  # output_schema:
  #   type: object
  #   properties:
  #     executive_summary: {type: string}
  #     key_findings: {type: array, items: {type: string}}
  #     recommendations: {type: array, items: {type: string}}

# ── Template: DOCX layout and formatting ────────────────────────────────────
template:
  page:
    size: A4                       # A4 or letter
    margins: {top: 2, bottom: 2, left: 2.5, right: 2}  # cm
    font: Calibri
    font_size: 10

  title:
    text: "Daily Operations Report"
    size: 18
    color: "#003366"
    alignment: center              # center, left, right

  # subtitle:
  #   text: "Automated Summary"
  #   size: 13
  #   color: "#003366"

  # meta_fields:                   # Key-value metadata block
  #   - label: "Date:"
  #     source: "header.date"
  #   - label: "Total:"
  #     source: "overview.count"

  sections:
    - heading: "Overview"
      # heading_style: styled-header  # styled-header (dark bg) or plain
      # heading_color: "#2F5496"
      elements:
        - type: paragraph
          text: "Report overview goes here."

        # - type: paragraph
        #   text: "Total entries: {value}"
        #   source: "overview.count"

        # - type: key_value         # Renders dict as key: value pairs
        #   source: "sheets.data.aggregates"

        # - type: table             # Renders list[dict] as a table
        #   source: "sheets.data.rows"
        #   columns:
        #     - {field: "id", header: "ID", width_pct: 20}
        #     - {field: "category", header: "Category", width_pct: 40}
        #     - {field: "amount", header: "Amount", width_pct: 40}

        # - type: list              # Renders list as bullet points
        #   source: "some_list_field"
        #   item_format: "- {value}"

        # - type: comments          # Renders a list with a label
        #   source: "data.comments"
        #   label: "Notes"
        #   empty_text: "Nothing to report"

  footer:
    text: "Confidential — For Internal Use Only"
    include_timestamp: true

# ── Metadata: injected into template placeholders ───────────────────────────
metadata:
  generated_at: "auto"
  generator: "excel2docx"
"""


def main():
    """CLI entry point."""
    ap = argparse.ArgumentParser(
        description="excel2docx — General-purpose Excel-to-DOCX report pipeline"
    )

    # Mutually exclusive: either --init OR the pipeline args
    group = ap.add_mutually_exclusive_group()
    group.add_argument(
        "--init", action="store_true",
        help="Generate a sample config YAML and exit. Use --output to write to a file."
    )
    # These are required unless --init
    ap.add_argument("--excel", default=None, help="Path to Excel workbook")
    ap.add_argument("--config", default=None, help="Path to pipeline config (YAML/JSON)")
    ap.add_argument("--output", default=None, help="Output .docx path (or output path for --init config)")
    ap.add_argument("--debug", action="store_true", help="Save intermediate JSON files")
    ap.add_argument("--api-key", default=None, help="LLM API key (if using LLM mode)")

    args = ap.parse_args()

    # Handle --init
    if args.init:
        if args.output:
            out_path = Path(args.output)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(SAMPLE_CONFIG_YAML)
            print(f"Sample config written to: {out_path}")
        else:
            print(SAMPLE_CONFIG_YAML)
        return 0

    # Validate pipeline args
    if not args.excel or not args.config or not args.output:
        ap.error("--excel, --config, and --output are required (or use --init to generate a config)")

    # LLM client factory (if config requires it)
    config = load_config(args.config)
    transform_cfg = config.get("transform", {})
    llm_client = None

    if transform_cfg.get("mode") == "llm":
        llm_cfg = transform_cfg.get("llm_config", {})
        api_key = args.api_key or _load_llm_key()

        # Generic OpenAI-compatible client
        import urllib.request, urllib.error

        def _llm_client(prompt: str, response_schema: dict | None = None, **kwargs) -> dict:
            """Generic LLM client — works with any OpenAI-compatible API."""
            model = kwargs.get("model", llm_cfg.get("model", "gpt-4o"))
            temperature = kwargs.get("temperature", llm_cfg.get("temperature", 0.3))
            max_tokens = kwargs.get("max_tokens", llm_cfg.get("max_tokens", 4000))
            base_url = kwargs.get("base_url", llm_cfg.get("base_url", "https://api.openai.com/v1"))

            payload = json.dumps({
                "model": model,
                "messages": [
                    {"role": "system", "content": "You are a structured report generator. Output ONLY valid JSON."},
                    {"role": "user", "content": prompt},
                ],
                "temperature": temperature,
                "max_tokens": max_tokens,
                **({"response_format": {"type": "json_object"}} if response_schema else {}),
            }).encode()

            req = urllib.request.Request(
                f"{base_url}/chat/completions",
                data=payload,
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {api_key}",
                },
            )
            try:
                with urllib.request.urlopen(req, timeout=120) as resp:
                    result = json.loads(resp.read())
                content = result["choices"][0]["message"]["content"].strip()
                import re
                content = re.sub(r"^```(?:json)?\s*", "", content)
                content = re.sub(r"\s*```$", "", content)
                return json.loads(content)
            except urllib.error.HTTPError as e:
                body = e.read().decode()[:500]
                raise RuntimeError(f"LLM API error {e.code}: {body}")

        llm_client = _llm_client

    result = run(args.excel, args.config, args.output, llm_client=llm_client, debug=args.debug)
    print(f"\nSummary: {result['stats']['total_rows']} rows → {result['output']} ({result['mode']} mode)")
    return 0


def _load_llm_key() -> str:
    """Load LLM API key from environment."""
    import os
    for var in ["OPENAI_API_KEY", "DEEPSEEK" + "_API_KEY", "ANTHROPIC_API_KEY"]:
        key = os.environ.get(var)
        if key:
            return key
    # Try Hermes .env
    env_f = Path.home() / ".hermes" / ".env"
    if env_f.exists():
        for line in env_f.read_text().splitlines():
            for prefix in ["OPENAI_API_KEY=", "DEEPSEEK" + "_API_KEY="]:
                if line.startswith(prefix):
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    raise RuntimeError("No LLM API key found. Set OPENAI_API_KEY or DEEPSEEK_API_KEY.")


if __name__ == "__main__":
    raise SystemExit(main())
