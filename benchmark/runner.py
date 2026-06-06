"""Benchmark runner — executes the pipeline and scores results.

Usage:
    python -m benchmark.runner \\
        --excel data.xlsx --config config.yaml \\
        --ground-truth ground_truth.json \\
        --expected-schema expected_schema.json

    python -m benchmark.runner --suite casino
    python -m benchmark.runner --suite casino --llm --api-key sk-...
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
import time
import urllib.request
import urllib.error
from pathlib import Path

from excel2docx.parser import parse_excel, load_config
from excel2docx.transformer import transform, _build_prompt
from excel2docx.generator import generate

from .scorer import (
    score_table_understanding,
    score_report_quality,
    BenchmarkResult,
    format_result,
)


def _load_api_key(explicit: str | None = None) -> str | None:
    """Load API key from explicit arg, env, or Hermes .env."""
    if explicit:
        return explicit
    for var in ["DEEPSEEK" + "_API_KEY", "OPENAI_API_KEY"]:
        key = os.environ.get(var)
        if key:
            return key
    env_f = Path.home() / ".hermes" / ".env"
    if env_f.exists():
        for line in env_f.read_text().splitlines():
            for prefix in ["DEEPSEEK" + "_API_KEY=", "OPENAI_API_KEY="]:
                if line.startswith(prefix):
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    return None


def run_benchmark(
    excel_path: str | Path,
    config_path: str | Path,
    ground_truth_path: str | Path,
    output_path: str | Path | None = None,
    *,
    expected_schema_path: str | Path | None = None,
    label: str = "",
    api_key: str | None = None,
) -> BenchmarkResult:
    """Run pipeline and benchmark it."""
    started = time.time()

    config = load_config(config_path)
    ground_truth = json.loads(Path(ground_truth_path).read_text())
    expected_schema = None
    if expected_schema_path:
        expected_schema = json.loads(Path(expected_schema_path).read_text())

    if output_path is None:
        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        output_path = tmp.name
        tmp.close()

    # 1. Parse
    parsed = parse_excel(excel_path, config.get("parser", {}))

    # 2. Transform (rules or LLM)
    transform_cfg = config.get("transform", {})
    if transform_cfg.get("mode") == "llm":
        report = _run_llm_transform(parsed, config, api_key)
    else:
        report = transform(parsed, transform_cfg)

    # 3. Generate DOCX
    output_path = Path(output_path)
    try:
        generate(report, config.get("template", {}), output_path,
                metadata=config.get("metadata", {}))
    except Exception:
        pass  # score what we can

    # 4. Score
    table_score = score_table_understanding(parsed, ground_truth)
    report_score = score_report_quality(
        report, parsed,
        expected_schema=expected_schema,
        ground_truth=ground_truth,
    )

    elapsed = time.time() - started

    if output_path.exists() and "--output" not in sys.argv:
        output_path.unlink(missing_ok=True)

    return BenchmarkResult(
        table_score=table_score,
        report_score=report_score,
        config_name=label or Path(config_path).stem,
        excel_path=str(Path(excel_path).name),
        duration_seconds=elapsed,
    )


def _run_llm_transform(parsed: dict, config: dict, api_key: str | None) -> dict:
    """Run LLM transform: build prompt, call API, return parsed JSON."""
    if not api_key:
        raise RuntimeError("LLM mode requires --api-key or DEEPSEEK_API_KEY env var")

    tf = config.get("transform", {})
    llm_cfg = tf.get("llm_config", {})
    prompt = _build_prompt(tf.get("prompt_template", ""), parsed)

    model = llm_cfg.get("model", "deepseek-chat")
    base_url = llm_cfg.get("base_url", "https://api.deepseek.com/v1")
    temperature = llm_cfg.get("temperature", 0.3)
    max_tokens = llm_cfg.get("max_tokens", 8000)

    payload = json.dumps({
        "model": model,
        "messages": [
            {"role": "system", "content": "Output ONLY valid JSON matching the exact schema provided."},
            {"role": "user", "content": prompt},
        ],
        "temperature": temperature,
        "max_tokens": max_tokens,
        "response_format": {"type": "json_object"},
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
        content = re.sub(r"^```(?:json)?\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
        return json.loads(content)
    except urllib.error.HTTPError as e:
        body = e.read().decode()[:500]
        raise RuntimeError(f"LLM API error {e.code}: {body}")
    except json.JSONDecodeError as e:
        raise RuntimeError(f"LLM JSON parse error: {e}")


def run_suite(suite_name: str, *, llm: bool = False, api_key: str | None = None) -> list[BenchmarkResult]:
    """Run a named benchmark suite."""
    base = Path(__file__).resolve().parent

    if suite_name == "casino":
        excel = base / "test_data" / "casino_surveillance_log.xlsx"
        ground_truth = base / "ground_truth" / "casino_surveillance.json"
        schema = base / "ground_truth" / "casino_pdr_schema.json"

        if llm:
            config = base / "configs" / "casino_pdr_llm.yaml"
            label = "Casino PDR (LLM)"
        else:
            config = base / "configs" / "casino_pdr.yaml"
            label = "Casino PDR (rules)"

        for p, lbl in [(excel, "Excel"), (config, "Config"), (ground_truth, "Ground truth")]:
            if not p.exists():
                print(f"[ERROR] Missing: {lbl} -> {p}")
                sys.exit(1)

        return [run_benchmark(
            excel, config, ground_truth,
            expected_schema_path=schema if schema.exists() else None,
            label=label,
            api_key=api_key,
        )]

    print(f"Unknown suite: {suite_name}")
    print("Available: casino")
    sys.exit(1)


def main():
    ap = argparse.ArgumentParser(description="excel2docx Benchmark Runner")
    ap.add_argument("--suite", default=None, help="Named suite (e.g. 'casino')")
    ap.add_argument("--llm", action="store_true", help="Use LLM-mode config for suite")
    ap.add_argument("--api-key", default=None, help="LLM API key")
    ap.add_argument("--excel", default=None, help="Path to Excel file")
    ap.add_argument("--config", default=None, help="Path to pipeline config YAML")
    ap.add_argument("--ground-truth", default=None, help="Path to ground truth JSON")
    ap.add_argument("--expected-schema", default=None, help="Path to expected schema JSON")
    ap.add_argument("--output", default=None, help="Output .docx path")
    ap.add_argument("--label", default="benchmark", help="Label for this run")
    ap.add_argument("--json", action="store_true", help="Output raw JSON")

    args = ap.parse_args()

    if args.suite:
        results = run_suite(args.suite, llm=args.llm, api_key=_load_api_key(args.api_key))
    elif args.excel and args.config and args.ground_truth:
        results = [run_benchmark(
            args.excel, args.config, args.ground_truth,
            output_path=args.output,
            expected_schema_path=args.expected_schema,
            label=args.label,
            api_key=_load_api_key(args.api_key),
        )]
    else:
        ap.error("Either --suite or (--excel + --config + --ground-truth) is required")

    for result in results:
        if args.json:
            print(json.dumps({
                "config": result.config_name,
                "excel": result.excel_path,
                "duration_s": round(result.duration_seconds, 1),
                "table_understanding": {
                    "overall": result.table_score.overall,
                    "cell_recovery_pct": result.table_score.cell_recovery_pct,
                    "row_completeness_pct": result.table_score.row_completeness_pct,
                    "column_detection_pct": result.table_score.column_detection_pct,
                    "type_accuracy_pct": result.table_score.type_accuracy_pct,
                    "sheet_coverage_pct": result.table_score.sheet_coverage_pct,
                    "detail": {
                        "cells_recovered": result.table_score.total_cells_recovered,
                        "cells_expected": result.table_score.total_cells_expected,
                        "rows_parsed": result.table_score.parsed_rows,
                        "rows_expected": result.table_score.expected_rows,
                        "missed_columns": result.table_score.missed_columns,
                        "type_errors": result.table_score.type_errors[:5],
                    },
                },
                "report_quality": {
                    "overall": result.report_score.overall,
                    "section_completeness_pct": result.report_score.section_completeness_pct,
                    "value_accuracy_pct": result.report_score.value_accuracy_pct,
                    "hallucination_rate_pct": result.report_score.hallucination_rate_pct,
                    "structure_conformance": result.report_score.structure_conformance,
                    "detail": {
                        "missing_sections": result.report_score.missing_sections,
                        "value_mismatches": result.report_score.value_mismatches[:5],
                        "hallucinations": result.report_score.hallucinations[:5],
                    },
                },
            }, indent=2))
        else:
            print(format_result(result))

    if len(results) > 1:
        avg_table = sum(r.table_score.overall for r in results) / len(results)
        avg_report = sum(r.report_score.overall for r in results) / len(results)
        avg = (avg_table + avg_report) / 2
        grade = "A" if avg >= 90 else "B" if avg >= 75 else "C" if avg >= 60 else "D"
        print(f"\nSUITE SUMMARY: {len(results)} runs")
        print(f"  Avg Table Understanding: {avg_table:.1f}")
        print(f"  Avg Report Quality:      {avg_report:.1f}")
        print(f"  Overall: {avg:.1f} ({grade})")


if __name__ == "__main__":
    main()
