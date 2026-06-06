"""Benchmark runner — executes the pipeline and scores results.

Usage:
    python -m benchmark.runner \\
        --excel data.xlsx --config config.yaml \\
        --ground-truth ground_truth.json \\
        --expected-schema expected_schema.json

    python -m benchmark.runner --suite casino
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

from excel2docx.parser import parse_excel, load_config
from excel2docx.transformer import transform
from excel2docx.generator import generate

from .scorer import (
    score_table_understanding,
    score_report_quality,
    BenchmarkResult,
    format_result,
)


def run_benchmark(
    excel_path: str | Path,
    config_path: str | Path,
    ground_truth_path: str | Path,
    output_path: str | Path | None = None,
    *,
    expected_schema_path: str | Path | None = None,
    label: str = "",
) -> BenchmarkResult:
    """Run pipeline and benchmark it.

    Args:
        excel_path: Path to Excel file
        config_path: Path to pipeline config YAML
        ground_truth_path: Path to ground truth JSON
        output_path: Optional path for output .docx (temp if None)
        expected_schema_path: Optional expected schema for report quality scoring
        label: Human-readable label for this benchmark run
    """
    started = time.time()

    # Load references
    config = load_config(config_path)
    ground_truth = json.loads(Path(ground_truth_path).read_text())

    expected_schema = None
    if expected_schema_path:
        expected_schema = json.loads(Path(expected_schema_path).read_text())

    # Output path
    if output_path is None:
        import tempfile
        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        output_path = tmp.name
        tmp.close()

    # 1. Parse
    parsed = parse_excel(excel_path, config.get("parser", {}))

    # 2. Transform
    transform_cfg = config.get("transform", {})
    if transform_cfg.get("mode") == "llm":
        # For LLM mode, the config typically includes prompt_template
        # We don't actually call LLM here unless an API key is provided
        report = transform(parsed, transform_cfg)
    else:
        report = transform(parsed, transform_cfg)

    # 3. Generate DOCX
    output_path = Path(output_path)
    try:
        generate(report, config.get("template", {}), output_path,
                metadata=config.get("metadata", {}))
    except Exception as e:
        # Generation failed — still score what we can
        pass

    # 4. Score
    table_score = score_table_understanding(parsed, ground_truth)
    report_score = score_report_quality(
        report, parsed,
        expected_schema=expected_schema,
        ground_truth=ground_truth,
    )

    elapsed = time.time() - started

    # Cleanup
    if output_path and output_path.exists() and "--output" not in sys.argv:
        output_path.unlink(missing_ok=True)

    return BenchmarkResult(
        table_score=table_score,
        report_score=report_score,
        config_name=label or Path(config_path).stem,
        excel_path=str(Path(excel_path).name),
        duration_seconds=elapsed,
    )


def run_suite(suite_name: str) -> list[BenchmarkResult]:
    """Run a named benchmark suite.

    Available suites:
        casino  — Casino PDR (surveillance log + winners/losers)
    """
    base = Path(__file__).resolve().parent

    if suite_name == "casino":
        excel = base / "test_data" / "casino_surveillance_log.xlsx"
        config = base / "configs" / "casino_pdr.yaml"
        ground_truth = base / "ground_truth" / "casino_surveillance.json"
        schema = base / "ground_truth" / "casino_pdr_schema.json"

        # Check files exist
        missing = []
        for p, label in [(excel, "Excel"), (config, "Config"), (ground_truth, "Ground truth")]:
            if not p.exists():
                missing.append(f"{label}: {p}")
        if missing:
            print(f"[ERROR] Missing files for 'casino' suite:")
            for m in missing:
                print(f"  {m}")
            sys.exit(1)

        result = run_benchmark(
            excel, config, ground_truth,
            expected_schema_path=schema if schema.exists() else None,
            label="Casino PDR",
        )
        return [result]

    print(f"Unknown suite: {suite_name}")
    print("Available: casino")
    sys.exit(1)


def main():
    ap = argparse.ArgumentParser(description="excel2docx Benchmark Runner")

    ap.add_argument("--suite", default=None,
                   help="Run a named benchmark suite (e.g., 'casino')")
    ap.add_argument("--excel", default=None, help="Path to Excel file")
    ap.add_argument("--config", default=None, help="Path to pipeline config YAML")
    ap.add_argument("--ground-truth", default=None, help="Path to ground truth JSON")
    ap.add_argument("--expected-schema", default=None,
                   help="Path to expected schema JSON (for report quality scoring)")
    ap.add_argument("--output", default=None, help="Output .docx path (temp if omitted)")
    ap.add_argument("--label", default="benchmark", help="Label for this run")
    ap.add_argument("--json", action="store_true",
                   help="Output raw JSON instead of formatted report")

    args = ap.parse_args()

    if args.suite:
        results = run_suite(args.suite)
    elif args.excel and args.config and args.ground_truth:
        results = [run_benchmark(
            args.excel, args.config, args.ground_truth,
            output_path=args.output,
            expected_schema_path=args.expected_schema,
            label=args.label,
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

    # Overall
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
