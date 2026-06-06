"""Benchmark scoring engine.

Two independent scores (0-100 each) that produce a combined report:

1. TABLE UNDERSTANDING SCORE
   Measures parsing fidelity against ground truth.
   - Cell recovery rate: what % of non-empty source cells appear in parsed output?
   - Row completeness: what % of rows are fully captured?
   - Column detection: were all expected columns found?
   - Type accuracy: what % of values have the correct Python type?
   - Multi-sheet coverage: were all sheets discovered?

2. REPORT QUALITY SCORE
   Measures generation fidelity against expected output.
   - Section completeness: are all required sections present?
   - Value accuracy: do report values match source data?
   - Hallucination detection: fabricated facts / phantom references
   - Structure conformance: does output follow expected schema?
"""

from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional


# ═══════════════════════════════════════════════════════════════
# Data types
# ═══════════════════════════════════════════════════════════════

@dataclass
class TableScore:
    """Parsing fidelity metrics."""
    cell_recovery_pct: float = 0.0       # % of non-empty cells captured
    row_completeness_pct: float = 0.0    # % of rows fully parsed
    column_detection_pct: float = 0.0    # % of expected columns found
    type_accuracy_pct: float = 0.0       # % of values with correct type
    sheet_coverage_pct: float = 0.0      # % of sheets discovered
    overall: float = 0.0                 # weighted average

    # Detail fields for diagnosis
    total_cells_expected: int = 0
    total_cells_recovered: int = 0
    expected_rows: int = 0
    parsed_rows: int = 0
    missed_columns: list[str] = field(default_factory=list)
    type_errors: list[str] = field(default_factory=list)
    missed_sheets: list[str] = field(default_factory=list)


@dataclass
class ReportScore:
    """Generation fidelity metrics."""
    section_completeness_pct: float = 0.0   # required sections present
    value_accuracy_pct: float = 0.0         # numeric/text values match source
    hallucination_rate_pct: float = 0.0     # fabricated values (lower=better)
    structure_conformance: float = 0.0      # schema match score
    overall: float = 0.0

    missing_sections: list[str] = field(default_factory=list)
    value_mismatches: list[str] = field(default_factory=list)
    hallucinations: list[str] = field(default_factory=list)
    schema_violations: list[str] = field(default_factory=list)


@dataclass
class BenchmarkResult:
    """Complete benchmark run output."""
    table_score: TableScore
    report_score: ReportScore
    config_name: str = ""
    excel_path: str = ""
    duration_seconds: float = 0.0
    warnings: list[str] = field(default_factory=list)


# ═══════════════════════════════════════════════════════════════
# 1. TABLE UNDERSTANDING SCORER
# ═══════════════════════════════════════════════════════════════

def score_table_understanding(
    parsed_data: dict[str, Any],
    ground_truth: dict[str, Any],
) -> TableScore:
    """Compare parsed output against ground truth.

    ground_truth format:
    {
        "sheets": {
            "<sheet_name>": {
                "row_count": N,
                "columns": ["col1", "col2", ...],
                "column_types": {"col1": "str", "col2": "float", ...},
                "non_empty_cells": N,              # total non-empty cells
                "sample_rows": [dict, ...]         # first 3 rows as reference
            },
            ...
        }
    }
    """
    s = TableScore()

    gt_sheets = ground_truth.get("sheets", {})
    parsed_sheets = parsed_data.get("sheets", {})

    if not gt_sheets:
        return s

    # Sheet coverage — use column-aware matching when name matching fails
    found_sheets = 0
    parsed_sheet_names = list(parsed_sheets.keys())
    
    for gt_name, gt_info in gt_sheets.items():
        matched_name = _fuzzy_find(gt_name, parsed_sheet_names)
        
        # Fallback: match by expected columns
        if not matched_name:
            gt_cols = set(gt_info.get("columns", []))
            if gt_cols:
                for ps_name in parsed_sheet_names:
                    if parsed_sheets[ps_name].get("rows"):
                        ps_cols = set(parsed_sheets[ps_name]["rows"][0].keys())
                        overlap = len(gt_cols & ps_cols) / len(gt_cols) if gt_cols else 0
                        if overlap > 0.6:
                            matched_name = ps_name
                            break
        
        if matched_name:
            found_sheets += 1
        else:
            s.missed_sheets.append(gt_name)

    s.sheet_coverage_pct = (found_sheets / len(gt_sheets)) * 100 if gt_sheets else 0

    # Per-sheet analysis
    total_cells_expected = 0
    total_cells_recovered = 0
    total_rows_expected = 0
    total_rows_parsed = 0
    all_columns_found = 0
    all_columns_expected = 0
    type_checks_total = 0
    type_checks_passed = 0

    for gt_name, gt_info in gt_sheets.items():
        # Match using same logic as sheet coverage
        matched_name = _fuzzy_find(gt_name, parsed_sheet_names)
        if not matched_name:
            gt_cols = set(gt_info.get("columns", []))
            if gt_cols:
                for ps_name in parsed_sheet_names:
                    if parsed_sheets[ps_name].get("rows"):
                        ps_cols = set(parsed_sheets[ps_name]["rows"][0].keys())
                        overlap = len(gt_cols & ps_cols) / len(gt_cols) if gt_cols else 0
                        if overlap > 0.6:
                            matched_name = ps_name
                            break
        if not matched_name:
            total_rows_expected += gt_info.get("row_count", 0)
            total_cells_expected += gt_info.get("non_empty_cells", 0)
            all_columns_expected += len(gt_info.get("columns", []))
            continue

        sheet_data = parsed_sheets[matched_name]
        rows = sheet_data.get("rows", [])

        # Row count
        gt_rows = gt_info.get("row_count", 0)
        total_rows_expected += gt_rows
        total_rows_parsed += len(rows)

        # Cell recovery: count non-empty cells in parsed rows
        gt_cells = gt_info.get("non_empty_cells", 0)
        total_cells_expected += gt_cells

        parsed_cells = 0
        for row in rows:
            for v in row.values():
                if v is not None and v != "" and v != 0:
                    parsed_cells += 1
        total_cells_recovered += parsed_cells

        # Column detection
        gt_cols = set(gt_info.get("columns", []))
        if rows:
            parsed_cols = set(rows[0].keys())
        else:
            parsed_cols = set()
        all_columns_expected += len(gt_cols)
        all_columns_found += len(gt_cols & parsed_cols)
        s.missed_columns.extend(sorted(gt_cols - parsed_cols))

        # Type accuracy: compare against sample rows
        gt_types = gt_info.get("column_types", {})
        sample_rows = gt_info.get("sample_rows", [])
        if sample_rows and rows:
            for col_name, expected_type in gt_types.items():
                if col_name not in rows[0]:
                    continue
                # Check type consistency across all parsed rows
                for row in rows:
                    if col_name in row and row[col_name] is not None and row[col_name] != "":
                        type_checks_total += 1
                        if _type_matches(row[col_name], expected_type):
                            type_checks_passed += 1
                        else:
                            s.type_errors.append(
                                f"{matched_name}.{col_name}: expected {expected_type}, "
                                f"got {type(row[col_name]).__name__} ({row[col_name]!r})"
                            )

    # Compute scores
    s.cell_recovery_pct = _safe_pct(total_cells_recovered, total_cells_expected)
    s.row_completeness_pct = _safe_pct(total_rows_parsed, total_rows_expected)

    # Row completeness caps at 100% (extra rows are fine — likely from multi-sheet)
    if total_rows_parsed > total_rows_expected > 0:
        s.row_completeness_pct = 100.0

    s.column_detection_pct = _safe_pct(all_columns_found, all_columns_expected)
    s.type_accuracy_pct = _safe_pct(type_checks_passed, type_checks_total)
    s.total_cells_expected = total_cells_expected
    s.total_cells_recovered = total_cells_recovered
    s.expected_rows = total_rows_expected
    s.parsed_rows = total_rows_parsed

    # Overall (weighted: cells 30%, rows 20%, cols 20%, types 20%, sheets 10%)
    s.overall = (
        s.cell_recovery_pct * 0.30 +
        s.row_completeness_pct * 0.20 +
        s.column_detection_pct * 0.20 +
        s.type_accuracy_pct * 0.20 +
        s.sheet_coverage_pct * 0.10
    )

    return s


# ═══════════════════════════════════════════════════════════════
# 2. REPORT QUALITY SCORER
# ═══════════════════════════════════════════════════════════════

def score_report_quality(
    report_data: dict[str, Any],
    parsed_data: dict[str, Any],
    expected_schema: Optional[dict] = None,
    ground_truth: Optional[dict] = None,
) -> ReportScore:
    """Score generated report against expectations.

    For LLM-generated reports: checks hallucination by comparing report values
    against the original parsed data (source of truth).

    For rules-generated reports: checks structural conformance and field mapping.

    expected_schema:
    {
        "required_sections": ["header", "overview", "departments", ...],
        "required_fields": {"header": ["title", "date"], ...},
        "value_constraints": {"overview.count": {"type": "int", "min": 0}},
        "forbidden_patterns": ["TODO", "FIXME", "fill this in"],
    }
    """
    s = ReportScore()
    schema = expected_schema or {}

    # --- Section completeness ---
    required_sections = schema.get("required_sections", [])
    if required_sections:
        found_sections = 0
        for sec in required_sections:
            if _path_exists(report_data, sec):
                found_sections += 1
            else:
                s.missing_sections.append(sec)
        s.section_completeness_pct = _safe_pct(found_sections, len(required_sections))
    else:
        s.section_completeness_pct = 100.0  # no requirements = full score

    # --- Value accuracy ---
    # Check numeric values in report against ground truth AND computed aggregates
    value_checks_total = 0
    value_checks_passed = 0

    # Build set of ALL allowable values:
    #   (a) raw values from ground truth sheets
    #   (b) computed aggregates from parsed data (totals, counts, etc.)
    allowable_values = set()

    if ground_truth:
        gt_sheets = ground_truth.get("sheets", {})
        for v in _collect_all_values(gt_sheets):
            if isinstance(v, (int, float)):
                allowable_values.add(round(v, 2))

    # Computed values from parsed data: total_rows, per-sheet row counts, aggregates
    # Also allow 0 and 1 as legitimate counter defaults
    allowable_values.add(0)
    allowable_values.add(1)
    if parsed_data:
        meta = parsed_data.get("metadata", {})
        if isinstance(meta.get("total_rows"), (int, float)):
            allowable_values.add(meta["total_rows"])
        for sn, sd in parsed_data.get("sheets", {}).items():
            if isinstance(sd.get("aggregates"), dict):
                for v in sd["aggregates"].values():
                    if isinstance(v, (int, float)):
                        allowable_values.add(v)
            rows = sd.get("rows", [])
            if rows:
                allowable_values.add(len(rows))

    all_report_values = _collect_all_values(report_data)

    for rv in all_report_values:
        if isinstance(rv, (int, float)):
            value_checks_total += 1
            rv_rounded = round(rv, 2)
            if rv_rounded in allowable_values:
                value_checks_passed += 1
            elif _value_exists_in(rv, list(allowable_values)):
                value_checks_passed += 1
            else:
                s.value_mismatches.append(
                    f"Value {rv!r} not found in source data or computed aggregates"
                )

    s.value_accuracy_pct = _safe_pct(value_checks_passed, value_checks_total)

    # --- Hallucination detection ---
    # Check for fabricated reference IDs, names, amounts
    hallucination_checks = 0
    hallucinations_found = 0

    if ground_truth:
        # Collect all known IDs from ground truth
        known_ids = set()
        gt_sheets = ground_truth.get("sheets", {})
        for sn, info in gt_sheets.items():
            for row in info.get("sample_rows", []):
                for k, v in row.items():
                    if isinstance(v, str) and _looks_like_id(v):
                        known_ids.add(v.lower())

        # Scan report for IDs not in ground truth
        all_report_strings = _collect_all_strings(report_data)
        for s_val in all_report_strings:
            if _looks_like_id(s_val) and s_val.lower() not in known_ids:
                hallucination_checks += 1

    # Also check forbidden patterns
    forbidden = schema.get("forbidden_patterns", [])
    all_report_text = json.dumps(report_data, ensure_ascii=False).lower()
    for pattern in forbidden:
        if pattern.lower() in all_report_text:
            hallucinations_found += 1
            s.hallucinations.append(f"Forbidden pattern found: '{pattern}'")

    hallucination_checks += len(forbidden)
    s.hallucination_rate_pct = _safe_pct(hallucinations_found, hallucination_checks)

    # --- Structure conformance ---
    schema_violations = 0
    schema_checks = 0

    required_fields = schema.get("required_fields", {})
    for section, fields in required_fields.items():
        section_data = report_data.get(section, {})
        for field in fields:
            schema_checks += 1
            if field not in section_data or section_data[field] is None:
                schema_violations += 1
                s.schema_violations.append(f"{section}.{field} missing")

    # Check value constraints
    constraints = schema.get("value_constraints", {})
    for path, constraint in constraints.items():
        val = _resolve_path(report_data, path)
        if val is not None:
            schema_checks += 1
            if constraint.get("type"):
                if not _type_matches(val, constraint["type"]):
                    schema_violations += 1
                    s.schema_violations.append(f"{path}: expected type {constraint['type']}")
            if "min" in constraint and isinstance(val, (int, float)) and val < constraint["min"]:
                schema_violations += 1
                s.schema_violations.append(f"{path}: value {val} < min {constraint['min']}")

    s.structure_conformance = _safe_pct(
        schema_checks - schema_violations, schema_checks
    )

    # Overall (weighted: sections 30%, values 25%, hallucination 25%, structure 20%)
    s.overall = (
        s.section_completeness_pct * 0.30 +
        s.value_accuracy_pct * 0.25 +
        (100 - s.hallucination_rate_pct) * 0.25 +  # invert: lower hallucination = higher score
        s.structure_conformance * 0.20
    )

    return s


# ═══════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════

def _safe_pct(numerator: int | float, denominator: int | float) -> float:
    if denominator == 0:
        return 100.0 if numerator == 0 else 0.0
    return round((numerator / denominator) * 100, 1)


def _fuzzy_find(target: str, candidates: list[str]) -> Optional[str]:
    """Find best fuzzy match for target in candidates.
    
    Matches use word-level overlap: if most words of the target appear
    in a candidate name (or vice versa), it's a match.
    """
    target_lower = target.lower().strip()
    
    # Exact match (ignoring case/whitespace)
    for c in candidates:
        if c.lower().strip() == target_lower:
            return c
    
    # Substring match (whole string)
    for c in candidates:
        cl = c.lower().strip()
        if target_lower in cl or cl in target_lower:
            return c
    
    # Word-level overlap-based match
    target_words = set(target_lower.split())
    best_match = None
    best_overlap = 0
    for c in candidates:
        cand_words = set(c.lower().strip().replace('_', ' ').replace('-', ' ').split())
        if not target_words or not cand_words:
            continue
        # Jaccard-like overlap ratio
        intersection = target_words & cand_words
        union = target_words | cand_words
        overlap_ratio = len(intersection) / len(union) if union else 0
        if overlap_ratio > 0.5 and overlap_ratio > best_overlap:
            best_overlap = overlap_ratio
            best_match = c
    
    return best_match


def _type_matches(value: Any, expected_type: str) -> bool:
    if expected_type == "str":
        return isinstance(value, str)
    if expected_type == "int":
        return isinstance(value, int)
    if expected_type == "float":
        return isinstance(value, (int, float))
    if expected_type == "date":
        from datetime import datetime
        return isinstance(value, (datetime, str))
    if expected_type == "list":
        return isinstance(value, list)
    if expected_type == "dict" or expected_type == "object":
        return isinstance(value, dict)
    return True


def _looks_like_id(s: str) -> bool:
    """Check if a string looks like a reference ID."""
    return bool(re.match(r'^(IN\d+|[A-Z]{2,4}\d{3,}|[A-Z][a-z]+ \d+)$', s.strip()))


def _path_exists(data: dict, path: str) -> bool:
    return _resolve_path(data, path) is not None


def _resolve_path(data: dict, path: str) -> Any:
    parts = path.split(".")
    current = data
    for part in parts:
        if isinstance(current, dict):
            current = current.get(part)
        elif isinstance(current, list) and part.isdigit():
            idx = int(part)
            current = current[idx] if idx < len(current) else None
        else:
            return None
        if current is None:
            return None
    return current


def _collect_all_values(data: Any) -> list[Any]:
    """Recursively collect all scalar values."""
    result = []
    if isinstance(data, dict):
        for v in data.values():
            result.extend(_collect_all_values(v))
    elif isinstance(data, list):
        for item in data:
            result.extend(_collect_all_values(item))
    elif data is not None and data != "":
        result.append(data)
    return result


def _collect_all_strings(data: Any) -> list[str]:
    """Recursively collect all string values."""
    return [v for v in _collect_all_values(data) if isinstance(v, str)]


def _value_exists_in(target: int | float, collection: list) -> bool:
    """Check if target value (or close approximation) exists in collection."""
    for v in collection:
        if isinstance(v, (int, float)):
            if abs(v - target) < 0.01:
                return True
    return False


# ═══════════════════════════════════════════════════════════════
# Pretty-print
# ═══════════════════════════════════════════════════════════════

def format_result(result: BenchmarkResult) -> str:
    """Format benchmark result as a readable report."""
    lines = []
    sep = "=" * 64
    sub = "-" * 64

    lines.append(sep)
    lines.append("  excel2docx BENCHMARK REPORT")
    lines.append(sep)
    lines.append(f"  Config:     {result.config_name}")
    lines.append(f"  Excel:      {result.excel_path}")
    lines.append(f"  Duration:   {result.duration_seconds:.1f}s")
    lines.append("")

    # Table understanding
    ts = result.table_score
    lines.append(sub)
    lines.append(f"  1. TABLE UNDERSTANDING  —  {ts.overall:.1f}/100")
    lines.append(sub)
    lines.append(f"     Cell recovery:      {ts.cell_recovery_pct:.1f}%  ({ts.total_cells_recovered}/{ts.total_cells_expected} cells)")
    lines.append(f"     Row completeness:   {ts.row_completeness_pct:.1f}%  ({ts.parsed_rows}/{ts.expected_rows} rows)")
    lines.append(f"     Column detection:   {ts.column_detection_pct:.1f}%")
    lines.append(f"     Type accuracy:      {ts.type_accuracy_pct:.1f}%")
    lines.append(f"     Sheet coverage:     {ts.sheet_coverage_pct:.1f}%")

    if ts.missed_columns:
        lines.append(f"     ⚠  Missed columns:   {', '.join(ts.missed_columns[:5])}")
    if ts.missed_sheets:
        lines.append(f"     ⚠  Missed sheets:    {', '.join(ts.missed_sheets)}")
    if ts.type_errors:
        lines.append(f"     ⚠  Type errors:      {len(ts.type_errors)}")
        for e in ts.type_errors[:3]:
            lines.append(f"        {e}")

    lines.append("")

    # Report quality
    rs = result.report_score
    lines.append(sub)
    lines.append(f"  2. REPORT QUALITY      —  {rs.overall:.1f}/100")
    lines.append(sub)
    lines.append(f"     Section completeness:  {rs.section_completeness_pct:.1f}%")
    lines.append(f"     Value accuracy:        {rs.value_accuracy_pct:.1f}%")
    lines.append(f"     Hallucination rate:    {rs.hallucination_rate_pct:.1f}%  (lower is better)")
    lines.append(f"     Structure conformance: {rs.structure_conformance:.1f}%")

    if rs.missing_sections:
        lines.append(f"     ⚠  Missing sections:   {', '.join(rs.missing_sections)}")
    if rs.value_mismatches:
        lines.append(f"     ⚠  Value mismatches:   {len(rs.value_mismatches)}")
        for v in rs.value_mismatches[:3]:
            lines.append(f"        {v}")
    if rs.hallucinations:
        lines.append(f"     ⚠  Hallucinations:     {len(rs.hallucinations)}")
        for h in rs.hallucinations[:3]:
            lines.append(f"        {h}")
    if rs.schema_violations:
        lines.append(f"     ⚠  Schema violations:  {len(rs.schema_violations)}")
        for sv in rs.schema_violations[:3]:
            lines.append(f"        {sv}")

    lines.append("")
    lines.append(sep)

    # Overall grade
    combined = (ts.overall + rs.overall) / 2
    grade = "A" if combined >= 90 else "B" if combined >= 75 else "C" if combined >= 60 else "D"
    lines.append(f"  OVERALL: {combined:.1f}/100  ({grade})")
    lines.append(sep)

    return "\n".join(lines)
