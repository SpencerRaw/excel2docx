"""Benchmark framework for excel2docx.

Measures two dimensions:
  1. Table Understanding  — how well the parser captures Excel data
  2. Report Quality        — how well the pipeline generates accurate DOCX

Usage:
    python -m benchmark.runner --excel data.xlsx --config config.yaml --ground-truth ground_truth.json
    python -m benchmark.runner --suite casino
"""
