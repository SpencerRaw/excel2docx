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


def main():
    """CLI entry point."""
    ap = argparse.ArgumentParser(
        description="excel2docx — General-purpose Excel-to-DOCX report pipeline"
    )
    ap.add_argument("--excel", required=True, help="Path to Excel workbook")
    ap.add_argument("--config", required=True, help="Path to pipeline config (YAML/JSON)")
    ap.add_argument("--output", required=True, help="Output .docx path")
    ap.add_argument("--debug", action="store_true", help="Save intermediate JSON files")
    ap.add_argument("--api-key", default=None, help="LLM API key (if using LLM mode)")

    args = ap.parse_args()

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
