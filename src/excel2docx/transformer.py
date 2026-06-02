"""Data transformation layer.

Two modes:
  1. Rule-based — mechanical field mapping, aggregation, formatting (no LLM needed)
  2. LLM-powered — send structured data to an LLM for narrative generation

Config defines which mode + the rules/prompt template.
"""

from __future__ import annotations

from typing import Any, Optional


def transform(
    parsed: dict[str, Any],
    config: dict,
    *,
    llm_client: Optional[callable] = None,
) -> dict[str, Any]:
    """Transform parsed Excel data into report-ready structured data.

    Config structure:
        mode: "rules" | "llm"
        rules:                 # for mode=rules
          - field: <str>
            source: <str>      # "sheets.<name>.rows[0].<field>" or "sheets.<name>.aggregates"
            format: <str>      # optional format template
        prompt_template: <str> # for mode=llm
        llm_config:
          model: <str>
          temperature: <float>
          max_tokens: <int>
        output_schema: <dict>  # expected JSON schema for LLM output
    """
    mode = config.get("mode", "rules")

    if mode == "rules":
        return _transform_rules(parsed, config.get("rules", []))
    elif mode == "llm":
        if llm_client is None:
            raise ValueError("LLM mode requires an llm_client callable")
        return _transform_llm(parsed, config, llm_client)
    else:
        raise ValueError(f"Unknown transform mode: {mode}")


def _transform_rules(parsed: dict, rules: list[dict]) -> dict:
    """Apply mechanical field mapping rules."""
    result = {}
    for rule in rules:
        field = rule["field"]
        source = rule.get("source", field)
        fmt = rule.get("format")
        default = rule.get("default", "")

        value = _resolve_source(parsed, source)
        if value is None:
            value = default
        if fmt and value:
            try:
                value = fmt.format(value=value)
            except (KeyError, ValueError, TypeError):
                pass

        # Support nested keys like "header.title"
        keys = field.split(".")
        target = result
        for k in keys[:-1]:
            target = target.setdefault(k, {})
        target[keys[-1]] = value

    return result


def _transform_llm(
    parsed: dict,
    config: dict,
    llm_client: callable,
) -> dict:
    """Send parsed data to LLM for narrative transformation."""
    import json
    import re

    prompt_template = config.get("prompt_template", "")
    output_schema = config.get("output_schema", {})
    llm_config = config.get("llm_config", {})

    # Build prompt from data
    prompt = _build_prompt(prompt_template, parsed)

    # Call LLM
    response = llm_client(
        prompt=prompt,
        response_schema=output_schema,
        **llm_config,
    )

    # Parse JSON output
    if isinstance(response, dict):
        return response

    content = response if isinstance(response, str) else str(response)
    # Strip markdown fences
    content = re.sub(r"^```(?:json)?\s*", "", content.strip())
    content = re.sub(r"\s*```$", "", content)
    return json.loads(content)


def _resolve_source(data: dict, source: str) -> Any:
    """Resolve a dot-path source reference like 'sheets.main.rows.0.name'."""
    parts = source.split(".")
    current = data
    for part in parts:
        if isinstance(current, dict):
            current = current.get(part)
        elif isinstance(current, list) and part.isdigit():
            current = current[int(part)] if int(part) < len(current) else None
        else:
            return None
        if current is None:
            return None
    return current


def _build_prompt(template: str, data: dict) -> str:
    """Build prompt from template + parsed data using simple substitution."""
    import json

    result = template

    # Inject parsed data as JSON
    if "{{DATA}}" in result or "{{ data }}" in result:
        result = result.replace("{{DATA}}", json.dumps(data, indent=2, ensure_ascii=False))
        result = result.replace("{{ data }}", json.dumps(data, indent=2, ensure_ascii=False))

    # Inject partial data for specific sheets
    for sheet_name, sheet_data in data.get("sheets", {}).items():
        placeholder_rows = f"{{{{{sheet_name}.rows}}}}"
        placeholder_agg = f"{{{{{sheet_name}.aggregates}}}}"
        if placeholder_rows in result:
            result = result.replace(placeholder_rows, json.dumps(sheet_data.get("rows", []), indent=2, ensure_ascii=False))

    return result
