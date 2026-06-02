"""excel2docx — General-purpose Excel-to-DOCX report generation pipeline.

Architecture:
  1. Parser  — read Excel sheets via schema config
  2. Transformer — rule-based mapping or LLM-powered text generation
  3. Generator — produce formatted .docx from structured data + template

Domain-agnostic. Config-driven. LLM is optional — rule-based transforms
work for purely mechanical field mapping.
"""

__version__ = "0.2.0"
__all__ = ["parser", "transformer", "generator", "pipeline"]
