"""Official document generation skill."""

from __future__ import annotations

import argparse
import json
import os
import sys

import yaml

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))
from app.providers.llm import get_llm  # noqa: E402
from skills.generate_docx.prompts import build_prompt, parse_generated_json  # noqa: E402
from skills.generate_docx.render import render_docx  # noqa: E402


def load_terms(path):
    if not path or not os.path.exists(path):
        return []
    data = yaml.safe_load(open(path, encoding="utf-8")) or {}
    return list(data.get("terms", []))


def validate_generated(data: dict) -> None:
    if not isinstance(data, dict):
        raise ValueError("LLM output must be a JSON object")
    if not data.get("title"):
        raise ValueError("Generated document is missing title")
    if not isinstance(data.get("paragraphs", []), list):
        raise ValueError("paragraphs must be a list")
    if not isinstance(data.get("sections", []), list):
        raise ValueError("sections must be a list")


def generate_docx(
    brief: str,
    document_type: str = "notice",
    terms_path: str | None = None,
    terms: list[str] | None = None,
    style_card: dict | None = None,
    template_path: str | None = None,
    out_path: str | None = None,
) -> dict:
    loaded_terms = load_terms(terms_path)
    all_terms = list(dict.fromkeys(loaded_terms + list(terms or [])))
    system, user = build_prompt(brief, document_type, all_terms, style_card=style_card)
    raw = get_llm().chat(system, user)
    data = parse_generated_json(raw)
    validate_generated(data)

    out_path = out_path or f"generated_{document_type}.docx"
    rendered = render_docx(data, out_path, template_path=template_path)
    return {
        "out": rendered["out"],
        "title": data.get("title"),
        "document_type": document_type,
        "render_engine": rendered["engine"],
        "template_path": rendered.get("template_path"),
        "template_error": rendered.get("template_error"),
    }


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("brief")
    ap.add_argument("--document-type", default="notice")
    ap.add_argument("--terms", default="assets/terms/terms.yaml")
    ap.add_argument("--template", default=None)
    ap.add_argument("--out", default=None)
    args = ap.parse_args()
    result = generate_docx(
        args.brief,
        document_type=args.document_type,
        terms_path=args.terms,
        template_path=args.template,
        out_path=args.out,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
