"""Company/department context loading.

The service keeps business context outside skill code so every entrypoint
(web UI, OpenClaw, enterprise chat bots) can reuse the same vocabulary and
template preferences.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

from app.paths import ASSETS_DIR, ROOT_DIR


DEFAULT_CONTEXT = "default"


@dataclass(frozen=True)
class OfficeContext:
    name: str
    terms_path: str | None
    extra_terms: list[str]
    style_cards: dict[str, str]
    templates: dict[str, str]
    polish_author: str
    default_polish_level: str


def _resolve_path(value: str | None) -> str | None:
    if not value:
        return None
    path = Path(value)
    if not path.is_absolute():
        path = ROOT_DIR / path
    return str(path)


def load_context(name: str | None = None) -> OfficeContext:
    context_name = (name or DEFAULT_CONTEXT).strip() or DEFAULT_CONTEXT
    context_path = ASSETS_DIR / "contexts" / f"{context_name}.yaml"
    if not context_path.exists() and context_name != DEFAULT_CONTEXT:
        context_name = DEFAULT_CONTEXT
        context_path = ASSETS_DIR / "contexts" / f"{context_name}.yaml"

    data: dict[str, Any] = {}
    if context_path.exists():
        data = yaml.safe_load(context_path.read_text(encoding="utf-8")) or {}

    polish = data.get("polish") or {}
    examples = data.get("examples") or {}
    style_cards = {
        str(key): str(_resolve_path(value))
        for key, value in (examples.get("style_cards") or {}).items()
        if value
    }
    templates = {
        str(key): str(_resolve_path(value))
        for key, value in (data.get("templates") or {}).items()
        if value
    }
    return OfficeContext(
        name=str(data.get("name") or context_name),
        terms_path=_resolve_path(data.get("terms_path") or "assets/terms/terms.yaml"),
        extra_terms=list(data.get("extra_terms") or []),
        style_cards=style_cards,
        templates=templates,
        polish_author=str(polish.get("author") or "AI润色"),
        default_polish_level=str(polish.get("default_level") or "medium"),
    )


def load_style_card(context: OfficeContext, document_type: str | None) -> dict[str, Any] | None:
    if not document_type:
        return None
    card_path = context.style_cards.get(document_type)
    if not card_path:
        return None

    path = Path(card_path)
    if not path.exists():
        return None
    return yaml.safe_load(path.read_text(encoding="utf-8")) or None


def get_template_path(context: OfficeContext, document_type: str | None) -> str | None:
    if not document_type:
        return None
    template_path = context.templates.get(document_type)
    if not template_path:
        return None
    path = Path(template_path)
    return str(path) if path.exists() else None


def list_contexts() -> list[dict[str, Any]]:
    contexts: list[dict[str, Any]] = []
    for context_path in sorted((ASSETS_DIR / "contexts").glob("*.yaml")):
        context = load_context(context_path.stem)
        contexts.append(
            {
                "id": context_path.stem,
                "name": context.name,
                "document_types": sorted(context.style_cards.keys()),
                "templates": sorted(context.templates.keys()),
                "default_polish_level": context.default_polish_level,
            }
        )
    return contexts
