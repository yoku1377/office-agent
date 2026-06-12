"""Skill discovery helpers."""
from __future__ import annotations

import re
from pathlib import Path

from app.paths import ROOT_DIR


SKILLS_DIR = ROOT_DIR / "skills"


def _parse_frontmatter(text: str) -> dict[str, str]:
    match = re.match(r"^---\s*\n(.*?)\n---", text, flags=re.S)
    if not match:
        return {}

    meta: dict[str, str] = {}
    for line in match.group(1).splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        meta[key.strip()] = value.strip()
    return meta


def list_skills() -> list[dict[str, str]]:
    skills: list[dict[str, str]] = []
    for skill_md in sorted(SKILLS_DIR.glob("*/SKILL.md")):
        text = skill_md.read_text(encoding="utf-8")
        meta = _parse_frontmatter(text)
        name = meta.get("name") or skill_md.parent.name
        skills.append(
            {
                "name": name,
                "description": meta.get("description", ""),
                "path": str(skill_md.relative_to(ROOT_DIR)),
            }
        )
    return skills


def has_skill(name: str) -> bool:
    return any(skill["name"] == name for skill in list_skills())
