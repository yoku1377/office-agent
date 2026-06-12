"""Small file-backed task store and runner for v1."""
from __future__ import annotations

import json
import re
import traceback
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from app.context import get_template_path, load_context, load_style_card
from app.paths import OUTPUT_DIR, TASK_DIR, UPLOAD_DIR, ensure_storage_dirs
from skills.generate_docx.generate import generate_docx
from skills.polish.polish import polish


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^\w.\-\u4e00-\u9fff]+", "_", name, flags=re.U).strip("._")
    return cleaned or "upload.docx"


class TaskStore:
    def __init__(self) -> None:
        ensure_storage_dirs()

    def task_path(self, task_id: str) -> Path:
        return TASK_DIR / f"{task_id}.json"

    def load(self, task_id: str) -> dict[str, Any] | None:
        path = self.task_path(task_id)
        if not path.exists():
            return None
        return json.loads(path.read_text(encoding="utf-8"))

    def save(self, task: dict[str, Any]) -> dict[str, Any]:
        task["updated_at"] = utc_now()
        self.task_path(task["id"]).write_text(
            json.dumps(task, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return task

    def create(
        self,
        *,
        skill: str,
        original_filename: str,
        input_path: str,
        params: dict[str, Any],
        department: str,
        user_id: str,
    ) -> dict[str, Any]:
        task_id = uuid.uuid4().hex
        task = {
            "id": task_id,
            "skill": skill,
            "status": "queued",
            "department": department,
            "user_id": user_id,
            "original_filename": original_filename,
            "input_path": input_path,
            "output_path": None,
            "params": params,
            "result": None,
            "error": None,
            "created_at": utc_now(),
            "updated_at": utc_now(),
        }
        return self.save(task)

    def update(self, task_id: str, **changes: Any) -> dict[str, Any]:
        task = self.load(task_id)
        if task is None:
            raise FileNotFoundError(task_id)
        task.update(changes)
        return self.save(task)


store = TaskStore()


def make_upload_path(task_id: str, filename: str) -> Path:
    return UPLOAD_DIR / f"{task_id}_{safe_filename(filename)}"


def make_output_path(task_id: str, filename: str, suffix: str) -> Path:
    base = safe_filename(filename)
    if base.lower().endswith(".docx"):
        base = base[:-5]
    return OUTPUT_DIR / f"{task_id}_{base}{suffix}.docx"


def run_task(task_id: str) -> None:
    task = store.load(task_id)
    if task is None:
        return

    try:
        store.update(task_id, status="running", error=None)
        if task["skill"] not in {"polish", "generate_docx"}:
            raise ValueError(f"Unsupported skill: {task['skill']}")

        context = load_context(task.get("department"))
        params = task.get("params", {})
        document_type = params.get("document_type")
        style_card = load_style_card(context, document_type)
        template_path = get_template_path(context, document_type)

        if task["skill"] == "generate_docx":
            out_path = make_output_path(
                task_id,
                task["original_filename"] or f"{document_type or 'document'}.docx",
                f"_{document_type or 'document'}",
            )
            result = generate_docx(
                params.get("brief", ""),
                document_type=document_type or "notice",
                terms_path=context.terms_path,
                terms=context.extra_terms,
                style_card=style_card,
                template_path=template_path,
                out_path=str(out_path),
            )
            store.update(
                task_id,
                status="succeeded",
                output_path=result["out"],
                result={
                    "title": result["title"],
                    "context": context.name,
                    "document_type": result["document_type"],
                    "style_card": style_card.get("name") if style_card else None,
                    "render_engine": result["render_engine"],
                    "template_path": result.get("template_path"),
                    "template_error": result.get("template_error"),
                },
            )
            return

        level = params.get("level") or context.default_polish_level
        out_path = make_output_path(task_id, task["original_filename"], f"_润色_{level}")
        result = polish(
            task["input_path"],
            level=level,
            terms_path=context.terms_path,
            terms=context.extra_terms,
            style_card=style_card,
            out_path=str(out_path),
            author=context.polish_author,
        )
        store.update(
            task_id,
            status="succeeded",
            output_path=result["out"],
            result={
                "applied": result["applied"],
                "rejected": result["rejected"],
                "context": context.name,
                "document_type": document_type,
                "style_card": style_card.get("name") if style_card else None,
            },
        )
    except Exception as exc:
        store.update(
            task_id,
            status="failed",
            error={"message": str(exc), "traceback": traceback.format_exc()},
        )
