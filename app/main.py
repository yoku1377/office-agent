"""office-agent service.

v1 exposes skills as reusable company APIs: upload a file, create a task,
poll status, and download the generated document. The old synchronous polish
endpoint is kept for compatibility with v0 demos.

启动： uvicorn app.main:app --host 0.0.0.0 --port 8080
"""
from __future__ import annotations

import os
import shutil
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.context import list_contexts, load_context, load_style_card
from app.skills import has_skill, list_skills
from app.tasks import make_upload_path, run_task, store
from skills.polish.polish import polish

app = FastAPI(title="office-agent", version="0.2")


def _assert_docx(file: UploadFile) -> None:
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="仅支持 .docx")


def _assert_document_type(document_type: str) -> None:
    allowed = {"notice", "opinion", "action_plan", "approval"}
    if document_type not in allowed:
        raise HTTPException(status_code=400, detail=f"document_type 仅支持: {', '.join(sorted(allowed))}")


def _task_response(task: dict) -> dict:
    payload = {
        "task_id": task["id"],
        "skill": task["skill"],
        "status": task["status"],
        "status_url": f"/tasks/{task['id']}",
        "download_url": None,
        "result": task.get("result"),
        "error": task.get("error"),
    }
    if task.get("status") == "succeeded" and task.get("output_path"):
        payload["download_url"] = f"/tasks/{task['id']}/download"
    return payload


@app.get("/healthz")
def healthz():
    return {"ok": True, "version": app.version}


@app.get("/skills")
def skills_endpoint():
    return {"skills": list_skills()}


@app.get("/contexts")
def contexts_endpoint():
    return {"contexts": list_contexts()}


@app.post("/tasks")
async def create_task(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    skill: str = Form("polish"),
    level: str = Form("medium"),
    department: str = Form("default"),
    document_type: str = Form("notice"),
    user_id: str = Form("anonymous"),
):
    _assert_docx(file)
    _assert_document_type(document_type)
    if skill != "polish" or not has_skill(skill):
        raise HTTPException(status_code=400, detail=f"暂不支持 skill: {skill}")
    if level not in {"light", "medium", "heavy"}:
        raise HTTPException(status_code=400, detail="level 仅支持 light / medium / heavy")

    task = store.create(
        skill=skill,
        original_filename=file.filename,
        input_path="",
        params={"level": level, "document_type": document_type},
        department=department,
        user_id=user_id,
    )
    upload_path = make_upload_path(task["id"], file.filename)
    with upload_path.open("wb") as f:
        shutil.copyfileobj(file.file, f)
    task = store.update(task["id"], input_path=str(upload_path))

    background_tasks.add_task(run_task, task["id"])
    return _task_response(task)


@app.post("/tasks/polish")
async def create_polish_task(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    level: str = Form("medium"),
    department: str = Form("default"),
    document_type: str = Form("notice"),
    user_id: str = Form("anonymous"),
):
    _assert_document_type(document_type)
    return await create_task(
        background_tasks=background_tasks,
        file=file,
        skill="polish",
        level=level,
        department=department,
        document_type=document_type,
        user_id=user_id,
    )


@app.post("/tasks/generate-docx")
async def create_generate_docx_task(
    background_tasks: BackgroundTasks,
    brief: str = Form(...),
    department: str = Form("admin"),
    document_type: str = Form("notice"),
    user_id: str = Form("anonymous"),
):
    _assert_document_type(document_type)
    if not has_skill("generate_docx"):
        raise HTTPException(status_code=400, detail="generate_docx skill 未安装")
    if not brief.strip():
        raise HTTPException(status_code=400, detail="brief 不能为空")

    task = store.create(
        skill="generate_docx",
        original_filename=f"{document_type}.docx",
        input_path="",
        params={"brief": brief, "document_type": document_type},
        department=department,
        user_id=user_id,
    )
    background_tasks.add_task(run_task, task["id"])
    return _task_response(task)

@app.post("/tasks/generate-pptx")
async def create_generate_pptx_task(
    background_tasks: BackgroundTasks,
    brief: str = Form(...),
    department: str = Form("admin"),
    template_path: str = Form(None),
    user_id: str = Form("anonymous"),
):
    if not has_skill("generate_pptx"):
        raise HTTPException(status_code=400, detail="generate_pptx skill 未安装")
    if not brief.strip():
        raise HTTPException(status_code=400, detail="brief 不能为空")

    task = store.create(
        skill="generate_pptx",
        original_filename="presentation.pptx",
        input_path="",
        params={"brief": brief, "template_path": template_path},
        department=department,
        user_id=user_id,
    )
    background_tasks.add_task(run_task, task["id"])
    return _task_response(task)


@app.get("/tasks/{task_id}")
def get_task(task_id: str):
    task = store.load(task_id)
    if task is None:
        raise HTTPException(status_code=404, detail="任务不存在")
    return _task_response(task)


@app.get("/tasks/{task_id}/download")
def download_task(task_id: str):
    task = store.load(task_id)
    if task is None:
        raise HTTPException(status_code=404, detail="任务不存在")
    if task.get("status") != "succeeded" or not task.get("output_path"):
        raise HTTPException(status_code=409, detail="任务尚未完成")

    output_path = Path(task["output_path"])
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="输出文件不存在")
    return FileResponse(output_path, filename=os.path.basename(output_path))


@app.post("/skills/polish")
async def polish_endpoint(
    file: UploadFile = File(...),
    level: str = Form("medium"),
    department: str = Form("default"),
    document_type: str = Form("notice"),
):
    """v0-compatible synchronous endpoint."""
    _assert_docx(file)
    _assert_document_type(document_type)
    if level not in {"light", "medium", "heavy"}:
        raise HTTPException(status_code=400, detail="level 仅支持 light / medium / heavy")

    task = store.create(
        skill="polish",
        original_filename=file.filename,
        input_path="",
        params={"level": level, "document_type": document_type},
        department=department,
        user_id="sync",
    )
    upload_path = make_upload_path(task["id"], file.filename)
    with upload_path.open("wb") as f:
        shutil.copyfileobj(file.file, f)
    context = load_context(department)
    style_card = load_style_card(context, document_type)
    out_path = str(upload_path).replace(".docx", f"_润色_{level}.docx")
    result = polish(
        str(upload_path),
        level=level,
        terms_path=context.terms_path,
        terms=context.extra_terms,
        style_card=style_card,
        out_path=out_path,
        author=context.polish_author,
    )
    store.update(
        task["id"],
        status="succeeded",
        input_path=str(upload_path),
        output_path=result["out"],
        result={
            "applied": result["applied"],
            "rejected": result["rejected"],
            "context": context.name,
            "document_type": document_type,
            "style_card": style_card.get("name") if style_card else None,
        },
    )
    return FileResponse(result["out"], filename=os.path.basename(result["out"]))
