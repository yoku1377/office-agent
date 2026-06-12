"""编排服务骨架（v0）：先以单个 HTTP 接口直连 skill，
Agent SDK 编排与任务队列在 v1 接入（见 README 路线）。

启动： uvicorn app.main:app --host 0.0.0.0 --port 8080
"""
import os
import shutil
import tempfile

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, JSONResponse

from skills.polish.polish import polish

app = FastAPI(title="office-agent", version="0.1")

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "storage", "outputs")


@app.post("/skills/polish")
async def polish_endpoint(file: UploadFile = File(...), level: str = Form("medium")):
    if not file.filename.endswith(".docx"):
        return JSONResponse({"error": "仅支持 .docx"}, status_code=400)
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        shutil.copyfileobj(file.file, tmp)
        in_path = tmp.name
    out_path = os.path.join(OUTPUT_DIR, file.filename.replace(".docx", f"_润色_{level}.docx"))
    result = polish(in_path, level=level, terms_path="assets/terms/terms.yaml", out_path=out_path)
    return FileResponse(result["out"], filename=os.path.basename(result["out"]))


@app.get("/healthz")
def healthz():
    return {"ok": True}
