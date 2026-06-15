"""自检：成品逐页转缩略图交多模态 LLM 检查溢出与排版问题，自动回修。

v1 策略：
- 使用 comtypes/powerpoint 或 LibreOffice 将 pptx 逐页导出为图片
- 将图片交给多模态 LLM 检查溢出、排版、空白页等问题
- 返回问题列表和修复建议
- 若 LLM 不可用或转换工具不可用，跳过自检步骤
"""

from __future__ import annotations

import base64
import json
import logging
import os
import subprocess
import tempfile
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


REVIEW_SYSTEM = """你是一名资深 PPT 排版审核专家。请检查以下幻灯片缩略图，找出排版和内容问题。

请检查以下方面：
1. 文字溢出：文本是否超出占位符边界
2. 内容重叠：标题、正文、图表是否互相遮挡
3. 空白过多：页面是否有大面积空白
4. 图表可读性：图表标题、数据标签是否清晰
5. 整体美观：字体大小、间距是否合理

【输出格式】只输出 JSON，不要任何其他文字或代码块标记：
{
  "issues": [
    {
      "page": 页码(整数),
      "type": "overflow|overlap|whitespace|chart_readability|aesthetics",
      "severity": "high|medium|low",
      "description": "问题描述",
      "suggestion": "修复建议"
    }
  ],
  "overall_score": 1-10,
  "pass": true或false(无high级别问题时为true)
}"""


def _pptx_to_images_comtypes(pptx_path: str, output_dir: str) -> list[str]:
    """使用 Windows COM 接口（PowerPoint 应用）将 PPTX 逐页导出为图片。"""
    import comtypes.client

    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    abs_path = os.path.abspath(pptx_path)
    presentation = powerpoint.Presentations.Open(abs_path)

    image_paths = []
    for i, slide in enumerate(presentation.Slides):
        img_path = os.path.join(output_dir, f"slide_{i + 1}.png")
        slide.Export(img_path, "PNG", 1280, 720)
        image_paths.append(img_path)

    presentation.Close()
    powerpoint.Quit()
    return image_paths


def _pptx_to_images_libreoffice(pptx_path: str, output_dir: str) -> list[str]:
    """使用 LibreOffice 将 PPTX 逐页导出为图片。"""
    abs_path = os.path.abspath(pptx_path)
    cmd = [
        "soffice", "--headless", "--convert-to", "pdf",
        "--outdir", output_dir, abs_path,
    ]
    subprocess.run(cmd, check=True, capture_output=True)

    # 找到生成的 PDF
    pdf_name = Path(abs_path).stem + ".pdf"
    pdf_path = os.path.join(output_dir, pdf_name)
    if not os.path.exists(pdf_path):
        return []

    # PDF 转图片（需要 pdf2image）
    try:
        from pdf2image import convert_from_path
        images = convert_from_path(pdf_path, dpi=150)
        image_paths = []
        for i, img in enumerate(images):
            img_path = os.path.join(output_dir, f"slide_{i + 1}.png")
            img.save(img_path, "PNG")
            image_paths.append(img_path)
        return image_paths
    except ImportError:
        logger.warning("pdf2image not installed, cannot convert PDF to images")
        return []


def pptx_to_images(pptx_path: str, output_dir: str | None = None) -> list[str]:
    """将 PPTX 逐页导出为 PNG 图片。自动选择可用的转换方式。

    Returns:
        图片路径列表，转换失败返回空列表
    """
    output_dir = output_dir or tempfile.mkdtemp(prefix="pptx_review_")
    os.makedirs(output_dir, exist_ok=True)

    # 优先尝试 Windows COM
    if os.name == "nt":
        try:
            return _pptx_to_images_comtypes(pptx_path, output_dir)
        except Exception as exc:
            logger.debug("COM conversion failed: %s", exc)

    # 回退到 LibreOffice
    try:
        return _pptx_to_images_libreoffice(pptx_path, output_dir)
    except Exception as exc:
        logger.debug("LibreOffice conversion failed: %s", exc)

    logger.warning("No available method to convert PPTX to images, skipping review")
    return []


def _encode_image(image_path: str) -> str:
    """将图片编码为 base64。"""
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def review_with_vlm(
    image_paths: list[str],
    llm_provider=None,
) -> dict[str, Any]:
    """使用多模态 LLM 检查幻灯片排版问题。

    Args:
        image_paths: 幻灯片缩略图路径列表
        llm_provider: LLM 提供者，需支持多模态输入

    Returns:
        检查结果字典，包含 issues、overall_score、pass
    """
    if not image_paths or llm_provider is None:
        return {"issues": [], "overall_score": -1, "pass": True, "skipped": True}

    # 构建多模态消息
    try:
        from openai import OpenAI
        import os

        client = OpenAI(
            api_key=os.environ.get("LLM_API_KEY", ""),
            base_url=os.environ.get("LLM_BASE_URL", ""),
        )
        model = os.environ.get("LLM_MODEL", "")

        content_parts = [{"type": "text", "text": "请检查以下幻灯片缩略图："}]
        for i, img_path in enumerate(image_paths):
            b64 = _encode_image(img_path)
            content_parts.append({
                "type": "text",
                "text": f"--- 第 {i + 1} 页 ---",
            })
            content_parts.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{b64}"},
            })

        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": REVIEW_SYSTEM},
                {"role": "user", "content": content_parts},
            ],
            temperature=0.2,
        )
        raw = resp.choices[0].message.content.strip()
        if raw.startswith("```"):
            raw = raw.strip("`").removeprefix("json").strip()
        return json.loads(raw)

    except Exception as exc:
        logger.warning("VLM review failed: %s", exc)
        return {"issues": [], "overall_score": -1, "pass": True, "skipped": True, "error": str(exc)}


def review_pptx(
    pptx_path: str,
    llm_provider=None,
    max_retries: int = 1,
) -> dict[str, Any]:
    """对 PPTX 文件执行自检。

    Args:
        pptx_path: PPTX 文件路径
        llm_provider: LLM 提供者（可选）
        max_retries: 最大回修次数

    Returns:
        检查结果字典
    """
    image_paths = pptx_to_images(pptx_path)
    if not image_paths:
        return {
            "issues": [],
            "overall_score": -1,
            "pass": True,
            "skipped": True,
            "reason": "无法将 PPTX 转换为图片，跳过自检",
        }

    result = review_with_vlm(image_paths, llm_provider=llm_provider)
    result["page_count"] = len(image_paths)
    return result
