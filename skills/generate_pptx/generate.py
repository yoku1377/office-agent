"""PPT 生成 skill 主流程：母版解析 → 大纲生成 → 渲染 → 自检 → 回修。"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys

import yaml

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))
from app.providers.llm import get_llm  # noqa: E402
from skills.generate_pptx.layout_parser import build_layout_menu, parse_template  # noqa: E402
from skills.generate_pptx.prompts import build_prompt, parse_generated_json  # noqa: E402
from skills.generate_pptx.render import render_pptx  # noqa: E402
from skills.generate_pptx.review import review_pptx  # noqa: E402

logger = logging.getLogger(__name__)


def load_terms(path):
    if not path or not os.path.exists(path):
        return []
    data = yaml.safe_load(open(path, encoding="utf-8")) or {}
    return list(data.get("terms", []))


def validate_generated(data: dict) -> None:
    """校验 LLM 生成的大纲 JSON 基本结构。"""
    if not isinstance(data, dict):
        raise ValueError("LLM output must be a JSON object")
    if not data.get("title"):
        raise ValueError("Generated PPT is missing title")
    pages = data.get("pages")
    if not isinstance(pages, list):
        raise ValueError("pages must be a list")
    if len(pages) == 0:
        raise ValueError("pages must not be empty")
    for i, page in enumerate(pages):
        if not isinstance(page, dict):
            raise ValueError(f"page[{i}] must be a JSON object")
        if not page.get("title") and not page.get("layout_name"):
            raise ValueError(f"page[{i}] must have title or layout_name")


def generate_pptx(
    brief: str,
    template_path: str | None = None,
    terms_path: str | None = None,
    terms: list[str] | None = None,
    out_path: str | None = None,
    skip_review: bool = False,
    max_review_retries: int = 2,
) -> dict:
    """PPT 生成主流程。

    步骤：
    1. 母版解析：扫描模板提取版式菜单
    2. 大纲生成：LLM 按版式菜单输出结构化 JSON
    3. 渲染：python-pptx 按版式填充内容
    4. 自检：缩略图 + 多模态 LLM 检查溢出与排版
    5. 回修：根据自检结果重新渲染（如有 high 级别问题）
    """
    # 1. 母版解析
    layouts = parse_template(template_path) if template_path else []
    layout_menu = build_layout_menu(layouts)

    # 2. 大纲生成
    loaded_terms = load_terms(terms_path)
    all_terms = list(dict.fromkeys(loaded_terms + list(terms or [])))
    system, user = build_prompt(brief, layout_menu, terms=all_terms)
    raw = get_llm().chat(system, user)
    data = parse_generated_json(raw)
    validate_generated(data)

    # 3. 渲染
    out_path = out_path or "generated_presentation.pptx"
    rendered = render_pptx(data, out_path, template_path=template_path)

    result = {
        "out": rendered["out"],
        "title": data.get("title"),
        "page_count": rendered["page_count"],
        "template_used": rendered["template_used"],
        "review": None,
    }

    # 4. 自检
    if not skip_review:
        review_result = review_pptx(rendered["out"], max_retries=max_review_retries)
        result["review"] = review_result

        # 5. 回修：如果有 high severity 问题，尝试重新生成
        if not review_result.get("pass") and max_review_retries > 0:
            high_issues = [
                iss for iss in review_result.get("issues", [])
                if iss.get("severity") == "high"
            ]
            if high_issues:
                logger.info(
                    "Review found %d high-severity issues, regenerating...",
                    len(high_issues),
                )
                # 将问题反馈给 LLM 重新生成，附带具体修复约束
                issues_text = json.dumps(high_issues, ensure_ascii=False, indent=2)
                fix_constraints = """【修复约束】
- 文字溢出：减少该页要点数量（每条不超过30字），或拆分为多页；
- 内容重叠：更换为更宽松的版式，或减少内容量；
- 空白过多：增加要点数量或补充图表数据；
- 图表不可读：减少 categories 和 series 数量；
- 修改时仅调整有问题的页面，其余页面保持不变。"""
                retry_system = system + f"\n\n【上一版自检发现的问题】\n{issues_text}\n{fix_constraints}\n请修正以上问题后重新生成完整大纲。"
                raw_retry = get_llm().chat(retry_system, user)
                data_retry = parse_generated_json(raw_retry)
                try:
                    validate_generated(data_retry)
                    rendered_retry = render_pptx(data_retry, out_path, template_path=template_path)
                    result["out"] = rendered_retry["out"]
                    result["page_count"] = rendered_retry["page_count"]
                    result["review_retry"] = review_pptx(rendered_retry["out"])
                except ValueError as exc:
                    logger.warning("Retry generation validation failed: %s", exc)

    return result


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    ap = argparse.ArgumentParser()
    ap.add_argument("brief", help="事项说明/演示主题")
    ap.add_argument("--template", default=None, help="PPT 模板路径 (.pptx)")
    ap.add_argument("--terms", default="assets/terms/terms.yaml", help="术语表路径")
    ap.add_argument("--out", default=None, help="输出文件路径")
    ap.add_argument("--skip-review", action="store_true", help="跳过自检步骤")
    ap.add_argument("--max-review-retries", type=int, default=2, help="自检回修最大次数")
    args = ap.parse_args()
    result = generate_pptx(
        args.brief,
        template_path=args.template,
        terms_path=args.terms,
        out_path=args.out,
        skip_review=args.skip_review,
        max_review_retries=args.max_review_retries,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
