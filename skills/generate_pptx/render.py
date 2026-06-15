"""渲染：python-pptx 按版式填充内容，图表使用 PPT 原生图表（可二次编辑）。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches, Pt

from skills.generate_pptx.layout_parser import get_layout_by_name

# 图表类型映射
CHART_TYPE_MAP = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar_clustered": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE,
    "pie": XL_CHART_TYPE.PIE,
    "bar_stacked": XL_CHART_TYPE.COLUMN_STACKED,
}


def _find_placeholder(slide, ph_type_id: int):
    """在幻灯片中查找指定类型的占位符。"""
    for ph in slide.placeholders:
        if ph.placeholder_format.type == ph_type_id:
            return ph
    return None


def _find_placeholder_by_idx(slide, idx: int):
    """在幻灯片中查找指定索引的占位符。"""
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph
    return None


def _fill_title(slide, title_text: str) -> None:
    """填充标题占位符。"""
    if not title_text:
        return
    # 尝试标题占位符
    ph = _find_placeholder(slide, PP_PLACEHOLDER.TITLE)
    if ph is None:
        ph = _find_placeholder(slide, PP_PLACEHOLDER.CENTER_TITLE)
    if ph is not None:
        ph.text = title_text
    else:
        # 尝试 idx=0
        ph = _find_placeholder_by_idx(slide, 0)
        if ph is not None:
            ph.text = title_text


def _fill_subtitle(slide, subtitle_text: str) -> None:
    """填充副标题占位符。"""
    if not subtitle_text:
        return
    ph = _find_placeholder(slide, PP_PLACEHOLDER.SUBTITLE)
    if ph is not None:
        ph.text = subtitle_text


def _fill_body(slide, bullets: list[str]) -> None:
    """填充正文占位符（要点列表）。"""
    if not bullets:
        return
    ph = _find_placeholder(slide, PP_PLACEHOLDER.BODY)
    if ph is None:
        ph = _find_placeholder(slide, PP_PLACEHOLDER.OBJECT)
    if ph is None:
        # 尝试 idx=1
        ph = _find_placeholder_by_idx(slide, 1)
    if ph is None:
        return

    tf = ph.text_frame
    tf.clear()
    for i, bullet in enumerate(bullets):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = str(bullet)
        p.level = 0


def _add_chart(slide, chart_data: dict, left=None, top=None, width=None, height=None) -> None:
    """在幻灯片中添加原生图表。"""
    chart_type_str = chart_data.get("type", "bar")
    chart_type = CHART_TYPE_MAP.get(chart_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED)

    categories = chart_data.get("categories", [])
    series_list = chart_data.get("series", [])

    data = CategoryChartData()
    data.categories = categories
    for s in series_list:
        data.add_series(s.get("name", ""), s.get("values", []))

    # 默认图表位置和大小
    if left is None:
        left = Inches(1.0)
    if top is None:
        top = Inches(2.0)
    if width is None:
        width = Inches(8.0)
    if height is None:
        height = Inches(4.5)

    chart_frame = slide.shapes.add_chart(
        chart_type, left, top, width, height, data
    )
    chart = chart_frame.chart
    chart.has_legend = True

    # 设置图表标题
    chart_title = chart_data.get("title", "")
    if chart_title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = chart_title


def _fill_notes(slide, notes_text: str) -> None:
    """填充备注栏讲稿。"""
    if not notes_text:
        return
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text


def render_slide(prs: Presentation, page: dict, layout_name: str | None = None) -> None:
    """渲染单页幻灯片。"""
    target_layout_name = layout_name or page.get("layout_name", "")

    layout = get_layout_by_name(prs, target_layout_name)
    if layout is None:
        layout = prs.slide_layouts[0]

    slide = prs.slides.add_slide(layout)

    # 填充标题
    _fill_title(slide, page.get("title", ""))

    # 填充副标题
    _fill_subtitle(slide, page.get("subtitle", ""))

    # 填充要点
    bullets = page.get("bullets")
    if bullets:
        _fill_body(slide, bullets)

    # 添加图表
    chart = page.get("chart")
    if chart:
        _add_chart(slide, chart)

    # 填充备注
    _fill_notes(slide, page.get("notes", ""))


def render_pptx(
    data: dict,
    out_path: str,
    template_path: str | None = None,
) -> dict:
    """将结构化大纲渲染为 PPTX 文件。

    Args:
        data: LLM 生成的大纲 JSON
        out_path: 输出文件路径
        template_path: PPT 模板路径（可选）

    Returns:
        包含输出路径和渲染信息的字典
    """
    path = Path(out_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    # 加载模板或创建空白演示文稿
    if template_path and Path(template_path).exists():
        prs = Presentation(str(template_path))
    else:
        prs = Presentation()

    pages = data.get("pages", [])
    for page in pages:
        render_slide(prs, page)

    prs.save(str(path))

    return {
        "out": str(path),
        "page_count": len(pages),
        "template_used": bool(template_path and Path(template_path).exists()),
    }
