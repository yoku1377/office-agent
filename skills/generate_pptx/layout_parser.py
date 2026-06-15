"""母版解析：扫描 PPT 模板，提取各版式占位符清单，作为大纲生成的"版式菜单"。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER


@dataclass
class PlaceholderInfo:
    idx: int
    name: str
    type_name: str
    type_id: int


@dataclass
class LayoutInfo:
    name: str
    placeholders: list[PlaceholderInfo] = field(default_factory=list)


# python-pptx 占位符类型映射
_PLACEHOLDER_TYPE_NAMES = {
    PP_PLACEHOLDER.TITLE: "title",
    PP_PLACEHOLDER.BODY: "body",
    PP_PLACEHOLDER.CENTER_TITLE: "center_title",
    PP_PLACEHOLDER.SUBTITLE: "subtitle",
    PP_PLACEHOLDER.OBJECT: "object",
    PP_PLACEHOLDER.CHART: "chart",
    PP_PLACEHOLDER.TABLE: "table",
    PP_PLACEHOLDER.PICTURE: "picture",
    PP_PLACEHOLDER.ORG_CHART: "org_chart",
    PP_PLACEHOLDER.MEDIA_CLIP: "media_clip",
    PP_PLACEHOLDER.SLIDE_NUMBER: "slide_number",
    PP_PLACEHOLDER.HEADER: "header",
    PP_PLACEHOLDER.FOOTER: "footer",
    PP_PLACEHOLDER.DATE: "date",
}


def _placeholder_type_name(ph_type: int) -> str:
    return _PLACEHOLDER_TYPE_NAMES.get(ph_type, f"unknown({ph_type})")


def parse_template(template_path: str | Path) -> list[LayoutInfo]:
    """解析 PPT 模板，返回所有版式及其占位符信息。"""
    path = Path(template_path)
    if not path.exists():
        return []

    prs = Presentation(str(path))
    layouts: list[LayoutInfo] = []

    for layout in prs.slide_layouts:
        info = LayoutInfo(name=layout.name)
        for ph in layout.placeholders:
            info.placeholders.append(
                PlaceholderInfo(
                    idx=ph.placeholder_format.idx,
                    name=ph.name,
                    type_name=_placeholder_type_name(ph.placeholder_format.type),
                    type_id=ph.placeholder_format.type,
                )
            )
        layouts.append(info)

    return layouts


def build_layout_menu(layouts: list[LayoutInfo]) -> str:
    """将版式信息格式化为 LLM 可读的版式菜单文本。"""
    if not layouts:
        return _default_layout_menu()

    lines = []
    for i, layout in enumerate(layouts, 1):
        ph_desc = ", ".join(
            f"{ph.name}(idx={ph.idx}, type={ph.type_name})"
            for ph in layout.placeholders
        )
        lines.append(f"  {i}. \"{layout.name}\" — 占位符: [{ph_desc}]" if ph_desc
                     else f"  {i}. \"{layout.name}\" — 无占位符")
    return "\n".join(lines)


def _default_layout_menu() -> str:
    """无模板时的默认版式菜单。"""
    return """  1. "Title Slide" — 占位符: [Title 1(idx=0, type=title), Subtitle 2(idx=1, type=subtitle)]
  2. "Title and Content" — 占位符: [Title 1(idx=0, type=title), Content Placeholder 2(idx=1, type=body)]
  3. "Section Header" — 占位符: [Title 1(idx=0, type=title), Text Placeholder 2(idx=1, type=body)]
  4. "Two Content" — 占位符: [Title 1(idx=0, type=title), Content Placeholder 2(idx=1, type=body), Content Placeholder 3(idx=2, type=body)]
  5. "Comparison" — 占位符: [Title 1(idx=0, type=title), Content Placeholder 2(idx=1, type=body), Content Placeholder 3(idx=2, type=body)]
  6. "Title Only" — 占位符: [Title 1(idx=0, type=title)]
  7. "Blank" — 无占位符
  8. "Content with Caption" — 占位符: [Title 1(idx=0, type=title), Content Placeholder 2(idx=1, type=body), Caption Placeholder 3(idx=2, type=body)]
  9. "Picture with Caption" — 占位符: [Title 1(idx=0, type=title), Picture Placeholder 2(idx=1, type=picture), Caption Placeholder 3(idx=2, type=body)]"""


def get_layout_by_name(prs: Presentation, name: str):
    """根据名称查找版式，找不到则回退到第一个版式。"""
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    # 回退：尝试模糊匹配
    for layout in prs.slide_layouts:
        if name.lower() in layout.name.lower():
            return layout
    # 最终回退到第一个版式
    return prs.slide_layouts[0] if prs.slide_layouts else None
