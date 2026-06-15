"""Build a default PPTX template with standard layouts for office use."""

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt


def build_default_template(out_path: str) -> None:
    """创建一个包含标准版式的默认 PPT 模板。"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 模板已自带默认版式，直接保存即可
    # python-pptx 创建的 Presentation 默认包含:
    # Title Slide, Title and Content, Section Header, Two Content,
    # Comparison, Title Only, Blank, Content with Caption, Picture with Caption
    path = Path(out_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    print(f"Default template saved to {path}")


if __name__ == "__main__":
    build_default_template("assets/templates/pptx/default.pptx")
