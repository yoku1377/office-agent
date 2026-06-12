"""Smoke-check docx templates by rendering sample data."""

from __future__ import annotations

import sys
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from skills.generate_docx.render import render_docx  # noqa: E402


SAMPLE = {
    "title": "关于开展红途智盒 IB-200 现场巡检工作的通知",
    "recipient": "各部门",
    "paragraphs": [
        "为进一步加强设备运行安全管理，现就开展现场巡检工作有关事项通知如下。"
    ],
    "sections": [
        {"heading": "一、巡检范围", "items": ["覆盖已部署红途智盒 IB-200 的相关现场。"]},
        {"heading": "二、工作要求", "items": ["请各部门提前准备，发现问题及时反馈。"]},
    ],
    "closing": "请按要求组织落实。",
    "signer": "综合管理部",
    "date": "2026年6月12日",
}


def check_template(template_path: Path, out_dir: Path) -> tuple[bool, str]:
    out_path = out_dir / f"{template_path.stem}.docx"
    rendered = render_docx(SAMPLE, str(out_path), template_path=str(template_path))
    if rendered["engine"] != "docxtpl":
        return False, f"fallback: {rendered.get('template_error')}"

    text = "\n".join(p.text for p in Document(out_path).paragraphs)
    if any(tag in text for tag in ("{{", "}}", "{%", "%}")):
        return False, "template tags remain in output"
    return True, "ok"


def main() -> int:
    template_root = ROOT / "assets" / "templates" / "docx"
    templates = sorted(template_root.rglob("*.docx"))
    if not templates:
        print("no templates found")
        return 1

    failed = 0
    with TemporaryDirectory() as tmp:
        out_dir = Path(tmp)
        for template_path in templates:
            ok, message = check_template(template_path, out_dir)
            rel = template_path.relative_to(ROOT)
            print(f"{'ok' if ok else 'fail'} {rel}: {message}")
            failed += 0 if ok else 1
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
