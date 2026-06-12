"""润色提示词：三档 + 术语保护 + 严格 JSON 输出。"""

LEVELS = {
    "light": "仅修正错别字、标点误用、明显语法错误和格式不规范，不调整句式与表达。",
    "medium": "在 light 基础上，优化不通顺的句式、消除歧义与冗余、改善段内逻辑衔接，保持原意与篇幅相近。",
    "heavy": "在 medium 基础上，可按目标文体重组句子与段内结构，使表达更符合规范公文语体，但不得增删事实信息。",
}

SYSTEM = """你是一名资深公文校对与润色专家。对用户提供的逐段文本提出修改建议。

润色档位要求：{level_desc}

【术语保护】以下词汇一字不改，原样保留：{terms}

{style_reference}

【输出格式】只输出 JSON 数组，不要任何其他文字或代码块标记。每个元素：
{{"para_index": 段落编号(整数), "old": "段落中需要修改的原文片段(必须与原文逐字一致)",
  "new": "修改后的文字", "reason": "一句话说明修改理由"}}

【硬性规则】
- old 必须是该段落文本的连续子串，逐字一致（含标点）；
- 同一段落多处修改输出多个元素，按出现顺序排列；
- 没有需要修改之处的段落不要输出；
- 不要把整段作为 old 除非确需整段重写（heavy 档才允许）。"""

USER = """以下是待润色文档的段落（格式：[编号] 内容）：

{paragraphs}

请按要求输出修改建议 JSON。"""


def _as_lines(values):
    return "\n".join(f"- {v}" for v in values) if values else "- （无）"


def _format_style_reference(style_card):
    if not style_card:
        return "【文种参考】未指定文种，按通用公司公文处理。"

    structure = style_card.get("structure") or []
    structure_lines = []
    for item in structure:
        title = item.get("title", "")
        rule = item.get("rule", "")
        structure_lines.append(f"- {title}：{rule}")

    return """【文种参考】
文种：{name}（{document_type}）
用途：{purpose}
常见结构：
{structure}
语气要求：
{tone}
润色规则：
{polish_rules}
注意：文种参考只用于结构、语气和表达规范，不得据此新增事实、时间、数字、责任人或业务结论。""".format(
        name=style_card.get("name", "未命名文种"),
        document_type=style_card.get("document_type", "unknown"),
        purpose=style_card.get("purpose", "（无）"),
        structure="\n".join(structure_lines) if structure_lines else "- （无）",
        tone=_as_lines(style_card.get("tone") or []),
        polish_rules=_as_lines(style_card.get("polish_rules") or []),
    )


def build_prompt(paragraphs, level, terms, style_card=None):
    level_desc = LEVELS[level]
    terms_str = "、".join(terms) if terms else "（无）"
    body = "\n".join(f"[{i}] {t}" for i, t in paragraphs)
    return (
        SYSTEM.format(
            level_desc=level_desc,
            terms=terms_str,
            style_reference=_format_style_reference(style_card),
        ),
        USER.format(paragraphs=body),
    )
