"""Prompts for structured official document generation."""

import json


SYSTEM = """你是一名资深公司公文写作助手。请根据用户提供的事项说明，生成结构化公文内容。

【文种参考】
文种：{document_name}（{document_type}）
用途：{purpose}
常见结构：
{structure}
语气要求：
{tone}
生成规则：
{generation_rules}

【术语保护】以下词汇一字不改，原样保留：{terms}

【硬性规则】
- 不得编造用户未提供的事实、时间、数字、责任人、客户名称、政策依据或审批结论；
- 信息不足时，在对应字段中使用“待补充”，不要自行发挥；
- 输出必须是 JSON 对象，不要输出代码块标记或解释文字；
- 正文应适合公司内部使用，正式但不过度机关化。

【输出 JSON 结构】
{{
  "title": "标题",
  "recipient": "主送对象，没有则为待补充",
  "paragraphs": ["导语或正文段落"],
  "sections": [
    {{"heading": "一、章节标题", "items": ["条目1", "条目2"]}}
  ],
  "closing": "结尾要求或待补充",
  "signer": "落款，没有则为待补充",
  "date": "日期，没有则为待补充"
}}"""

USER = """用户事项说明：

{brief}

请生成结构化公文 JSON。"""


def _as_lines(values):
    return "\n".join(f"- {v}" for v in values) if values else "- （无）"


def _format_structure(style_card):
    lines = []
    for item in style_card.get("structure") or []:
        lines.append(f"- {item.get('title', '')}：{item.get('rule', '')}")
    return "\n".join(lines) if lines else "- （无）"


def build_prompt(brief, document_type, terms, style_card=None):
    style_card = style_card or {}
    system = SYSTEM.format(
        document_name=style_card.get("name") or document_type,
        document_type=document_type,
        purpose=style_card.get("purpose") or "（无）",
        structure=_format_structure(style_card),
        tone=_as_lines(style_card.get("tone") or []),
        generation_rules=_as_lines(style_card.get("generation_rules") or []),
        terms="、".join(terms) if terms else "（无）",
    )
    return system, USER.format(brief=brief)


def parse_generated_json(raw):
    raw = raw.strip()
    if raw.startswith("```"):
        raw = raw.strip("`")
        raw = raw.removeprefix("json").strip()
    return json.loads(raw)
