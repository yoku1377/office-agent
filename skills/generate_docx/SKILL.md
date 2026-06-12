---
name: generate_docx
description: 公文生成。根据用户提供的事项说明、部门上下文和文种规则卡，生成通知、意见、方案/行动计划、批复等 Word 文档。第一版使用结构化 JSON + python-docx 渲染，后续可替换为 docxtpl 模板。
---

# 公文生成 skill

## 流程

1. 加载公司/部门上下文、术语表和文种规则卡
2. LLM 根据用户事项说明输出结构化 JSON
3. 校验 JSON 基本结构，不合格则失败
4. 用 `python-docx` 渲染为 `.docx`

## 输入

- `brief`：用户提供的事项说明
- `document_type`：`notice` / `opinion` / `action_plan` / `approval`
- `department`：由上层服务加载上下文

## 输出

Word 文档，包含标题、主送对象、正文段落、分节内容、落款和日期。

## 已知约束（v1）

- 暂未使用真实公司 docxtpl 模板
- 不自动生成红头、版记、页码等正式机关公文版式
- 模型不得编造用户未提供的时间、数字、责任人、审批结论
