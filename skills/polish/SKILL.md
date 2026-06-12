---
name: polish
description: 公文润色。对 Word 文档(.docx)按轻/中/重三档提出修改建议，以 Word 修订标记（可逐条接受/拒绝）写回原文件，每处改动附批注说明理由，术语表词汇保护不改。当用户要求"润色/校对/修改文风/检查错别字"且对象是 docx 时使用本 skill。
---

# 公文润色 skill

## 流程
1. 提取段落（含编号），加载术语表（assets/terms/terms.yaml）
2. 按档位构建提示词调用 LLM，要求输出 {para_index, old, new, reason} JSON
3. 三道校验闸：段落存在 / old 可在原文逐字定位 / 术语未被改动，不合格建议直接丢弃
4. 修订写回：w:del(原文) + w:ins(新文) + 批注(理由)，作者名"AI润色"

## 使用
    python -m skills.polish.polish 文件.docx --level medium

## 档位
- light：错别字、标点、明显语法
- medium：+ 句式通顺、消歧义、段内衔接（默认）
- heavy：+ 按公文语体重写句子，不得增删事实

## 已知约束（v0）
- 修改片段跨段落不支持；段内多处修改按出现顺序应用
- 重建段落保留段落样式与首 run 字符格式；段内混合格式（局部加粗等）会被统一为首 run 格式，公文场景影响极小
