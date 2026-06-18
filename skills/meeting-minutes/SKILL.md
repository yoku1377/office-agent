---
name: meeting-minutes
description: "会议纪要生成。将音频通过可用的语音模型转写，LLM 提取要点，输出 Word 文档。"
---

# Meeting Minutes — 会议纪要生成

## 概述

将会议音频自动转写为文字（含逐句时间戳和 speaker 编号），由 LLM 提取要点并推断发言人身份供用户一键确认，最终输出结构化 Word 文档。

| 阶段 | 方式 | 说明 |
|------|------|------|
| 音频定位 | 文件系统 | 在 `~/.openclaw/media/inbound/` 查找音频 |
| 语音转写 | 配置中可用的音频模型 | 自动检测并使用支持音频输入的模型 |
| 发言人推断 | 当前会话 LLM | 根据对话内容推测各 speaker 真实姓名/职务，由用户确认 |
| 纪要生成 | 当前会话 LLM | 由 agent 主模型生成 |
| 文档输出 | docx skill | 生成 .docx 文件 |

---

## 触发条件

当用户提供音频文件（`.wav` / `.mp3` / `.m4a` / `.flac` / `.webm` / `.ogg` / `.opus` / `.aac`），或说"帮我生成会议纪要"、"转写这段录音"、"把录音整理成纪要"并附带音频时触发。

---

## 工作流

### 第一步：定位音频文件

1. 用户可能提供绝对路径或文件名关键词
2. 若未提供完整路径，在 `~/.openclaw/media/inbound/` 目录下搜索匹配的音频文件（按最新修改时间优先）
3. 确认文件存在、格式受支持

### 第二步：ASR 语音转写（逐句 + speaker 编号）

**核心原则：不绑定任何特定模型或提供商，从配置中自动发现可用的音频模型。**

#### 2.1 发现可用音频模型

读取 openclaw.json，查找 input 包含 `audio` 的模型：

```python
import json, os

config_path = os.path.expanduser('~/.openclaw/openclaw.json')
with open(config_path, 'r', encoding='utf-8-sig') as f:
    config = json.load(f)

audio_models = []
for provider_name, provider_cfg in config.get('models', {}).get('providers', {}).items():
    for model in provider_cfg.get('models', []):
        if 'audio' in model.get('input', []):
            audio_models.append({
                'provider': provider_name,
                'model_id': model['id'],
                'base_url': provider_cfg.get('baseUrl', ''),
                'api_key': provider_cfg.get('apiKey', ''),
                'api_type': provider_cfg.get('api', '')
            })

if not audio_models:
    raise Exception("配置中未找到支持音频输入的模型，请在 openclaw.json 中添加支持 audio input 的模型。")
```

#### 2.2 调用音频模型进行逐句转写

使用标准的 OpenAI-compatible API 格式，要求返回逐句、带 speaker 编号和时间戳的结果：

```python
import base64, requests, re

model_info = audio_models[0]
audio_path = "<audio_path>"

with open(audio_path, 'rb') as f:
    audio_b64 = base64.b64encode(f.read()).decode('utf-8')

ext = audio_path.rsplit('.', 1)[-1].lower()
mime_map = {
    'wav': 'audio/wav', 'mp3': 'audio/mpeg', 'm4a': 'audio/mp4',
    'flac': 'audio/flac', 'webm': 'audio/webm', 'ogg': 'audio/ogg',
    'opus': 'audio/opus', 'aac': 'audio/aac'
}
mime_type = mime_map.get(ext, f'audio/{ext}')

url = f"{model_info['base_url'].rstrip('/')}/chat/completions"
headers = {
    "Authorization": f"Bearer {model_info['api_key']}",
    "Content-Type": "application/json"
}

# 要求逐句转写 + speaker 编号 + 时间戳
asr_prompt = (
    "请逐句转写这段会议录音。每句话输出一行，格式如下：\n"
    "[speaker_N] MM:SS.sss --> MM:SS.sss 说话内容\n\n"
    "要求：\n"
    "1. speaker_N：同一说话人用同一编号，不同人用不同编号\n"
    "2. 时间戳范围：起始时间 --> 结束时间（从录音开始计算）\n"
    "3. 逐句完整转写，不要合并句子"
)

payload = {
    "model": model_info['model_id'],
    "messages": [{
        "role": "user",
        "content": [
            {
                "type": "input_audio",
                "input_audio": {
                    "data": f"data:{mime_type};base64,{audio_b64}",
                    "format": ext
                }
            },
            {"type": "text", "text": asr_prompt}
        ]
    }]
}

response = requests.post(url, headers=headers, json=payload, timeout=200)
response.raise_for_status()
transcript = response.json()['choices'][0]['message']['content']
```

#### 2.3 解析为结构化数据

```python
pattern = re.compile(r'\[(speaker_\d+)\]\s+(\S+(?:\s*-->\s*\S+)?)\s+(.*)')
utterances = []
for line in transcript.strip().split('\n'):
    line = line.strip()
    m = pattern.match(line)
    if m:
        utterances.append({
            'speaker': m.group(1),
            'timestamp': m.group(2),
            'text': m.group(3)
        })
```

#### 2.4 转写验证

- 若 `len(utterances) == 0` 或所有文本总长度 < 50 字符：提示用户音频质量可能不佳
- 若 API 调用失败：检查 apiKey、模型可用性、网络连通性
- exec timeout 设置 200+（音频转写可能需要 60-180 秒）

**注意**：如果 OpenAI-compatible 格式调用失败，检查模型的 `api_type` 字段。某些提供商可能需要使用 native SDK。

### 第三步：发言人身分推断（LLM + 用户确认）

#### 3.1 LLM 推断

将逐句转写文本交给当前会话 LLM，让模型根据对话内容推断每个 speaker 的可能身份：

```markdown
以下是一次会议录音的逐句转写（含 speaker 编号）：

[speaker_1] 00:00.000 --> 00:02.880 大家好，下面开始项目例会。
[speaker_1] 00:03.240 --> 00:12.160 本周完成了xxx的初步开发...
...

请推断每位 speaker 可能的真实姓名或职务身份。考虑：
- 说话人的语气和口吻（领导/员工/客户/技术负责人等）
- 讨论的内容领域
- 任何对话中提到的称呼线索

输出 JSON 格式：
{"推测": {"speaker_1": "推测姓名或职务"}, "置信度": "高/中/低", "推理依据": "依据说明"}
```

#### 3.2 展示给用户确认

以简洁格式呈现给用户，等待确认或修改：

```markdown
📋 发言人身份推测，请确认或修改（回复"确认"或修正内容）：

  speaker_1 → 项目经理（置信度：高）
  speaker_2 → 产品经理（置信度：中）

如确认无误，直接回复"确认"即可。如需修改，请回复"speaker_1=张伟, speaker_2=李娜"。
```

#### 3.3 用户确认后替换

收到用户确认或修改后，将原文中的 speaker 编号替换为真实姓名，得到干净的对话文本：

```
[项目经理 00:00.000 --> 00:02.880] 大家好，下面开始项目例会。
[项目经理 00:03.240 --> 00:12.160] 本周完成了xxx的初步开发...
```

### 第四步：生成会议纪要

agent 主模型根据带真实姓名的对话文本，生成结构化纪要：

```markdown
# [会议标题] 会议纪要

**会议时间：** YYYY-MM-DD HH:MM
**会议地点：** 
**参会人员：** 
**主持人：** 
**记录人：** AI 助手

---

### 一、会议议题

1. xxx

---

### 二、讨论要点

**议题1：xxx**
- [发言人姓名]：xxx

---

### 三、会议决议

1. xxx

---

### 四、行动计划（Action Items）

| 序号 | 任务描述 | 负责人 | 截止日期 | 优先级 |
|------|----------|--------|----------|--------|
| 1 | xxx | xxx | YYYY-MM-DD | 高/中/低 |

---

### 五、遗留问题 / 待定事项

- [ ] xxx

---

### 六、下次会议安排

**日期：** 
**主题：** 

---
*本纪要由 AI 助手根据录音转写自动生成，如有疏漏请指正。*
*生成时间：YYYY-MM-DD HH:MM*
```

**生成要求**：
- 未提及的信息留空，不编造
- 行动项要具体、可执行
- 调用真实姓名替换后的文本，不要再用 speaker_N
- 输出语言与用户提问语言一致

### 第五步：输出 Word 文档

使用 docx skill 生成 .docx 文件，保存到 workspace 根目录，命名：`meeting_minutes_YYYYMMDD.docx`。

若文档生成失败，回退输出 Markdown 文本。

---

## 输出

1. **逐句转写**（中间产物，可展示给用户或仅用于后续步骤）
2. **发言人身份确认**（交互步骤，需要用户确认/修改）
3. **Word 文档**（必需）— 用 `MEDIA:` 指令附加
4. **纪要预览**（必需）— 在对话中展示全文

---

## 数据流

```
音频文件
    │
    ▼
ASR（音频模型 → 逐句 + speaker_N + 时间戳）
    │
    ▼
结构化解析 → 逐句JSON数组
    │
    ▼
LLM 身分推断 → 推测 speaker_N 对应身份
    │
    ▼
用户确认/修改 → 得到 speaker_N → 真实姓名映射
    │
    ▼
替换 speaker_N 为真实姓名 → 干净对话文本
    │
    ▼
LLM 生成纪要
    │
    ▼
Word 文档输出
```

---

## 错误处理

| 场景 | 处理 |
|------|------|
| 音频文件不存在 | 提示检查路径 |
| 音频格式不支持 | 列出支持的格式 |
| 无可用音频模型 | 提示在 openclaw.json 中添加支持 audio input 的模型 |
| ASR 未返回有效转写 | 提示音频质量可能不佳 |
| ASR API 调用失败 | 检查 apiKey、模型是否支持音频、网络连通性 |
| 身份推断置信度过低 | 提示用户自行判断，提供对话文本作为参考 |
| 用户长时间未确认身份 | 保留 speaker_N 编号生成纪要，标注"待确认" |
| 文档生成失败 | 回退输出 Markdown |

---

## 注意事项

- 录音时长建议 ≤ 1 小时
- ASR 耗时约 60-180 秒，exec timeout 建议 200+
- **不硬编码任何模型名称、提供商或 API 端点**
- 始终从 openclaw.json 动态发现可用模型
- 身分推断由当前会话 LLM 执行，不额外调用其他模型
- 用户确认步骤是可选但推荐的，能显著提升纪要质量
