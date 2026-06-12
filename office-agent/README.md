# office-agent — 办公文档智能化平台（v0）

小公司内部文档平台：PPT 生成 / PPT 播报 / 公文（生成·纪要·润色）。
第一版全部模型能力走线上 API（阿里云百炼），本地零模型部署。

## 当前已实现
- **公文润色 skill**（skills/polish）：三档润色，Word 修订标记写回 + 批注理由 + 术语保护
- **Provider 适配层**（app/providers）：LLM/ASR/TTS 抽象接口，百炼实现 + 离线 Mock
- **HTTP 服务骨架**（app/main.py）：POST /skills/polish 上传 docx 直接返回润色件

## 快速开始
```bash
pip install -r requirements.txt
cp .env.example .env   # 填入 DASHSCOPE_API_KEY

# 命令行润色
python -m skills.polish.polish 文件.docx --level medium

# 或起服务
uvicorn app.main:app --port 8080
curl -F "file=@文件.docx" -F "level=medium" http://localhost:8080/skills/polish -o 润色件.docx
```

## 离线验证（无需 Key）
```bash
python tests/make_sample.py
MOCK_LLM_FILE=tests/sample/mock_response.json \
  python -m skills.polish.polish tests/sample/sample_notice.docx \
  --out tests/sample/out.docx
```
样例含一个故意改动术语（Ascend→昇腾）的用例，应被校验闸拒绝。

## 目录
```
app/            编排服务（FastAPI 骨架 + provider 适配层）
skills/polish/  润色 skill（SKILL.md / prompts / writeback）
assets/terms/   术语保护表    assets/templates/  PPT母版与公文模板（待填）
tests/          样例与离线验证
```

## 路线（对应方案文档第五节）
1. ✅ 润色  2. 公文生成（docxtpl 模板）  3. PPT 生成  4. 会议纪要  5. PPT 播报
v1 引入 Claude Agent SDK 统一编排各 skill 与 Redis 任务队列。

## v0 已知约束
- 润色按段落内子串定位，不支持跨段修改；段内混合字符格式会统一为首 run 格式
- HTTP 接口为同步处理，并发量上来后切任务队列
