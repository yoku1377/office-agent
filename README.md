# office-agent — 办公文档智能化平台（v1）

小公司内部文档能力后端：PPT 生成 / PPT 播报 / 公文（生成、纪要、润色）。
第一版模型能力按方案走线上 API（阿里云百炼），本地仅运行编排服务。

## 当前已实现

- **公文润色 skill**（`skills/polish`）：三档润色，Word 修订标记写回 + 批注理由 + 术语保护
- **公文生成 skill**（`skills/generate_docx`）：按事项说明 + 文种规则卡生成 Word 文档，优先使用部门/文种模板
- **公司上下文**（`assets/contexts`）：按部门/场景加载术语表、默认档位、作者名
- **公开样例库**（`assets/examples`）：公开公文来源清单 + 文种风格规则卡，内部样例先留空
- **Provider 适配层**（`app/providers`）：百炼 OpenAI 兼容 LLM + 离线 Mock
- **任务式 HTTP API**（`app/main.py`）：上传 docx 创建任务，查询状态，下载结果
- **skill 发现**（`GET /skills`）：读取各 skill 的 `SKILL.md`

## 快速开始

```bash
pip install -r requirements.txt
cp .env.example .env

uvicorn app.main:app --host 0.0.0.0 --port 8080
```

默认大模型配置：

```env
DASHSCOPE_API_KEY=sk-xxxxxxxx
LLM_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
LLM_MODEL=qwen-max
```

后续如果要切到内网 OpenAI 兼容服务，再通过 `LLM_BASE_URL` / `LLM_MODEL` 调整，不改 skill 代码。

## v1 API

查看可用 skill：

```bash
curl http://localhost:8080/skills
```

查看可用部门上下文与文种：

```bash
curl http://localhost:8080/contexts
```

创建润色任务：

```bash
curl -F "skill=polish" \
  -F "file=@文件.docx" \
  -F "level=medium" \
  -F "department=admin" \
  -F "document_type=notice" \
  -F "user_id=zhangsan" \
  http://localhost:8080/tasks
```

创建公文生成任务：

```bash
curl -F "brief=请起草一份关于开展红途智盒 IB-200 现场巡检工作的通知，巡检时间为6月15日至6月30日，各部门需提前准备并及时反馈问题。" \
  -F "department=admin" \
  -F "document_type=notice" \
  -F "user_id=zhangsan" \
  http://localhost:8080/tasks/generate-docx
```

也可以使用专用路径：

```bash
curl -F "file=@文件.docx" \
  -F "level=medium" \
  -F "department=admin" \
  -F "document_type=notice" \
  http://localhost:8080/tasks/polish
```

查询任务：

```bash
curl http://localhost:8080/tasks/<task_id>
```

下载结果：

```bash
curl http://localhost:8080/tasks/<task_id>/download -o 润色件.docx
```

兼容 v0 的同步接口仍保留：

```bash
curl -F "file=@文件.docx" \
  -F "level=medium" \
  -F "department=admin" \
  -F "document_type=notice" \
  http://localhost:8080/skills/polish -o 润色件.docx
```

## 公司上下文

默认上下文在 `assets/contexts/default.yaml`：

```yaml
name: default
terms_path: assets/terms/terms.yaml
extra_terms: []
polish:
  default_level: medium
  author: AI润色
```

已预置：

- `assets/contexts/admin.yaml`：行政/综合部
- `assets/contexts/marketing.yaml`：营销部门
- `assets/contexts/product.yaml`：产品部门

请求里传 `department=admin`、`department=marketing` 或 `department=product`，服务会自动加载对应上下文；找不到时回退到 `default`。
请求里再传 `document_type=notice`、`opinion`、`action_plan` 或 `approval`，润色 prompt 会自动加载对应文种规则卡。

## 模板

公文生成优先使用 `docxtpl` 模板；如果模板缺失或 `docxtpl` 不可用，会回退到通用 `python-docx` 渲染。

当前已生成种子模板：

```text
assets/templates/docx/admin/
  notice.docx
  opinion.docx
  action_plan.docx
  approval.docx
assets/templates/docx/marketing/
  action_plan.docx
  opinion.docx
assets/templates/docx/product/
  action_plan.docx
```

模板路径由 `assets/contexts/*.yaml` 配置。后续如果有正式公司模板，直接覆盖对应 `.docx` 文件即可。

重建种子模板：

```bash
python scripts/build_docx_templates.py
```

检查所有模板能否正常渲染：

```bash
python scripts/check_docx_templates.py
```

## 公开样例库

公开样例放在 `assets/examples/public`：

- `official_sources.yaml`：公开权威来源清单，只记录来源、文种、适用场景和可学习点
- `style_cards/notice.yaml`：通知
- `style_cards/opinion.yaml`：意见
- `style_cards/action_plan.yaml`：方案 / 行动计划
- `style_cards/approval.yaml`：批复

内部样例放在 `assets/examples/internal`，目前先留空。等确认哪些历史文档可以复用后，再按部门和文种加入。

## 离线验证（无需 Key）

```bash
python tests/make_sample.py
MOCK_LLM_FILE=tests/sample/mock_response.json \
  python -m skills.polish.polish tests/sample/sample_notice.docx \
  --out tests/sample/out.docx
```

样例含一个故意改动术语（`Ascend -> 昇腾`）的用例，应被校验闸拒绝。

## 目录

```text
app/              FastAPI 服务、任务存储、上下文加载、provider 适配层
skills/polish/    润色 skill（SKILL.md / prompts / writeback）
skills/generate_docx/
                  公文生成 skill（structured JSON / docxtpl templates）
assets/contexts/  公司/部门上下文
assets/examples/  公开样例库与内部样例占位
assets/terms/     术语保护表
assets/templates/ 部门/文种模板
storage/          上传文件、输出文件、任务状态
tests/            样例与离线验证
```

## 路线

1. 已完成：润色 skill、术语保护、Word 修订写回
2. 已完成：v1 任务 API、上下文配置、skill 清单、公开样例库
3. 已完成：公文生成 skill 初版（docxtpl 模板优先，python-docx 回退）
4. 下一步：接入 Redis/RQ 或 Celery，替换当前文件型任务队列
5. 下一步：替换正式公司 docxtpl 模板、PPT 生成、会议纪要、PPT 播报
6. 后续：OpenClaw skill 作为入口，调用本服务 API
