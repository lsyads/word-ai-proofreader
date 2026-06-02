# Word AI 审校助手

面向出版社责任编辑的 Word AI 审校助手。当前 V1 基线支持在 Word 中选中一段正文做局部审校，也可以选择“全书正文”并上传 `.docx` 文件，由 FastAPI 后端完成全书抽取、分块审校、批注或修订+批注写回，并生成新的 Word 文件。

V2 目标是把项目升级为出版审校 Agent 工作台：Agent 不只是调用一次模型，而是围绕一本书建立审校项目，制定审校计划，调用文档处理工具，跨章节追踪术语、体例和一致性问题，汇总风险和证据，并在编辑确认后写回 Word。

## 项目结构

```text
.
├── backend/      # Python FastAPI 后端服务
├── scripts/      # 本地开发辅助脚本
├── word-addin/   # Office.js + TypeScript + Webpack Word 插件
├── AGENTS.md     # 协作约定
└── spec.md       # API 契约和验收标准
```

## 当前能力（V1 基线）

- 审校范围：当前选区或全书正文。当前选区由插件读取并回写；全书正文上传 `.docx`，后端处理目录可见文本、正文、表格和常见文本框文字。本版不支持 `.doc`，请先另存为 `.docx`。
- 分块规则：当前选区 `> 7000` 字时走后端内存异步任务；全书 `.docx` 先按章拆分，再按节拆分；仍超过 7000 字时再用可提取的目录小标题辅助拆分，最后按现有段落/句末规则切分。
- AI API：支持 OpenAI 兼容 Responses API 和 Chat Completions；插件可调整 temperature，默认 `0.2`；未配置 `AI_API_KEY` 时返回 mock 结果，方便本地联调。
- 书籍信息：插件要求填写书名，介绍可选；后端把书籍信息作为 prompt 背景，但只审校传入正文。
- 结果处理：后端把 AI 输出转换为结构化 `issues[]`，过滤纯空白差异，并按 `original` 计算 `start/end/locator`。
- Word 写回：当前选区审校完成后先展示结果，编辑筛选、勾选并确认后由插件写回批注或修订+批注；全书 `.docx` 由后端直接生成带批注或修订+批注的新文件，插件展示新文件名、保留期限和下载入口。
- 结果保留：全书 `.docx` 结果文件保存在后端 `DOCX_OUTPUT_DIR`，默认至少保留 7 天；历史记录里的下载入口在文件未过期且未被外部清理时可继续下载。
- Agent trace：后端用 LangGraph 编排审校流程，普通审校、分块任务和 DOCX 任务都会生成 `run_id`；Word 插件“运行过程”面板可刷新 trace 摘要，也可通过 `/api/agent/runs/{run_id}/trace` 查看节点、chunk、耗时、错误和重试次数。
- 历史记录：插件在本地保存最近 20 条新 schema 历史，支持清空、导出 JSON、导入 JSON；历史不保存完整正文或原始 DOCX。

完整 API 契约见 [spec.md](spec.md)，V2 Agent 工作台工程约束见 [AGENTS.md](AGENTS.md)。

## V2 目标：出版审校 Agent 工作台

V2 可以不兼容 V1 存量数据，包括本地历史、Agent trace、任务状态、DOCX 结果索引和旧任务快照；升级时允许重建或清空这些数据。V2 仍应复用 V1 已验证的核心文档处理逻辑，包括 DOCX 解析、章节/分块、AI 审校、结果归一化、原文定位、批注/修订写回和下载文件生成。

V2 的产品目标：

- 项目化审校：围绕一本书建立审校项目，保存审校目标、文档地图、章节进度、待确认问题、编辑决策和审校报告。
- Agent 规划：先生成全书审校计划，再按阶段调用工具，而不是把所有正文塞给一个 prompt。
- 文档地图：抽取章节、目录可见文本、正文、表格和文本框文字，建立可追踪的文档结构和位置映射。
- 出版规范记忆：沉淀本书术语、人名地名、机构名、体例规则、编辑确认过的偏好和跨章节一致性线索。
- 多轮复核：对候选问题做归并、去重、自检、证据绑定和风险分级，降低重复建议和幻觉建议。
- 人机协同写回：Agent 只生成候选建议和证据，编辑批准、拒绝或暂缓后，系统再批量写回批注或修订。
- 审校报告：输出本轮审校范围、问题分布、风险章节、未处理事项、编辑确认记录和可复查的运行摘要。

V2 路线图：

1. 项目与文档地图：建立 `project_id`、文档结构、章节状态和 V2 历史 schema。
2. 审校计划与工具注册表：让 Agent planner 选择分块审校、术语检查、体例检查、跨章节一致性检查等工具。
3. 记忆与规则库：支持本书级术语/体例记忆，默认不把完整正文写入长期记忆。
4. 复核与确认队列：将候选问题归并、证据绑定、自检后进入编辑确认队列。
5. 写回与报告：按编辑决策生成 Word 批注/修订，并输出可复查的审校报告。

当前代码已实现 V2 插件主界面和工作台闭环：创建当前选区或 DOCX 审校项目、建立文档地图、同步运行 Agent 审校、生成候选问题、批量批准/拒绝/暂缓、写回已批准问题、下载 DOCX 结果、查询审校报告和脱敏 run event trace。插件 UI 不再提供 V1 独立入口；当前选区写回由 Office.js 完成，DOCX 写回由后端完成。V2 第一版不迁移 V1 历史、trace、任务状态或 DOCX result index。

## 环境变量

复制模板后按需填写：

```bash
cp .env.example .env
```

常用配置：

```text
AI_API_KEY=local-omlx-dev-key
MIMO_API_KEY=...
AI_PROVIDER_API=responses
AI_PROFILES_JSON=
OPENAI_API_BASE_URL=http://127.0.0.1:8001/v1
OPENAI_MODEL=Qwen3.6-35B-A3B-4.4bit-msq
AI_REQUEST_TIMEOUT_SECONDS=180
AI_FAST_MAX_TOKENS=8192
AI_THINKING_MAX_TOKENS=16384
BACKEND_HOST=127.0.0.1
BACKEND_PORT=8000
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
DOCX_OUTPUT_DIR=var/docx-results
DOCX_RETENTION_DAYS=7
AGENT_TRACE_DIR=var/agent-traces
AGENT_WORKSPACE_DIR=var/agent-workspace
WORD_ADDIN_API_BASE_URL=http://127.0.0.1:8000
```

API Key 只配置在后端运行环境中。不要把真实 Key 写入 `manifest.xml`、前端源码、Webpack 配置、构建产物或文档。

不配置 `AI_PROFILES_JSON` 时，后端会用上面的旧变量生成 `Default AI (.env)`，插件里可直接选择。需要在插件中快速切换多个 OpenAI 兼容供应商时，可额外配置：

```text
OPENROUTER_API_KEY=...
LOCAL_OMLX_API_KEY=local-omlx-dev-key
MIMO_API_KEY=...
AI_PROFILES_JSON=[{"id":"openrouter-qwen","label":"OpenRouter / Qwen","api_base_url":"https://openrouter.ai/api/v1","api_key_env":"OPENROUTER_API_KEY","model":"qwen/xxx","default_api":"chat","supported_apis":["chat"]},{"id":"local-omlx","label":"本地 oMLX","api_base_url":"http://127.0.0.1:8001/v1","api_key_env":"LOCAL_OMLX_API_KEY","model":"Qwen3.6-35B-A3B-4.4bit-msq","default_api":"responses","supported_apis":["responses","chat"]},{"id":"xiaomi-mimo","label":"Xiaomi MiMo","api_base_url":"https://api.xiaomimimo.com/v1","api_key_env":"MIMO_API_KEY","model":"mimo-v2.5-pro","default_api":"chat","supported_apis":["chat"]}]
```

后端只把 profile 的 `id`、名称、模型和支持的 API 形态返回给插件，不返回 API Key。Xiaomi MiMo 当前按官方 OpenAI-compatible Chat Completions 接入，profile 只声明 `supported_apis=["chat"]`，不启用 Responses。

默认 SQLite 文件位置：

- Agent trace：`backend/var/agent-traces/traces.sqlite3`，由 `AGENT_TRACE_DIR` 控制。
- V2 工作台：`backend/var/agent-workspace/projects.sqlite3`，由 `AGENT_WORKSPACE_DIR` 控制。
- DOCX 下载索引：`backend/var/docx-results/results.sqlite3`，由 `DOCX_OUTPUT_DIR` 控制，记录下载恢复所需元数据和新任务的 `run_id`。

这两个目录都在 `backend/var/` 下，默认不提交到 Git。

## 启动 oMLX 本地 AI 服务

本地真实 AI 联调可用 oMLX 启动 OpenAI 兼容服务。脚本默认读取 `/Users/wulala/AI/models`，监听 `8001` 端口。

```bash
./scripts/start-omlx.sh
```

模型服务检查：

```bash
curl --noproxy 127.0.0.1 http://127.0.0.1:8001/v1/models \
  -H 'Authorization: Bearer local-omlx-dev-key'
```

如 `/v1/models` 返回的模型 ID 和 `.env` 不一致，请更新 `OPENAI_MODEL`。

## 启动后端

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000 --reload
```

健康检查：

```bash
curl --noproxy 127.0.0.1 http://127.0.0.1:8000/health
```

预期返回：

```json
{"status":"ok"}
```

## 启动 Word 插件

Windows 编辑电脑从零安装和本地试点部署见 [DEPLOYMENT.md](DEPLOYMENT.md)。

```bash
cd word-addin
npm install
npm run dev-server
```

插件开发服务默认运行在：

```text
https://localhost:3000/taskpane.html
```

旁加载到 Word：

```bash
cd word-addin
npm run start
```

## 联调流程

1. 启动 oMLX 或配置远程 OpenAI 兼容 API。
2. 启动后端，确认 `/health` 返回 `{"status":"ok"}`。
3. 启动 `word-addin` dev server。
4. 运行 `npm run start` 旁加载插件到 Word。
5. 当前选区审校：在 Word 文档中选中正文；全书审校：准备一个 `.docx` 文件。
6. 打开任务窗格，填写书名，按需选择审校范围、审校模式、temperature、API 模式和应用方式；全书模式需选择 `.docx` 文件。
7. 点击“AI 审校”。当前选区会先展示问题；全书 `.docx` 会展示分块进度并在完成后显示新文件名。
8. 当前选区可筛选、勾选、定位并点击“应用 N 条到 Word”；全书 `.docx` 点击“下载审校后 Word”获取后端生成的新文件，默认至少 7 天内可从历史记录再次下载。
9. 验证停止审校、重试当前分块、重试失败分块、历史导出和历史导入等常用流程。

## 测试与验证

后端测试：

```bash
cd backend
source .venv/bin/activate
python -m pytest -q
```

前端检查：

```bash
cd word-addin
npm run lint
npm run build
```

Manifest 联网校验：

```bash
cd word-addin
npm run validate
```

`npm run validate` 需要访问 Microsoft Office manifest validation service，离线或网络受限时可能失败。

## 文档维护约定

- 改接口契约时，更新 [spec.md](spec.md)。
- 改启动方式、端口、环境变量或联调流程时，更新本文件。
- 改 V2 Agent 工作台目标、架构边界或协作约束时，更新 [AGENTS.md](AGENTS.md) 和本文件中的 V2 路线图。
- 改 Windows 试点部署流程时，更新 [DEPLOYMENT.md](DEPLOYMENT.md)。
- 临时排障记录不要写进长期文档；需要留存时放到 Git 忽略的临时目录。
