# Word AI 审校助手

面向出版社责任编辑的 Word AI 审校助手。当前版本以 V2.2 出版审校 Agent 工作台为主流程：责任编辑可以围绕当前选区或一本 `.docx` 书稿建立审校项目，让 Agent 生成计划、分阶段运行、沉淀候选问题和证据，并在编辑批准后写回 Word。

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

## 当前能力（V2.2 工作台）

- 审校来源：支持当前选区和全书 DOCX。当前选区由插件读取并通过 Office.js 写回；全书 DOCX 上传到后端，后端处理目录可见文本、正文、表格和常见文本框文字。本版不支持 `.doc`，请先另存为 `.docx`。
- 项目化流程：插件用“开始审校/重新审校”主按钮创建项目、建立文档地图、生成审校计划、后台运行 Agent run，并刷新候选问题队列。
- AI API：支持 OpenAI 兼容 Responses API 和 Chat Completions。插件默认深度审校、Temperature `0.6`、修订+批注模式，并优先选择 `mimo-v2.5-pro` profile；后端 API schema 的默认 temperature 仍为 `0.2`，用于直接 API 调用。未配置 Key 时返回 mock 结果，方便本地联调。
- 书籍信息：插件要求填写书名，介绍可选；后端把书籍信息和审校目标作为 prompt 背景，但不把完整正文写入长期记忆。
- 本地高置信规则：非 AI 本地规则默认只生成可直接替换的机械体例候选，包括中文语境英文逗号、中文正文半角括号和连续同类句末标点；术语并用和重复数字一致性阶段保留可观测入口，但默认不生成低置信人工核查候选。
- 候选确认：Agent 只生成候选问题、风险说明、证据和建议；编辑可逐条批准/拒绝，也可批量处理所有待确认候选。
- Word 写回：当前选区项目写回已批准候选时由插件完成定位、批注或修订+批注，并标记为已写回；DOCX 项目只写回 `approved` 候选，由后端生成结果 DOCX，插件提供下载入口。
- 可观测性：V2 工作台保存项目、文档地图、审校计划、run event trace、候选问题、项目记忆和审校报告；trace 不记录完整正文、API Key、Authorization 或 Bearer token。
- 结果保留：DOCX 写回结果保存在后端 `AGENT_WORKSPACE_DIR` 的项目输出目录；旧 V1 DOCX 下载索引仍由 `DOCX_OUTPUT_DIR` 控制。下载入口在文件未过期且未被外部清理时可继续使用。

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

当前代码已升级到 V2.2 审校工作台闭环：插件用一个“开始审校/重新审校”主按钮完成创建项目、建立文档地图、后台运行分阶段 Agent 和刷新候选问题；默认使用深度审校、Temperature 0.6、修订+批注模式，并优先选择 `mimo-v2.5-pro` 配置。候选问题区合并定位、批准、拒绝、写回、DOCX 下载和报告摘要；文档地图、审校计划、trace 和本书规则默认收进高级信息。插件 UI 不再提供 V1 独立入口；当前选区写回和定位由 Office.js 完成，DOCX 写回由后端完成。V2.2 不迁移 V1/V2 旧历史、trace、任务状态或 DOCX result index。

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

这些目录都在 `backend/var/` 下，默认不提交到 Git。

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
6. 打开任务窗格，选择“当前选区”或“全书 DOCX”，填写书名和审校目标；全书模式需选择 `.docx` 文件。
7. 按需展开高级设置，确认 AI profile、API 模式、审校模式、temperature 和写回模式。
8. 点击“开始审校”。插件会创建项目、建立文档地图、启动后台 Agent run，并自动刷新候选问题。
9. 审校完成后，在候选问题区查看证据，逐条批准/拒绝或批量处理待确认候选。
10. 当前选区项目点击“写回已批准”后由插件写回 Word；DOCX 项目点击“写回已批准”后由后端生成结果文件，再点击“下载结果 DOCX”。
11. 验证刷新当前项目、继续等待长任务、刷新 trace、项目删除、DOCX 下载和报告摘要等常用流程。

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
