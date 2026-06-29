# Word AI 审校助手

面向出版社责任编辑的 Word AI 审校助手。当前版本以 V2.2 出版审校工作台为主流程：编辑在 Word 中选择文本或上传 `.docx` 书稿，点击开始审校，查看建议，接受或忽略后再写回 Word。

V2.2 围绕一本书建立审校项目，保存文档地图、审校计划、运行状态、候选建议、编辑决策和报告。V2 项目数据、历史、trace、任务快照和结果索引不迁移旧版本；代码仍保留底层直接 API，用于复用已验证的文档处理能力和开发调试。

## 项目结构

```text
.
├── backend/      # Python FastAPI 后端服务
├── scripts/      # 本地开发辅助脚本
├── word-addin/   # Office.js + TypeScript + Webpack Word 插件
├── AGENTS.md     # 协作约定
├── ARCHITECTURE.md # 代码入口、模块边界和 V2 数据流
├── DEPLOYMENT.md # Windows 本地试点部署
├── TESTING.md    # 测试命令和最小验证矩阵
└── spec.md       # API 契约和验收标准
```

## 当前能力（V2.2 工作台）

- 审校来源：支持当前选区和全书 DOCX。当前选区由插件读取并通过 Office.js 写回；全书 DOCX 上传到后端，后端处理目录可见文本、正文、表格和常见文本框文字。本版不支持 `.doc`，请先另存为 `.docx`。
- 普通使用流程：插件主界面保留“审校来源、书名、开始审校、本次进度、审校建议、写回/下载”，最近审校和排障信息默认折叠。
- AI API：支持 OpenAI 兼容 Responses API 和 Chat Completions。插件默认深度审校、Temperature `0.6`、修订+批注模式，并优先匹配 id、model 或 label 包含 `mimo-v2.5-pro` 的 profile；这些配置放在“更多设置（试点支持）”中。后端 API schema 的默认 temperature 仍为 `0.2`，用于直接 API 调用。未配置 Key 时返回 mock 结果，方便本地联调。
- 书籍信息：插件要求填写书名，介绍可选；后端把书籍信息和审校目标作为 prompt 背景，但不把完整正文写入长期记忆。
- 责任编辑边界：当前审校计划只展示 5 个实际执行阶段：生成计划、基础语言审校、候选归并、二次复核、等待编辑确认。术语、本书约定和一致性不再作为独立计划步骤展示；标点符号、空格、全半角和中英文符号转换等机械校对属于校对公司任务，AI 或本地规则默认都不会让这类建议进入责任编辑确认队列。
- 建议处理：Agent 只生成审校建议、风险说明和证据；编辑逐条接受/忽略，也可批量处理所有待处理建议。无论置信度多高，都必须接受后才会写回。
- Word 写回：当前选区写回已接受建议时由插件完成定位、批注或修订+批注；DOCX 审校只写回已接受建议，由后端生成审校后文件，插件提供下载入口。
- 可观测性：V2 工作台保存本次审校、文档地图、审校计划、run event trace、建议、项目记忆和审校报告；这些技术信息默认收进“排障信息（技术支持）”。trace 不记录完整正文、API Key、Authorization 或 Bearer token。
- 结果保留：V2 DOCX 写回结果保存在后端 `AGENT_WORKSPACE_DIR` 的项目输出目录，当前不返回保留期限或过期时间；下载入口在项目输出文件仍存在时可继续使用。底层 DOCX 任务结果索引仍由 `DOCX_OUTPUT_DIR` 和 `DOCX_RETENTION_DAYS` 控制。

文档导航：

- [AGENTS.md](AGENTS.md)：AI coding 协作约定、V2 目标和边界。
- [ARCHITECTURE.md](ARCHITECTURE.md)：代码入口、模块边界、V2 工作台数据流和常见改动入口。
- [spec.md](spec.md)：API 契约、状态语义、环境变量和验收标准。
- [TESTING.md](TESTING.md)：后端、前端、manifest 和 Word 手工联调的测试矩阵。
- [DEPLOYMENT.md](DEPLOYMENT.md)：Windows 编辑电脑本地试点部署手册。

## V2 工作台说明

V2.2 已收敛为项目化审校闭环：选择当前选区或全书 DOCX、填写书名、开始审校、查看建议、接受/忽略、写回或下载审校后文件。当前选区写回和定位由 Office.js 完成，DOCX 写回由后端完成；插件 UI 不提供旧版独立入口。

产品目标和 AI coding 边界见 [AGENTS.md](AGENTS.md)，代码入口和端到端流程见 [ARCHITECTURE.md](ARCHITECTURE.md)，完整 API 契约和状态语义见 [spec.md](spec.md)。

## 环境变量

复制模板后按需填写：

```bash
cp .env.example .env
```

后端当前会读取的常用配置：

```text
AI_API_KEY=local-omlx-dev-key
MIMO_API_KEY=
AI_PROVIDER_API=responses
OPENAI_API_BASE_URL=http://127.0.0.1:8001/v1
OPENAI_MODEL=Qwen3.6-35B-A3B-4.4bit-msq
AI_REQUEST_TIMEOUT_SECONDS=180
AI_REQUEST_TIMEOUT_MIN_SECONDS=60
AI_REQUEST_TIMEOUT_MAX_SECONDS=900
AI_REQUEST_TIMEOUT_BASE_SECONDS=60
AI_FAST_TIMEOUT_SECONDS_PER_1K_TOKENS=45
AI_THINKING_TIMEOUT_SECONDS_PER_1K_TOKENS=75
AI_FAST_MAX_TOKENS=8192
AI_THINKING_MAX_TOKENS=16384
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
DOCX_OUTPUT_DIR=var/docx-results
DOCX_RETENTION_DAYS=7
AGENT_TRACE_DIR=var/agent-traces
AGENT_WORKSPACE_DIR=var/agent-workspace
```

API Key 只配置在后端运行环境中。不要把真实 Key 写入 `manifest.xml`、前端源码、Webpack 配置、构建产物或文档。

`.env.example` 中还保留 `BACKEND_HOST`、`BACKEND_PORT`、`WORD_ADDIN_API_BASE_URL` 作为人工启动命令和历史兼容说明；当前代码不会读取这些变量。后端监听地址由 `uvicorn ... --host/--port` 决定，插件开发代理在 `word-addin/webpack.config.js` 中指向 `http://127.0.0.1:8000`。

`.env.example` 默认包含 `local-omlx`、`openrouter-qwen`、`xiaomi-mimo` 三个 profile 的 `AI_PROFILES_JSON` 模板。插件会优先匹配 id、model 或 label 包含 `mimo-v2.5-pro` 的 profile；当前模板中 `xiaomi-mimo` profile 的 model 是 `mimo-v2.5-pro`，Key 来自 `MIMO_API_KEY`。如果删除或留空 `AI_PROFILES_JSON`，后端会用上面的旧变量生成 `Default AI (.env)`，插件里可直接选择。详细环境变量规则见 [spec.md](spec.md)，Windows 试点配置见 [DEPLOYMENT.md](DEPLOYMENT.md)。

后端只把 profile 的 `id`、名称、模型和支持的 API 形态返回给插件，不返回 API Key。

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
6. 打开任务窗格，选择“当前选区”或“全书 DOCX”，填写书名；全书模式需选择 `.docx` 文件。
7. 按需展开“补充信息”或“更多设置（试点支持）”，调整审校重点、模型配置、审校模式、temperature 和写回模式。
8. 点击“开始审校”。插件会自动刷新本次进度，完成后显示审校建议。
9. 审校完成后，在审校建议区查看原文、建议改为、修改说明和依据，逐条接受/忽略或批量处理待处理建议。
10. 当前选区点击“写回已接受建议”后由插件写回 Word；DOCX 点击写回后由后端生成审校后文件，再点击“下载审校后文件”。
11. 验证刷新进度、继续等待长任务、手动打开最近审校、排障信息、DOCX 下载和报告摘要等常用流程。

## 测试与验证

常用检查如下，详细测试矩阵见 [TESTING.md](TESTING.md)。

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
- 改代码入口、模块边界或 V2 数据流时，更新 [ARCHITECTURE.md](ARCHITECTURE.md)。
- 改测试命令、验证矩阵或手工联调要求时，更新 [TESTING.md](TESTING.md)。
- 改启动方式、端口、环境变量或联调流程时，更新本文件。
- 改 V2 Agent 工作台目标、架构边界或协作约束时，更新 [AGENTS.md](AGENTS.md)。
- 改 Windows 试点部署流程时，更新 [DEPLOYMENT.md](DEPLOYMENT.md)。
- 临时排障记录不要写进长期文档；需要留存时放到 Git 忽略的临时目录。
