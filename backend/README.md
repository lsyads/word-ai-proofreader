# Backend README

`backend/` 是 Word AI 审校助手的 FastAPI 服务。它接收 Word 插件传来的当前选区文本，或接收全书 `.docx` 文件；调用 OpenAI 兼容 AI provider 或 mock fallback，返回结构化审校问题，并负责文本定位、DOCX 批注/修订写回和结果文件保存。

完整 API 契约见仓库根目录 [spec.md](../spec.md)；通讯链路见 [docs/architecture.md](../docs/architecture.md)。

## 目录结构

```text
backend/
├── app/
│   ├── agents/          # LangGraph V1 编排、LangChain tools、run trace
│   ├── main.py          # FastAPI 入口、路由、SSE 响应
│   ├── schemas.py       # Pydantic 请求/响应模型
│   ├── settings.py      # 环境变量配置
│   └── services/
│       ├── ai_client.py # Responses / Chat Completions client
│       ├── ai_provider_service.py # Agent 调用 AI provider 的薄服务封装
│       ├── chunk_service.py # Agent 调用分块/全局位置转换的薄服务封装
│       ├── chunking.py  # 分块、聚合、全局位置转换
│       ├── docx.py      # DOCX 抽取、分块、OOXML 批注/修订写回
│       ├── docx_service.py # Agent 调用 DOCX 解析/写回的薄服务封装
│       ├── docx_store.py # DOCX 结果文件稳定存储和过期清理
│       ├── docx_tasks.py # DOCX 全书审校任务和下载文件管理
│       ├── locator_service.py # Agent 调用定位能力的薄服务封装
│       ├── proofread.py # 审校编排、mock fallback、定位
│       ├── sessions.py  # 本地 session ID
│       └── tasks.py     # 内存异步分块任务
├── tests/
└── requirements.txt
```

## 主要职责

- 提供 `GET /health`、审校接口、本地 session 接口、同步分块调试接口、异步分块任务接口和 DOCX 全书审校任务接口。
- 校验请求：`text` 非空，`book.title` 必填且非空。
- 在未配置 `AI_API_KEY` 时返回 mock issue，保证本地可联调。
- 在配置 `AI_API_KEY` 时按 `provider_api` 调用 OpenAI 兼容 Responses API 或 Chat Completions，并透传请求级 `temperature`，默认 `0.2`。
- 把 AI 精简输出转换为结构化 `issues[]`，过滤纯空白差异，计算 `start/end/locator`。
- 当前选区 `> 7000` 字时按默认 `chunk_size=5000` 分块；分块结果额外返回 `global_start/global_end`。
- 全书 `.docx` 任务仅支持 `.docx`，不支持旧二进制 `.doc`；后端抽取目录可见文本、正文、表格和常见文本框文字，先按章、节拆分，仍超过 7000 字时再按可提取的目录小标题辅助拆分，issue 绑定来源 chunk 并在写回前可按 chunk 二次定位，最后生成新的 `.docx` 结果文件。
- DOCX 结果文件保存到 `DOCX_OUTPUT_DIR`，SQLite 索引记录输出文件、任务统计和过期时间；`DOCX_RETENTION_DAYS` 默认并强制最少为 7 天，后端重启后未过期结果仍可下载。
- 通过 LangGraph `StateGraph` 编排 V1 审校流程，并用 LangChain tool/schema 薄封装既有解析、分块、AI 审校、定位和 DOCX 写回能力。
- 每次审校生成 `run_id`，Agent trace 写入 `AGENT_TRACE_DIR` 下的 SQLite，记录节点、chunk 状态、耗时、错误和重试次数；trace 不记录完整正文、API Key、Authorization 或 Bearer token。
- 通过内存任务提供分块进度、SSE、取消、当前分块重试和失败分块重试。

## 本地运行

从仓库根目录复制环境变量模板：

```bash
cp .env.example .env
```

启动后端：

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

## 常用环境变量

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
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
DOCX_OUTPUT_DIR=var/docx-results
DOCX_RETENTION_DAYS=7
AGENT_TRACE_DIR=var/agent-traces
```

说明：

- `AI_API_KEY` 为空时使用 mock fallback。
- 不配置 `AI_PROFILES_JSON` 时，后端根据旧变量生成 `Default AI (.env)`。
- `AI_PROFILES_JSON` 可选，用于配置多个 OpenAI 兼容 profile；插件通过 `/api/ai-profiles` 加载下拉选项，后端不会返回 API Key。
- Xiaomi MiMo profile 可配置为 `{"id":"xiaomi-mimo","label":"Xiaomi MiMo","api_base_url":"https://api.xiaomimimo.com/v1","api_key_env":"MIMO_API_KEY","model":"mimo-v2.5-pro","default_api":"chat","supported_apis":["chat"]}`；后端会按 MiMo Chat Completions 字段适配，不启用 Responses。
- `AI_PROVIDER_API` 支持 `responses`、`chat`，作为默认 profile 的 `default_api`，也可由请求体临时覆盖。
- `temperature` 是请求级参数，范围 `0` 到 `1.5`，无需环境变量。
- `local-omlx-dev-key` 只用于本机 oMLX 开发服务鉴权，不是真实云端密钥。
- 真实 API Key 不要提交到 Git，不要写入前端源码或文档。

## LangGraph / LangChain Agent V1

后端保留现有 FastAPI 接口作为兼容外壳，内部新增 Agent 编排层：

- `app/agents/graph.py` 用 LangGraph `StateGraph` 定义 V1 单图工作流。
- `app/agents/nodes.py` 放图节点函数，负责准备输入、分块、逐 chunk 审校和 finalize。
- `app/agents/tools.py` 用 LangChain `StructuredTool` + Pydantic schema 薄封装既有能力，tool 名称为 snake_case。
- `app/agents/service.py` 提供 `agent_runner`，供普通审校、同步分块、异步任务和 DOCX 任务调用。
- `app/agents/trace.py` 负责 run/node/chunk trace 的 SQLite 记录和查询。

V1 不是让 LLM 自主选择工具，也不是多 Agent 框架。现阶段的重点是把原有审校链路变成显式、可观测、可测试的单图工作流；业务逻辑仍在 `services/` 中复用。

### 查看一次审校 trace

任意审校响应或 SSE 事件里拿到 `run_id` 后，可以查询：

```bash
curl --noproxy 127.0.0.1 \
  http://127.0.0.1:8000/api/agent/runs/agent_run_xxx/trace
```

返回内容包含：

- `nodes`：节点名、状态、开始/结束时间、耗时、错误信息。
- `chunks`：chunk 序号、范围、长度、状态、问题数、失败原因、重试次数。
- `metadata`：审校范围、模型 profile、API 模式、审校模式等脱敏信息。

trace 不保存完整正文、API Key、Authorization header 或 Bearer token。

### SQLite 文件位置

后端当前有两个 SQLite 文件，默认都位于 `backend/var/`，该目录不提交到 Git：

```text
backend/var/agent-traces/traces.sqlite3   # Agent run/node/chunk trace
backend/var/docx-results/results.sqlite3  # DOCX 结果文件索引
```

对应环境变量：

```text
AGENT_TRACE_DIR=var/agent-traces
DOCX_OUTPUT_DIR=var/docx-results
```

相对路径会按 `backend/` 目录解析。若要把 trace 或 DOCX 结果索引放到外部持久化目录，请配置绝对路径，例如：

```text
AGENT_TRACE_DIR=/data/word-ai-proofreader/agent-traces
DOCX_OUTPUT_DIR=/data/word-ai-proofreader/docx-results
```

本地可用 `sqlite3` 查看表结构：

```bash
sqlite3 backend/var/agent-traces/traces.sqlite3 ".tables"
sqlite3 backend/var/docx-results/results.sqlite3 ".tables"
```

## 测试

```bash
cd backend
source .venv/bin/activate
python -m pytest -q
```

测试覆盖重点：

- API 校验和错误映射。
- mock fallback。
- Responses / Chat 请求 payload。
- AI 输出清理、schema 校验和错误处理。
- `original` 定位、重复片段定位、`locator`、纯空白差异过滤。
- 分块规则、全局位置转换、异步任务和 SSE 事件。
- DOCX 文字抽取、章节分块、批注/修订写回、结果文件下载、稳定保留和过期清理。
- Agent graph 编译、LangChain tools schema、run trace 记录和脱敏。

## 开发注意事项

- API 契约变更必须先同步 [../spec.md](../spec.md)。
- 运行方式、端口或环境变量变更必须同步根目录 [../README.md](../README.md)。
- 临时排障记录放在 `docs/tmp/`，不要写进长期 README。
- 后端日志不得打印 API Key、Authorization header 或 Bearer token。
