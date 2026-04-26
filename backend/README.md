# Backend README

`backend/` 是 Word AI 审校助手的 FastAPI 服务。它接收 Word 插件传来的选区或全书正文文本，返回结构化审校问题，并按 `original` 计算每条问题在单段或分块文本中的位置。分块任务会额外返回 `global_start/global_end`，供 Word 插件在长选区和全书正文中定位批注或修订。issue 中的 `replacement` 表示可直接替换原文的新文本，供 Word 插件修订模式使用。未配置真实 AI Key 时，会返回 mock 审校结果，方便本地联调。

## 目录结构

```text
backend/
├── app/
│   ├── __init__.py
│   ├── main.py
│   ├── schemas.py
│   ├── settings.py
│   └── services/
│       ├── __init__.py
│       ├── ai_client.py
│       ├── chunking.py
│       ├── proofread.py
│       ├── tasks.py
│       └── sessions.py
├── tests/
│   ├── test_ai_client.py
│   ├── test_api.py
│   └── test_settings.py
└── requirements.txt
```

## 文件说明

- `app/main.py`
  - FastAPI 应用入口。
  - 注册 CORS 中间件。
  - 提供 `GET /health`、`POST /api/sessions`、`POST /api/proofread`、`POST /api/proofread/stream`、`POST /api/proofread/chunked` 和分块任务接口。
  - 将 AI client 抛出的 `AIClientError` 转换为 HTTP 502。
  - 流式接口仅用于 Responses 模式，使用 SSE 返回阶段进度、最终结果或错误事件；Chat 模式使用普通接口。

- `app/schemas.py`
  - 定义 API 请求和响应模型。
  - `ProofreadRequest` 校验 `text` 不能为空。
  - `ProofreadRequest` 必须携带 `book.title`；`book.introduction` 可选。书籍信息用于 prompt 背景，不属于待审正文。
  - `ProofreadRequest` 可携带 `session_id`、`provider_api`、`proofread_mode`；`session_id` 仅用于兼容插件请求，不用于续接 AI 上下文。
  - `ProofreadIssue` 定义单条审校问题结构。
  - `ProofreadResponse` 固定返回 `{ "issues": [...] }`。
  - `ChunkedProofreadRequest`、`ChunkedProofreadIssue`、`ChunkedProofreadResult` 定义 V2 分块审校和异步任务结构。
  - `SessionResponse` 定义本地 session ID 和创建时间。

- `app/settings.py`
  - 统一读取后端运行环境变量。
  - 使用 `pydantic-settings` 管理配置。
  - 负责 AI Key、OpenAI 兼容接口地址、模型、超时、不同审校模式的 max tokens、CORS origins 等配置。
  - API Key 只允许存在于后端运行环境中，不进入前端代码或构建产物。

- `app/services/proofread.py`
  - 审校服务编排层。
  - 未配置 `AI_API_KEY` 时返回 mock issue。
  - 配置 `AI_API_KEY` 时按 `provider_api` 调用 Responses 或 Chat。
  - 对 mock/AI 返回的 issues 统一按 `original` 搜索并填充 `start/end`。
  - 返回给插件前过滤纯空白差异 issue：`original` 与 `replacement` 去掉所有空白后完全一致时不返回。
  - 每次审校都是独立请求，不读取或写入上一轮 provider response ID。

- `app/services/chunking.py`
  - V2 分块审校服务。
  - 当前选区超过 5000 字时按约 3000 字分块；全书正文始终按约 3000 字分块。
  - 优先在段落换行、句末标点附近切分；找不到边界时硬切。
  - 每个 chunk 复用 `proofread_text`，并把 chunk 内 `start/end` 转换为全文 `global_start/global_end`。

- `app/services/tasks.py`
  - V2 内存异步任务服务。
  - 使用模块级内存 dict 保存任务状态、进度、聚合结果和 SSE 事件。
  - 顺序处理 chunk；支持查询、SSE 订阅和取消。
  - 任务仅用于本地运行期，服务重启后不可恢复。

- `app/services/ai_client.py`
  - OpenAI 兼容 Responses API 与 Chat Completions client。
  - Responses 模式使用 `/v1/responses` 和 `text.format.type=json_object`，不发送 `previous_response_id`。
  - Chat 模式使用 `/v1/chat/completions`，并从 `choices[0].message.content` 读取 JSON。
  - 统一将 provider HTTP 错误、非 JSON 响应、schema 不匹配转换为 `AIClientError`。

- `app/services/sessions.py`
  - 创建轻量本地 session ID，用于兼容当前插件启动和“新建对话”流程。
  - 不保存 AI 上下文、不保存 provider `response.id`。

- `requirements.txt`
  - 后端 Python 依赖清单。
  - 包含 FastAPI、httpx、pydantic-settings、pytest、uvicorn。

- `tests/test_api.py`
  - API 层测试。
  - 覆盖健康检查、空文本校验、mock fallback、真实 AI 分支调用、AI 错误转 502、分块接口和异步任务。

- `tests/test_chunking.py`
  - 分块服务测试。
  - 覆盖短文本单段、长选区分块、全书分块、边界切分、硬切和全局位置转换。

- `tests/test_ai_client.py`
  - AI client 单元测试。
  - 使用 fake `httpx.AsyncClient` 验证请求 payload、timeout、max tokens、错误处理。

- `tests/test_settings.py`
  - 配置解析测试。
  - 覆盖默认配置、CORS origins 解析、环境变量读取。

## API

### `GET /health`

用于本地联调和部署健康检查。

```json
{"status":"ok"}
```

### `POST /api/proofread`

请求：

```json
{
  "text": "需要审校的 Word 选区文本",
  "book": {
    "title": "书名",
    "introduction": "可选书籍介绍"
  },
  "session_id": "session_xxx",
  "provider_api": "responses",
  "proofread_mode": "fast",
  "context": {
    "source": "word-addin"
  }
}
```

响应：

```json
{
  "issues": [
    {
      "id": "issue-1",
      "category": "style",
      "severity": "medium",
      "original": "原文片段",
      "replacement": "可直接替换原文的新文本",
      "suggestion": "修改建议说明",
      "start": 0,
      "end": 4
    }
  ]
}
```

`book.title` 必填，去掉首尾空白后不能为空；`book.introduction` 可选，空白会归一为 `null`。后端会把书籍信息加入 prompt 作为背景，但仍要求 AI 只审校请求里的 Word 选区文本。AI 原始输出不包含 `start/end/comment`；后端在返回给插件前计算 `start/end`。AI 原始输出可以包含 `replacement`，空字符串会归一为 `null`。如果 `original` 与 `replacement` 去掉所有空白后完全一致，说明只是加/删/改空白，后端会过滤该 issue，不返回给 Word 插件。`provider_api` 支持 `responses`、`chat`，`proofread_mode` 支持 `fast`、`thinking`。

### `POST /api/proofread/stream`

请求与 `/api/proofread` 相同，但仅支持 Responses 模式；Chat 模式请使用 `/api/proofread`。响应类型为 `text/event-stream`。正常情况下会返回阶段状态和最终结果：

```text
event: status
data: {"stage":"received","message":"已接收选区文本。"}

event: status
data: {"stage":"calling_ai","message":"正在调用 AI 审校。"}

event: status
data: {"stage":"calling_ai","message":"AI 审校仍在运行（约 1 秒）。"}

event: status
data: {"stage":"normalizing","message":"正在整理结构化审校结果。"}

event: result
data: {"issues":[]}

event: status
data: {"stage":"completed","message":"审校完成。"}
```

AI provider 异常时返回 `error` 事件，错误信息沿用非流式接口的脱敏错误文案。

### `POST /api/proofread/chunked`

同步分块审校接口。请求字段沿用 `/api/proofread`，增加：

```json
{
  "scope": "document",
  "chunk_size": 3000
}
```

响应为聚合后的分块结果：

```json
{
  "task_id": null,
  "scope": "document",
  "status": "succeeded",
  "total_chunks": 2,
  "completed_chunks": 2,
  "failed_chunks": 0,
  "issues": [
    {
      "id": "issue-1",
      "category": "style",
      "severity": "medium",
      "original": "原文片段",
      "replacement": null,
      "suggestion": "修改建议说明",
      "start": 0,
      "end": 4,
      "chunk_index": 0,
      "global_start": 0,
      "global_end": 4
    }
  ],
  "error_message": null
}
```

### 分块任务接口

- `POST /api/proofread/tasks`：创建内存异步任务，返回 `task_id` 和初始进度。
- `GET /api/proofread/tasks/{task_id}`：查询任务状态、进度和聚合结果。
- `GET /api/proofread/tasks/{task_id}/events`：订阅任务 SSE，事件包括 `queued`、`running`、`chunk_started`、`chunk_completed`、`chunk_failed`、`completed`、`cancelled`、`error`。
- `DELETE /api/proofread/tasks/{task_id}`：标记取消任务；当前 chunk 完成后停止后续 chunk。

任务只保存在后端内存中；服务重启或任务被清理后，查询会返回 404。

### `POST /api/sessions`

创建新的本地 session ID，用于兼容插件启动和“新建对话”流程。后端不保存 AI 上下文，也不会把上一轮 provider `response.id` 作为 `previous_response_id` 发给 `/v1/responses`；每次审校都是独立请求。

响应：

```json
{
  "session_id": "session_xxx",
  "created_at": "2026-04-25T04:00:00+00:00"
}
```

## 环境变量

本地开发推荐在仓库根目录创建 `.env`，并通过 `uvicorn --env-file ../.env` 加载。

```text
AI_API_KEY=local-omlx-dev-key
AI_PROVIDER_API=responses
OPENAI_API_BASE_URL=http://127.0.0.1:8001/v1
OPENAI_MODEL=Qwen3.6-35B-A3B-4.4bit-msq
AI_REQUEST_TIMEOUT_SECONDS=180
AI_MAX_TOKENS=32768
AI_FAST_MAX_TOKENS=16384
AI_THINKING_MAX_TOKENS=32768
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
```

说明：

- `AI_API_KEY` 为空时走 mock fallback。
- `AI_API_KEY` 有值时走真实 OpenAI 兼容 Responses API 或 Chat Completions。
- `AI_PROVIDER_API=responses` 表示默认使用 `/v1/responses`；也可在请求中传 `provider_api=chat` 临时切到 `/chat/completions`。
- `AI_FAST_MAX_TOKENS` 和 `AI_THINKING_MAX_TOKENS` 分别控制快速审校与深度审校的输出上限。
- `BACKEND_LOG_LEVEL` 控制后端日志级别，默认 `INFO`；本地调试报文时可临时设为 `DEBUG`。
- 本地真实 AI 联调推荐先运行仓库根目录的 `./scripts/start-omlx.sh`，默认 oMLX 地址为 `http://127.0.0.1:8001/v1`，并会将 `Qwen3.6-35B-A3B-4.4bit-msq` 配置为 default + pinned 以便启动时预加载。
- `local-omlx-dev-key` 只用于本机 oMLX 开发服务鉴权，不是真实云端密钥。
- `BACKEND_CORS_ORIGINS` 使用英文逗号分隔。
- 真实 API Key 不要提交到 Git，不要写入 README、manifest、前端源码或构建产物。

`INFO` 日志会打印请求入口、provider、审校模式、文本长度、AI HTTP 状态、问题数和定位数量，不打印选中文本全文。`DEBUG` 日志会打印后端请求/返回体、AI provider 请求 payload 和返回内容，可能包含选区文本、书名、介绍和 AI 输出；所有模式都不会打印 API Key、Authorization header 或 Bearer token。DEBUG 仅建议本地调试使用，不建议生产开启。

## 本地运行

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000 --reload
```

## 测试

```bash
cd backend
source .venv/bin/activate
python -m pytest -q
```

当前测试目标：

- API 行为稳定。
- mock fallback 可用。
- 后端能按 `original` 计算 `start/end`，重复片段按顺序定位，找不到时返回 `null`。
- 后端能保留有效 `replacement`，并将缺失或空字符串 `replacement` 归一为 `null`。
- 后端能过滤纯空白差异 issue。
- 真实 AI 分支 Responses API 与 Chat Completions 请求 payload 正确。
- Responses 请求不会携带 `previous_response_id`，同一 session 的多次审校也互不续接上下文。
- AI provider 异常不会泄露敏感信息，并统一返回 502。
- 流式审校接口会返回阶段事件、最终 `result` 事件或 `error` 事件。
