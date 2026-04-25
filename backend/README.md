# Backend README

`backend/` 是 Word AI 审校助手的 FastAPI 服务。它接收 Word 插件传来的选区文本，返回结构化审校问题，并按 `original` 计算每条问题在选区文本中的 `start/end`。issue 中的 `replacement` 表示可直接替换原文的新文本，供 Word 插件修订模式使用。未配置真实 AI Key 时，会返回 mock 审校结果，方便本地联调。

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
│       ├── proofread.py
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
  - 提供 `GET /health`、`POST /api/sessions`、`POST /api/proofread` 和 `POST /api/proofread/stream`。
  - 将 AI client 抛出的 `AIClientError` 转换为 HTTP 502。
  - 流式接口使用 SSE 返回阶段进度、最终结果或错误事件。

- `app/schemas.py`
  - 定义 API 请求和响应模型。
  - `ProofreadRequest` 校验 `text` 不能为空。
  - `ProofreadRequest` 可携带 `session_id` 续接 provider 原生 Responses session，并可携带 `provider_api`、`proofread_mode`。
  - `ProofreadIssue` 定义单条审校问题结构。
  - `ProofreadResponse` 固定返回 `{ "issues": [...] }`。
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
  - Responses 模式读取并更新当前 session 的 provider `response.id`。

- `app/services/ai_client.py`
  - OpenAI 兼容 Responses API 与 Chat Completions client。
  - Responses 模式使用 `/v1/responses`、`previous_response_id` 和 `text.format.type=json_object`。
  - Chat 模式使用 `/v1/chat/completions` 和 `response_format.type=json_object`，第一版不维护 provider session。
  - 统一将 provider HTTP 错误、非 JSON 响应、schema 不匹配转换为 `AIClientError`。

- `app/services/sessions.py`
  - 维护轻量内存 session 映射。
  - 保存 `session_id -> last_response_id`，真正上下文由 provider 原生 Responses session 续接。

- `requirements.txt`
  - 后端 Python 依赖清单。
  - 包含 FastAPI、httpx、pydantic-settings、pytest、uvicorn。

- `tests/test_api.py`
  - API 层测试。
  - 覆盖健康检查、空文本校验、mock fallback、真实 AI 分支调用、AI 错误转 502。

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

AI 原始输出不包含 `start/end/comment`；后端在返回给插件前计算 `start/end`。AI 原始输出可以包含 `replacement`，空字符串会归一为 `null`。`provider_api` 支持 `responses`、`chat`，`proofread_mode` 支持 `fast`、`thinking`。

### `POST /api/proofread/stream`

请求与 `/api/proofread` 相同，响应类型为 `text/event-stream`。正常情况下会返回阶段状态和最终结果：

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

### `POST /api/sessions`

创建新的 AI 对话 session。后端只保存本地 `session_id` 到 provider `response.id` 的映射；后续审校请求携带该 `session_id`，后端会把上一轮 provider `response.id` 作为 `previous_response_id` 发给 `/v1/responses`。

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
AI_REQUIRE_NATIVE_SESSION=true
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
- `BACKEND_LOG_LEVEL` 控制后端日志级别，默认 `INFO`；调试定位细节时可设为 `DEBUG`。
- `AI_REQUIRE_NATIVE_SESSION=true` 表示 Responses 模式下 provider 不支持原生 session 时明确报错，不静默降级。
- 本地真实 AI 联调推荐先运行仓库根目录的 `./scripts/start-omlx.sh`，默认 oMLX 地址为 `http://127.0.0.1:8001/v1`，并会将 `Qwen3.6-35B-A3B-4.4bit-msq` 配置为 default + pinned 以便启动时预加载。
- `local-omlx-dev-key` 只用于本机 oMLX 开发服务鉴权，不是真实云端密钥。
- `BACKEND_CORS_ORIGINS` 使用英文逗号分隔。
- 真实 API Key 不要提交到 Git，不要写入 README、manifest、前端源码或构建产物。

日志会打印请求入口、provider、审校模式、文本长度、AI HTTP 状态、问题数和定位数量。日志不会打印 API Key、Authorization header 或选中文本全文。

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
- 真实 AI 分支 Responses API 与 Chat Completions 请求 payload 正确。
- 同一 session 的第二次审校会带 `previous_response_id`。
- AI provider 异常不会泄露敏感信息，并统一返回 502。
- 流式审校接口会返回阶段事件、最终 `result` 事件或 `error` 事件。
