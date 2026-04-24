# Backend README

`backend/` 是 Word AI 审校助手的 FastAPI 服务。它接收 Word 插件传来的选区文本，返回结构化审校问题。未配置真实 AI Key 时，会返回 mock 审校结果，方便本地联调。

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
│       └── proofread.py
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
  - 提供 `GET /health` 和 `POST /api/proofread`。
  - 将 AI client 抛出的 `AIClientError` 转换为 HTTP 502。

- `app/schemas.py`
  - 定义 API 请求和响应模型。
  - `ProofreadRequest` 校验 `text` 不能为空。
  - `ProofreadIssue` 定义单条审校问题结构。
  - `ProofreadResponse` 固定返回 `{ "issues": [...] }`。

- `app/settings.py`
  - 统一读取后端运行环境变量。
  - 使用 `pydantic-settings` 管理配置。
  - 负责 AI Key、OpenAI 兼容接口地址、模型、超时、max tokens、CORS origins 等配置。
  - API Key 只允许存在于后端运行环境中，不进入前端代码或构建产物。

- `app/services/proofread.py`
  - 审校服务编排层。
  - 未配置 `AI_API_KEY` 时返回 mock issue。
  - 配置 `AI_API_KEY` 时调用真实 AI client。

- `app/services/ai_client.py`
  - OpenAI 兼容 Chat Completions client。
  - 负责构造审校 prompt、发送请求、解析模型返回 JSON。
  - 统一将 provider HTTP 错误、非 JSON 响应、schema 不匹配转换为 `AIClientError`。

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
      "suggestion": "修改建议",
      "comment": "给责任编辑看的批注内容",
      "start": 0,
      "end": 4
    }
  ]
}
```

## 环境变量

本地开发推荐在仓库根目录创建 `.env`，并通过 `uvicorn --env-file ../.env` 加载。

```text
AI_API_KEY=
OPENAI_API_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-4o-mini
AI_REQUEST_TIMEOUT_SECONDS=30
AI_MAX_TOKENS=1200
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
```

说明：

- `AI_API_KEY` 为空时走 mock fallback。
- `AI_API_KEY` 有值时走真实 OpenAI 兼容接口。
- `BACKEND_CORS_ORIGINS` 使用英文逗号分隔。
- 真实 API Key 不要提交到 Git，不要写入 README、manifest、前端源码或构建产物。

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
- 真实 AI 分支请求 payload 正确。
- AI provider 异常不会泄露敏感信息，并统一返回 502。
