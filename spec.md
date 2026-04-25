# Word AI 审校助手 MVP 开发计划与技术方案

## 目标

面向出版社责任编辑，完成第一版最小闭环：在 Word 中选中一段文字，点击“AI 审校”，插件调用 FastAPI 后端，后端返回结构化审校问题，插件把审校结果作为一条汇总批注插入当前选区。

## MVP 范围

- 包含：Word 选区读取、provider 原生 AI session、后端审校 API、阶段进度流、结构化问题返回、当前选区汇总批注、停止审校、本地历史记录、基础错误提示。
- 不包含：登录、云端审校历史、全文扫描、token 级模型文本流、多模型切换、逐条问题精准定位批注。
- 第一版后端使用 OpenAI 兼容接口；没有 `AI_API_KEY` 时返回 mock 结果，保证本地可联调。
- 本地真实 AI 联调可使用 oMLX 启动 OpenAI 兼容服务，默认地址为 `http://127.0.0.1:8001/v1`。

## 技术架构

```text
Word 选区
  -> word-addin 读取选区文本
  -> POST /api/sessions 创建 AI session
  -> POST /api/proofread/stream
  -> backend/FastAPI 调用 provider /v1/responses 或 mock service
  -> 同一 session 使用 previous_response_id 续接上下文
  -> 返回阶段进度和 issues[]
  -> word-addin 汇总 issues
  -> 有问题时 selection.insertComment(...)
```

### 前端

- 目录：`word-addin/`
- 技术栈：Office.js、TypeScript、Webpack。
- 本地地址：`https://localhost:3000/taskpane.html`。
- 后端地址：开发环境先请求同源 `/api/sessions` 创建 AI 对话，再优先请求同源 `/api/proofread/stream` 获取阶段进度，并在任务窗格“运行过程”区域逐条展示；流式不可用时回退 `/api/proofread`。这些接口均由 Webpack dev server 代理到 `http://127.0.0.1:8000`。

### 后端

- 目录：`backend/`
- 技术栈：Python、FastAPI、Pydantic、pydantic-settings、httpx、pytest。
- 本地地址：`http://127.0.0.1:8000`。
- 配置来源：后端运行环境变量；本地可通过 `uvicorn --env-file ../.env` 加载。
- 本地真实 AI：oMLX 服务建议监听 `http://127.0.0.1:8001/v1`，后端通过 OpenAI 兼容 Responses API 调用，并使用 `previous_response_id` 做 provider 原生 session 续接。

## API 契约

### `GET /health`

Response:

```json
{
  "status": "ok"
}
```

### `POST /api/proofread`

Request:

```json
{
  "text": "需要审校的 Word 选区文本",
  "session_id": "session_xxx",
  "context": {
    "source": "word-addin"
  }
}
```

Response:

```json
{
  "issues": [
    {
      "id": "issue-1",
      "category": "typo",
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

字段约定：

- `category`：问题类别，第一版允许自由字符串，例如 `typo`、`grammar`、`style`、`fact`。
- `severity`：严重程度，使用 `low`、`medium`、`high`。
- `start`、`end`：相对请求文本的字符偏移；MVP 暂不使用它们做精准定位。
- `issues` 为空表示未发现明显问题。
- `issues` 为空时，Word 插件只在任务窗格显示结果，不插入批注。

### `POST /api/proofread/stream`

Request 与 `/api/proofread` 相同。

Response 使用 `text/event-stream`：

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

AI provider 异常时返回 `error` 事件：

```text
event: error
data: {"message":"AI provider returned HTTP 500"}
```

### `POST /api/sessions`

Response:

```json
{
  "session_id": "session_xxx",
  "created_at": "2026-04-25T04:00:00+00:00"
}
```

字段约定：

- `session_id`：Word 插件当前 AI 对话 ID；后续审校请求必须携带它才能续接 provider 原生 session。
- 后端只保存 `session_id -> last_response_id`；真实上下文续接通过 `/v1/responses` 的 `previous_response_id` 完成。

## 开发任务

1. 搭建 FastAPI 后端骨架，提供 `GET /health` 和 `POST /api/proofread`。
2. 定义 Pydantic schema，校验空文本，固定响应结构。
3. 实现 mock 审校服务，未配置 `AI_API_KEY` 时返回可预测的本地结果。
4. 实现 OpenAI 兼容 Responses API client，配置 `AI_API_KEY` 后请求 `/v1/responses` 并解析 JSON。
5. 改造 Word 插件任务窗格，只保留“AI 审校”正式入口、状态提示和结果展示。
6. 插件内部读取 Word 当前选区，优先调用流式接口展示阶段进度，失败时回退普通接口。
7. 后端返回问题时插入一条汇总批注；未发现问题时只更新任务窗格，不插入批注。
8. 插件支持新建对话、停止审校、本地历史记录。
9. 补充后端测试、插件 lint/build 验证和本地联调说明。

## 本地运行

后端：

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000 --reload
```

AI 配置：

```text
AI_API_KEY=local-omlx-dev-key
AI_PROVIDER_API=responses
AI_REQUIRE_NATIVE_SESSION=true
OPENAI_API_BASE_URL=http://127.0.0.1:8001/v1
OPENAI_MODEL=Qwen3.6-35B-A3B-4.4bit-msq
AI_REQUEST_TIMEOUT_SECONDS=180
AI_MAX_TOKENS=1200
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
```

`AI_API_KEY` 只允许存在于后端运行环境中，不进入 Word 插件代码、manifest、Webpack 构建变量或前端产物。未配置 `AI_API_KEY` 时，后端返回 mock 审校结果；已配置时调用 OpenAI 兼容 Responses API。`local-omlx-dev-key` 只用于本机 oMLX 开发服务鉴权，不是真实云端密钥。AI HTTP 错误、非 JSON 返回、schema 不匹配统一转换为后端 502。`AI_REQUIRE_NATIVE_SESSION=true` 时，如果 provider 不支持 `/v1/responses` 原生 session 能力，后端明确报错，不静默降级。

本地 oMLX：

```bash
./scripts/start-omlx.sh
curl --noproxy 127.0.0.1 http://127.0.0.1:8001/v1/models \
  -H 'Authorization: Bearer local-omlx-dev-key'
```

`scripts/start-omlx.sh` 默认将 `Qwen3.6-35B-A3B-4.4bit-msq` 写入 `~/.omlx/model_settings.json`，设置为 default + pinned，使 oMLX 启动时预加载该模型。可通过 `OMLX_PRELOAD_MODEL` 覆盖模型 ID，或设置 `OMLX_CONFIGURE_MODEL_SETTINGS=0` 跳过该配置步骤。

流式接口 smoke test：

```bash
curl --no-buffer --noproxy 127.0.0.1 -X POST http://127.0.0.1:8000/api/proofread/stream \
  -H 'Content-Type: application/json' \
  -H 'Accept: text/event-stream' \
  -d '{"text":"这是一段需要审校的文本。","session_id":"session_xxx","context":{"source":"manual-curl"}}'
```

插件：

```bash
cd word-addin
npm install
npm run dev-server
```

联调：

1. 在 Word 中旁加载 `word-addin/manifest.xml`。
2. 打开任务窗格。
3. 选中一段正文。
4. 点击“AI 审校”。
5. 确认任务窗格“运行过程”区域显示阶段进度，并在“审校结果”区域显示最终结果。
6. 如果存在审校问题，确认当前选区出现一条包含审校问题的 Word 批注；如果没有问题，确认不会插入批注。
7. 点击“新建对话”，确认后续审校使用新的 AI session。
8. 审校运行中点击“停止审校”，确认请求停止且不会插入批注。

## 验收标准

- `GET /health` 返回 200 和 `{ "status": "ok" }`。
- 空文本请求 `POST /api/proofread` 返回 422。
- 未配置 `AI_API_KEY` 时，后端返回 mock `issues[]`。
- 配置 `AI_API_KEY` 时，后端调用真实 Responses API；AI provider 异常时返回 502，且错误信息不包含 Key 或 Authorization header。
- `POST /api/sessions` 返回唯一 `session_id`。
- 同一 `session_id` 的第二次审校会带上 provider 上一次 `response.id` 作为 `previous_response_id`。
- provider 不支持原生 Responses session 时，后端明确返回错误，不降级。
- 流式接口返回阶段进度事件和最终 `result` 事件。
- `npm run lint` 通过。
- `npm run build` 通过。
- Word 中空选区点击“AI 审校”时显示错误，不调用后端。
- 后端不可用时显示错误，不插入空批注。
- 后端返回非空 `issues[]` 时，插件在当前选区插入一条汇总批注。
- 后端返回空 `issues[]` 时，插件显示未发现明显问题，且不插入批注。
- 审校运行中点击“停止审校”时，插件中断请求、恢复按钮、不插入批注。
- 插件本地保存最近 20 条审校历史，可回看结果。
