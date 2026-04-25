# Word AI 审校助手 MVP 开发计划与技术方案

## 目标

面向出版社责任编辑，完成第一版可用闭环：在 Word 中选中一段文字，点击“AI 审校”，插件调用 FastAPI 后端，后端返回结构化审校问题和后端计算出的原文位置，插件可把审校建议批注到对应原文片段，也可在修订模式下把可直接替换的建议写成 Word 原生修订。

## MVP 范围

- 包含：Word 选区读取、provider 原生 AI session、后端审校 API、阶段进度流、结构化问题返回、后端原文定位、逐条精准批注、修订模式替换、快速/深度审校、Responses/Chat API 切换、停止审校、本地历史记录清空/导出/导入、基础错误提示。
- 不包含：登录、云端审校历史、全文扫描、token 级模型文本流、无人工确认地默认改正文。
- 第一版后端使用 OpenAI 兼容接口；没有 `AI_API_KEY` 时返回 mock 结果，保证本地可联调。
- 本地真实 AI 联调可使用 oMLX 启动 OpenAI 兼容服务，默认地址为 `http://127.0.0.1:8001/v1`。

## 技术架构

```text
Word 选区
  -> word-addin 读取选区文本
  -> POST /api/sessions 创建 AI session
  -> Responses: POST /api/proofread/stream；Chat: POST /api/proofread
  -> backend/FastAPI 调用 provider /v1/responses、/v1/chat/completions 或 mock service
  -> AI 返回精简 issues[]，包含 original/replacement/suggestion，不返回 start/end/comment
  -> backend 在选区文本中搜索 issue.original 并填充 start/end
  -> word-addin 按应用方式插入批注或生成 Word 修订，定位失败时回退汇总批注
```

### 前端

- 目录：`word-addin/`
- 技术栈：Office.js、TypeScript、Webpack。
- 本地地址：`https://localhost:3000/taskpane.html`。
- 后端地址：开发环境先请求同源 `/api/sessions` 创建 AI 对话；Responses 模式优先请求同源 `/api/proofread/stream` 获取阶段进度，流式不可用时回退 `/api/proofread`；Chat 模式直接请求 `/api/proofread`。插件请求会携带 `proofread_mode` 和 `provider_api`。这些接口均由 Webpack dev server 代理到 `http://127.0.0.1:8000`。

### 后端

- 目录：`backend/`
- 技术栈：Python、FastAPI、Pydantic、pydantic-settings、httpx、pytest。
- 本地地址：`http://127.0.0.1:8000`。
- 配置来源：后端运行环境变量；本地可通过 `uvicorn --env-file ../.env` 加载。
- 本地真实 AI：oMLX 服务建议监听 `http://127.0.0.1:8001/v1`。后端支持 OpenAI 兼容 Responses API 和 Chat Completions；Responses 模式使用 `previous_response_id` 做 provider 原生 session 续接，Chat 模式第一版为单轮审校。

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
  "provider_api": "responses",
  "proofread_mode": "fast",
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
      "replacement": "可直接替换原文的新文本",
      "suggestion": "修改建议说明",
      "start": 0,
      "end": 4
    }
  ]
}
```

字段约定：

- `category`：问题类别，第一版允许自由字符串，例如 `typo`、`grammar`、`style`、`fact`。
- `severity`：严重程度，使用 `low`、`medium`、`high`。
- `provider_api`：可选，支持 `responses`、`chat`；未传时使用后端环境变量 `AI_PROVIDER_API`，默认 `responses`。
- `proofread_mode`：可选，支持 `fast`、`thinking`；默认 `fast`。`fast` 只抓明显问题、优先响应速度；`thinking` 更细审、使用更高输出上限，优先审校质量。两种模式均不限制返回条数。
- AI 原始输出不包含 `start`、`end`、`comment`，只包含 `id`、`category`、`severity`、`original`、`replacement`、`suggestion`。
- `replacement`：可选，表示可直接替换 `original` 的正文文本；事实待核、需人工判断、体例疑问等不能直接替换的问题返回 `null`。空字符串会被后端归一为 `null`。
- `start`、`end`：后端按 `original` 在请求文本中计算，`start` 为包含式起点，`end` 为不包含式终点，均相对请求文本。重复 `original` 按 issue 顺序匹配下一处；找不到时返回 `null`。
- `issues` 为空表示未发现明显问题。
- `issues` 为空时，Word 插件只在任务窗格显示结果，不插入批注。

### `POST /api/proofread/stream`

Request 与 `/api/proofread` 相同，但仅用于 Responses 模式。Chat 模式直接使用 `/api/proofread`，后端按标准 Chat Completions request/response 调用 `/v1/chat/completions`。

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

- `session_id`：Word 插件当前 AI 对话 ID；Responses 模式后续审校请求携带它才能续接 provider 原生 session。
- 后端只保存 `session_id -> last_response_id`；Responses 模式真实上下文续接通过 `/v1/responses` 的 `previous_response_id` 完成。Chat 模式第一版不维护 provider 上下文。

## 开发任务

1. 搭建 FastAPI 后端骨架，提供 `GET /health` 和 `POST /api/proofread`。
2. 定义 Pydantic schema，校验空文本，固定响应结构。
3. 实现 mock 审校服务，未配置 `AI_API_KEY` 时返回可预测的本地结果。
4. 实现 OpenAI 兼容 Responses API 与 Chat Completions client，配置 `AI_API_KEY` 后请求对应 provider API 并解析精简 JSON。
5. 改造 Word 插件任务窗格，只保留“AI 审校”正式入口、状态提示和结果展示。
6. 插件内部读取 Word 当前选区：Responses 模式优先调用流式接口展示阶段进度，失败时回退普通接口；Chat 模式直接调用普通接口。
7. 后端返回问题时，插件按应用方式处理：批注模式按 `start/end` 和 `original` 精准插入逐条批注；修订模式临时开启 Word 修订跟踪，将可定位且有 `replacement` 的问题替换为 Word 原生修订；定位失败或无 `replacement` 的问题汇总插入当前选区 fallback 批注；未发现问题时只更新任务窗格，不插入批注。
8. 插件支持新建对话、停止审校、快速/深度审校、Responses/Chat API 切换、本地历史记录清空/导出/导入。
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
AI_MAX_TOKENS=32768
AI_FAST_MAX_TOKENS=16384
AI_THINKING_MAX_TOKENS=32768
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
```

`AI_API_KEY` 只允许存在于后端运行环境中，不进入 Word 插件代码、manifest、Webpack 构建变量或前端产物。未配置 `AI_API_KEY` 时，后端返回 mock 审校结果；已配置时按请求或环境配置调用 OpenAI 兼容 Responses API 或 Chat Completions。`local-omlx-dev-key` 只用于本机 oMLX 开发服务鉴权，不是真实云端密钥。AI HTTP 错误、非 JSON 返回、schema 不匹配统一转换为后端 502。`AI_FAST_MAX_TOKENS` 用于快速审校，`AI_THINKING_MAX_TOKENS` 用于深度审校。
`BACKEND_LOG_LEVEL` 默认 `INFO`，用于打印请求模式、文本长度、provider 状态、问题数和定位数量；不得打印 API Key 或选中文本全文。

本地 oMLX：

```bash
./scripts/start-omlx.sh
curl --noproxy 127.0.0.1 http://127.0.0.1:8001/v1/models \
  -H 'Authorization: Bearer local-omlx-dev-key'
```

`scripts/start-omlx.sh` 默认将 `Qwen3.6-35B-A3B-4.4bit-msq` 写入 `~/.omlx/model_settings.json`，设置为 default + pinned，使 oMLX 启动时预加载该模型。可通过 `OMLX_PRELOAD_MODEL` 覆盖模型 ID，或设置 `OMLX_CONFIGURE_MODEL_SETTINGS=0` 跳过该配置步骤。

Responses 流式接口 smoke test：

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
6. 如果存在审校问题，确认可定位问题批注在对应原文片段上；定位失败问题会作为一条 fallback 汇总批注插在当前选区；如果没有问题，确认不会插入批注。
7. 点击“新建对话”，确认后续审校使用新的 AI session。
8. 审校运行中点击“停止审校”，确认请求停止且不会插入批注。
9. 切换“快速审校/深度审校”和“Responses/Chat”，确认后续请求使用对应模式。
10. 切换“批注模式/修订模式”，确认批注模式不改正文，修订模式生成可接受/拒绝的 Word 修订。
11. 使用历史记录“清空、另存为、导入”，确认本地历史可管理。

## 验收标准

- `GET /health` 返回 200 和 `{ "status": "ok" }`。
- 空文本请求 `POST /api/proofread` 返回 422。
- 未配置 `AI_API_KEY` 时，后端返回 mock `issues[]`。
- 配置 `AI_API_KEY` 时，后端按 `provider_api` 调用真实 Responses API 或 Chat Completions；AI provider 异常时返回 502，且错误信息不包含 Key 或 Authorization header。
- `POST /api/sessions` 返回唯一 `session_id`。
- 同一 `session_id` 的第二次审校会带上 provider 上一次 `response.id` 作为 `previous_response_id`。
- Responses 模式下 provider 不支持原生 Responses session 时，后端明确返回错误，不降级；Chat 模式第一版不维护 provider session。
- AI 输出不包含 `start/end` 时，后端能按 `original` 计算 `start/end`。
- AI 输出含 `replacement` 时，后端响应保留该字段；AI 输出无 `replacement` 或空字符串时，响应为 `replacement: null`。
- 重复 `original` 会按 issue 顺序定位不同 occurrence；找不到 `original` 时返回 `start/end: null`。
- `proofread_mode=fast` 与 `proofread_mode=thinking` 使用不同 prompt 和 token 上限。
- Responses 流式接口返回阶段进度事件和最终 `result` 事件；Chat 模式不走 SSE。
- `npm run lint` 通过。
- `npm run build` 通过。
- Word 中空选区点击“AI 审校”时显示错误，不调用后端。
- 后端不可用时显示错误，不插入空批注。
- 后端返回非空 `issues[]` 时，批注模式对可定位问题逐条插入原文片段批注，对不可定位问题插入 fallback 汇总批注。
- 修订模式下，插件临时将 `document.changeTrackingMode` 设为 `TrackAll`，对可定位且有 `replacement` 的问题替换正文，完成后恢复原修订设置；无 `replacement` 或定位失败的问题插入 fallback 汇总批注。
- 后端返回空 `issues[]` 时，插件显示未发现明显问题，且不插入批注。
- 审校运行中点击“停止审校”时，插件中断请求、恢复按钮、不插入批注。
- 插件本地保存最近 20 条审校历史，可回看结果，并支持清空、另存为 JSON、导入 JSON。
