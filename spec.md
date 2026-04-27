# Word AI 审校助手 V2 开发计划与技术方案

## 目标

面向出版社责任编辑，在 Word 中完成统一范围审校闭环：小选区直接审校，长选区和全书正文自动分块审校；后端返回结构化审校问题和原文位置，插件可把建议批注到对应原文片段，也可在修订模式下把可直接替换的建议写成 Word 原生修订。

## V2 范围

- 包含：Word 选区读取、本地 session 流程、后端审校 API、阶段进度流、结构化问题返回、后端原文定位、纯空白差异过滤、结果筛选、逐条勾选、单条定位预览、逐条精准批注、修订模式替换、快速/深度审校、Responses/Chat API 切换、停止审校、本地历史记录清空/导出/导入、基础错误提示。
- 新增：当前选区超过 5000 字自动分块审校；全书正文按约 3000 字自动分块审校；分块任务使用后端内存异步任务、SSE 进度、heartbeat 和轮询兜底；全书结果支持二次确认后批注或修订。
- 不包含：登录、云端审校历史、跨服务重启恢复任务、页眉页脚/脚注/文本框扫描、token 级模型文本流、无人工确认地默认改正文。
- 第一版后端使用 OpenAI 兼容接口；没有 `AI_API_KEY` 时返回 mock 结果，保证本地可联调。
- 本地真实 AI 联调可使用 oMLX 启动 OpenAI 兼容服务，默认地址为 `http://127.0.0.1:8001/v1`。

## 技术架构

```text
Word 选区
  -> word-addin 读取选区文本
  -> POST /api/sessions 创建本地 session
  -> 小选区: Responses POST /api/proofread/stream；Chat POST /api/proofread
  -> 长选区/全书: POST /api/proofread/tasks + SSE/轮询任务进度
  -> backend/FastAPI 调用 provider /v1/responses、/v1/chat/completions 或 mock service
  -> AI 返回精简 issues[]，包含 original/replacement/suggestion，不返回 start/end/comment
  -> backend 在单段或分块文本中搜索 issue.original 并填充 start/end；分块结果额外填充 global_start/global_end
  -> word-addin 先展示结果，用户确认后按应用方式插入批注或生成 Word 修订
```

### 前端

- 目录：`word-addin/`
- 技术栈：Office.js、TypeScript、Webpack。
- 本地地址：`https://localhost:3000/taskpane.html`。
- 后端地址：开发环境先请求同源 `/api/sessions` 创建本地 session；短选区 Responses 模式优先请求同源 `/api/proofread/stream` 获取阶段进度，流式不可用时回退 `/api/proofread`；短选区 Chat 模式直接请求 `/api/proofread`。长选区和全书请求 `/api/proofread/tasks`，优先用 `/api/proofread/tasks/{task_id}/events` 获取 SSE 进度，失败时轮询 `/api/proofread/tasks/{task_id}`。这些接口均由 Webpack dev server 代理到 `http://127.0.0.1:8000`。

### 后端

- 目录：`backend/`
- 技术栈：Python、FastAPI、Pydantic、pydantic-settings、httpx、pytest。
- 本地地址：`http://127.0.0.1:8000`。
- 配置来源：后端运行环境变量；本地可通过 `uvicorn --env-file ../.env` 加载。
- 本地真实 AI：oMLX 服务建议监听 `http://127.0.0.1:8001/v1`。后端支持 OpenAI 兼容 Responses API 和 Chat Completions；两种模式都按单轮审校处理，不发送 `previous_response_id`。

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
  "book": {
    "title": "书名",
    "introduction": "可选书籍介绍"
  },
  "session_id": "session_xxx",
  "provider_api": "responses",
  "proofread_mode": "fast",
  "reasoning_enabled": false,
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
- `book`：必填。`title` 为书名，去掉首尾空白后不能为空；`introduction` 为可选书籍介绍，空白会归一为 `null`。缺少 `book` 或空书名返回 422。
- `provider_api`：可选，支持 `responses`、`chat`；未传时使用后端环境变量 `AI_PROVIDER_API`，默认 `responses`。
- `proofread_mode`：可选，支持 `fast`、`thinking`；默认 `fast`。`fast` 只抓明显问题、优先响应速度；`thinking` 更细审、使用更高输出上限，优先审校质量。两种模式均不限制返回条数。
- `reasoning_enabled`：可选布尔值，默认 `false`。启用时，Chat Completions 请求携带 `"reasoning": {"enabled": true}`；关闭时携带 `"reasoning": {"enabled": false}`。该开关独立于 `proofread_mode`，不改变 prompt 文案。
- 后端会把书名和书籍介绍作为 prompt 背景传给 AI；书籍信息不属于待审正文，AI 仍只能对 `<text>` 内的 Word 选区文本返回可定位 issue。
- AI 原始输出不包含 `start`、`end`、`comment`，只包含 `id`、`category`、`severity`、`original`、`replacement`、`suggestion`。
- `replacement`：可选，表示可直接替换 `original` 的正文文本；事实待核、需人工判断、体例疑问等不能直接替换的问题返回 `null`。空字符串会被后端归一为 `null`。
- 纯空白差异过滤：如果 `original` 与 `replacement` 去掉所有空白后完全一致，后端会过滤该 issue，不返回给 Word 插件。
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

### `POST /api/proofread/chunked`

同步分块审校接口，用于测试、调试和较小分块任务。请求字段沿用 `/api/proofread`，增加：

```json
{
  "scope": "selection",
  "chunk_size": 3000
}
```

`scope` 支持 `selection` 和 `document`。`selection` 文本不超过 5000 字时返回单个 chunk；超过 5000 字时按约 3000 字分块。`document` 始终按约 3000 字分块。分块优先在段落换行和句末标点附近切分，找不到边界时硬切。

Response:

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
      "category": "typo",
      "severity": "low",
      "original": "原文片段",
      "replacement": "替换文本",
      "suggestion": "修改建议",
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

### 异步分块任务接口

创建任务：

```http
POST /api/proofread/tasks
```

请求与 `/api/proofread/chunked` 相同。Response 为 `ChunkedProofreadResult`，其中 `task_id` 必填，初始 `status` 为 `queued`。

查询任务：

```http
GET /api/proofread/tasks/{task_id}
```

返回当前进度、聚合结果和错误信息。任务只保存在后端内存中；任务不存在或服务重启后返回 404。

停止任务：

```http
DELETE /api/proofread/tasks/{task_id}
```

将任务标记为取消。正在运行的 chunk 完成后停止后续 chunk，最终状态为 `cancelled`。

任务 SSE：

```http
GET /api/proofread/tasks/{task_id}/events
Accept: text/event-stream
```

事件名包括 `queued`、`running`、`chunk_started`、`heartbeat`、`chunk_completed`、`chunk_failed`、`completed`、`cancelled`、`error`。事件数据包含 `task_id`、`scope`、`status`、`total_chunks`、`completed_chunks`、`failed_chunks`、`issue_count` 和 `message`；chunk 相关事件额外包含 `chunk_index`、`chunk_start`、`chunk_end`、`chunk_len`、`elapsed_seconds`，失败事件包含 `error_message`。部分 chunk 失败但至少一个 chunk 成功时，最终状态为 `partial_succeeded`。

### `POST /api/sessions`

Response:

```json
{
  "session_id": "session_xxx",
  "created_at": "2026-04-25T04:00:00+00:00"
}
```

字段约定：

- `session_id`：Word 插件当前本地 session ID，用于兼容插件流程和历史记录。
- 后端不保存 AI 上下文，不保存 provider `response.id`；Responses 和 Chat 模式每次审校都是独立请求。

## 开发任务

1. 搭建 FastAPI 后端骨架，提供 `GET /health` 和 `POST /api/proofread`。
2. 定义 Pydantic schema，校验空文本，固定响应结构。
3. 实现 mock 审校服务，未配置 `AI_API_KEY` 时返回可预测的本地结果。
4. 实现 OpenAI 兼容 Responses API 与 Chat Completions client，配置 `AI_API_KEY` 后请求对应 provider API 并解析精简 JSON。
5. 改造 Word 插件任务窗格，只保留“AI 审校”正式入口、状态提示和结果展示。
6. 插件内部读取 Word 当前选区：Responses 模式优先调用流式接口展示阶段进度，失败时回退普通接口；Chat 模式直接调用普通接口。
7. 后端返回问题时，插件先展示结果，不立即写回 Word；任务窗格支持按严重程度、类别、定位状态和是否可直接替换筛选，支持逐条勾选和批量选择，并可点击单条“定位”选中 Word 原文。
8. 用户点击“应用 N 条到 Word”后，插件只写回已勾选问题；批注模式按 `start/end` 或 `global_start/global_end` 和 `original` 精准插入逐条批注；修订模式临时开启 Word 修订跟踪，将已勾选、可定位且有 `replacement` 的问题替换为 Word 原生修订。
9. 修订模式下已勾选、可定位但无 `replacement` 的问题回退为原位批注；已勾选但定位失败的问题统一合并为一条简短汇总批注，锚定在审校范围起点的第一个非空字符，找不到非空字符时退回范围起点；未勾选问题不写回 Word。
10. 长选区和全书审校通过异步任务展示分块进度；停止审校时同时中断前端请求并调用后端取消任务接口。
11. 插件支持清空当前结果、停止审校、快速/深度审校、Responses/Chat API 切换、本地历史记录清空/导出/导入；历史记录保存 `selectedIssueIds` 和 `skippedIssueCount`，开发阶段不兼容旧历史数据。
12. 补充后端测试、插件 lint/build 验证和本地联调说明。

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
OPENAI_API_BASE_URL=http://127.0.0.1:8001/v1
OPENAI_MODEL=Qwen3.6-35B-A3B-4.4bit-msq
AI_REQUEST_TIMEOUT_SECONDS=180
AI_MAX_TOKENS=32768
AI_FAST_MAX_TOKENS=16384
AI_THINKING_MAX_TOKENS=32768
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
```

`AI_API_KEY` 只允许存在于后端运行环境中，不进入 Word 插件代码、manifest、Webpack 构建变量或前端产物。未配置 `AI_API_KEY` 时，后端返回 mock 审校结果；已配置时按请求或环境配置调用 OpenAI 兼容 Responses API 或 Chat Completions。Chat 模式根据 `reasoning_enabled` 写入 `reasoning.enabled`，默认关闭。`local-omlx-dev-key` 只用于本机 oMLX 开发服务鉴权，不是真实云端密钥。AI HTTP 错误、无法清理解析的非 JSON 返回、schema 不匹配统一转换为后端 502；模型返回 Markdown 代码块、前后解释、尾随逗号或未转义控制字符时，后端会先清理再做 schema 校验。`AI_FAST_MAX_TOKENS` 用于快速审校，`AI_THINKING_MAX_TOKENS` 用于深度审校。
`BACKEND_LOG_LEVEL` 默认 `INFO`，用于打印请求模式、文本长度、provider 状态、AI provider 返回报文、问题数、定位数量、分块失败编号和错误原因，不打印完整请求正文。AI 返回报文可能包含 `original` 原文摘录。临时设为 `DEBUG` 时会额外打印后端请求体和 AI provider 请求报文，可能包含选区文本、书名、介绍和 AI 输出；任何模式都不得打印 API Key、`Authorization` 或 Bearer token。

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
  -d '{"text":"这是一段需要审校的文本。","book":{"title":"测试书名","introduction":"这是一部用于联调的测试图书。"},"session_id":"session_xxx","context":{"source":"manual-curl"}}'
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
3. 选中一段正文，或准备使用“全书正文”范围。
4. 点击“AI 审校”。
5. 确认任务窗格“运行过程”区域显示阶段进度；长选区和全书正文显示分块进度，并在“审校结果”区域显示最终结果。
6. 如果存在审校问题，确认审校完成后只在任务窗格展示结果，不立即写入 Word。
7. 使用筛选器、复选框和批量选择按钮调整待应用问题；点击单条“定位”，确认 Word 选中对应原文。
8. 点击“应用 N 条到 Word”后，确认只有已勾选的可定位问题批注在对应原文片段上；修订模式下确认已勾选、可定位且有 `replacement` 的问题生成 Word 修订。
9. 确认已勾选、可定位但无 `replacement` 的问题回退为原位批注，已勾选但未定位问题合并为一条范围起点汇总批注，未勾选问题不写回。
10. 点击“清空当前结果”，确认任务窗格清空当前结果，并创建新的本地 session。
11. 审校运行中点击“停止审校”，确认请求停止且不会插入批注。
12. 切换“快速审校/深度审校”和“Responses/Chat”，确认后续请求使用对应模式。
13. 切换“批注模式/修订模式”，确认批注模式不改正文，修订模式生成可接受/拒绝的 Word 修订。
14. 使用历史记录“清空、另存为、导入”，确认当前开发版 schema 历史可管理，旧 schema 历史导入失败并显示格式不兼容。

## 验收标准

- `GET /health` 返回 200 和 `{ "status": "ok" }`。
- 空文本请求 `POST /api/proofread` 返回 422。
- 未配置 `AI_API_KEY` 时，后端返回 mock `issues[]`。
- 配置 `AI_API_KEY` 时，后端按 `provider_api` 调用真实 Responses API 或 Chat Completions；AI provider 异常时返回 502，且错误信息不包含 Key 或 Authorization header。
- 缺少 `book` 或 `book.title` 为空时，`POST /api/proofread` 返回 422。
- `BACKEND_LOG_LEVEL=INFO` 不打印完整请求正文，但会记录 AI provider 返回报文、分块失败的 chunk 编号、范围和错误原因；`DEBUG` 会额外打印后端请求体和 AI provider 请求报文，但不包含 API Key、`Authorization` 或 Bearer token。
- `POST /api/sessions` 返回唯一 `session_id`。
- Responses 请求不携带 `previous_response_id`；同一 `session_id` 的多次审校互不续接上下文。
- AI 输出不包含 `start/end` 时，后端能按 `original` 计算 `start/end`。
- AI 输出含 `replacement` 时，后端响应保留该字段；AI 输出无 `replacement` 或空字符串时，响应为 `replacement: null`。
- `original` 与 `replacement` 去掉所有空白后完全一致的 issue 会被后端过滤。
- 重复 `original` 会按 issue 顺序定位不同 occurrence；找不到 `original` 时返回 `start/end: null`。
- `proofread_mode=fast` 与 `proofread_mode=thinking` 使用不同 prompt 和 token 上限。
- `reasoning_enabled` 默认关闭；开启时 Chat 请求体包含 `"reasoning": {"enabled": true}`，关闭时包含 `"reasoning": {"enabled": false}`。
- Responses 流式接口返回阶段进度事件和最终 `result` 事件；Chat 模式不走 SSE。
- 当前选区超过 5000 字时，插件自动创建分块任务；全书正文始终创建分块任务。
- 分块任务返回全局位置 `global_start/global_end`，前端据此定位重复原文 occurrence。
- 任务 SSE 返回分块进度、`heartbeat`、当前块耗时和失败原因；前端收到 `chunk_started` 后本地每秒刷新当前块耗时，并用后端 `heartbeat` 校准进度；SSE 不可用时前端轮询任务状态。
- 部分 chunk 失败但仍有可用结果时，任务状态为 `partial_succeeded`，前端保留可用结果并显示失败块数、失败原因和累计问题数。
- 点击停止审校会中断当前请求并取消后端异步任务。
- `npm run lint` 通过。
- `npm run build` 通过。
- Word 中书名为空或空选区点击“AI 审校”时显示错误，不调用审校接口。
- 后端不可用时显示错误，不插入空批注。
- 后端返回非空 `issues[]` 时，插件只展示结果；筛选和勾选后点击“应用 N 条到 Word”才对已选的可定位问题逐条插入批注或修订。
- 单条“定位”可选中 Word 中对应原文；重复 `original` 场景应定位到按 `start/end` 计算出的 occurrence。
- 修订模式下，插件临时将 `document.changeTrackingMode` 设为 `TrackAll`，对已选、可定位且有 `replacement` 的问题替换正文，完成后恢复原修订设置。
- 修订模式下，已选且无 `replacement` 但可定位的问题回退为原位批注；已选但定位失败的问题合并为一条锚定在审校范围起点第一个非空字符的汇总批注；未选问题不写回。
- 后端返回空 `issues[]` 时，插件显示未发现明显问题，且不插入批注。
- 审校运行中点击“停止审校”时，插件中断请求、恢复按钮、不插入批注。
- 插件本地保存最近 20 条新 schema 审校历史，可回看结果，并支持清空、另存为 JSON、导入 JSON；历史记录包含 `selectedIssueIds` 和 `skippedIssueCount`，开发阶段不兼容旧历史数据。
