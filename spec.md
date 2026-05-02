# Word AI 审校助手 API 契约与验收标准

本文是项目的唯一 API 契约来源。其他文档只摘要接口或链接到本文，不重复维护完整 schema。

## 范围

目标是在 Word 中完成审校闭环：读取当前选区或全书正文，调用后端 AI 审校，返回结构化审校问题；插件先展示问题，用户确认后再把已选建议写成 Word 批注或修订。

包含：

- 当前选区审校和全书正文审校。
- Responses API、Chat Completions API 和未配置 Key 时的 mock fallback。
- 结构化 `issues[]`、后端原文定位、`locator`、分块全局位置。
- 审校进度 SSE、分块任务 SSE、轮询兜底、停止和重试分块。
- 本地历史记录清空、导出、导入和可回写的新 schema 历史。

不包含：

- 登录、云端历史、跨后端重启恢复任务。
- 页眉页脚、脚注、文本框等非正文扫描。
- token 级模型文本流。
- 未经人工确认自动改正文。

## 系统约定

- 后端默认地址：`http://127.0.0.1:8000`。
- Word 插件开发地址：`https://localhost:3000/taskpane.html`。
- 本地开发时，插件从同源 `/api/*` 请求，由 Webpack dev server 代理到后端。
- API Key 只存在于后端运行环境，不进入前端代码、manifest、Webpack 配置、构建产物或文档。
- 后端不保存 AI provider 上下文；Responses 和 Chat 每次审校都是独立请求，不发送 `previous_response_id`。

## 数据流

```text
Word 当前选区或全书正文
  -> word-addin 读取文本并创建/携带本地 session
  -> 小选区: Responses POST /api/proofread/stream；Chat POST /api/proofread
  -> 长选区/全书: POST /api/proofread/tasks + SSE/轮询
  -> backend 调用 OpenAI 兼容 provider 或 mock
  -> backend 解析 AI 精简 issues[]，过滤纯空白差异，计算 start/end/locator
  -> word-addin 展示结果
  -> 用户勾选并应用
  -> word-addin 按 locator 写入批注或修订，无法精准定位的问题降级为汇总批注
```

## 通用模型

### `BookInfo`

```json
{
  "title": "书名",
  "introduction": "可选书籍介绍"
}
```

- `title` 必填，去掉首尾空白后不能为空。
- `introduction` 可选；空白会归一为 `null`。
- 书籍信息仅作为 prompt 背景，不属于待审正文。

### `ProofreadIssue`

```json
{
  "id": "issue-1",
  "category": "typo",
  "severity": "medium",
  "original": "原文片段",
  "replacement": "可直接替换原文的新文本",
  "suggestion": "修改建议说明",
  "start": 0,
  "end": 4,
  "locator": {
    "key": "原文片段",
    "key_start": 0,
    "key_end": 4,
    "original_start_in_key": 0,
    "original_end_in_key": 4,
    "strategy": "original",
    "key_occurrence_index": 0
  }
}
```

- `category` 为自由字符串，例如 `typo`、`grammar`、`style`、`fact`。
- `severity` 只允许 `low`、`medium`、`high`。
- AI 原始输出只需要包含 `id/category/severity/original/replacement/suggestion`，不返回 `start/end/comment/locator`。
- `replacement` 可为 `null`；空字符串归一为 `null`。不能直接替换正文的问题必须返回 `null`。
- 如果 `original` 与 `replacement` 去掉所有空白后完全一致，后端过滤该 issue。
- `start/end` 由后端按 `original` 在请求文本中计算；找不到时为 `null`。
- `locator` 是 Word 精准写回定位包；不可靠时为 `null`，前端应用时降级为汇总批注。

### 分块规则

- 当前选区长度 `<= 7000`：单段审校。
- 当前选区长度 `> 7000`：分块任务。
- 全书正文：始终分块任务。
- 默认 `chunk_size=5000`，允许范围 `500..10000`。
- 分块优先在目标长度前的段落换行和句末标点附近切分；找不到时向后延伸到下一个边界，不硬切自然句。
- 极端情况下全文无任何边界时，保留剩余文本为一个 chunk。

## API 契约

### `GET /health`

Response:

```json
{
  "status": "ok"
}
```

### `POST /api/sessions`

创建本地 session ID，用于插件启动、清空当前结果和历史记录关联；不承担 AI 上下文续接。

Response:

```json
{
  "session_id": "session_xxx",
  "created_at": "2026-04-25T04:00:00+00:00"
}
```

### `POST /api/proofread`

普通审校接口。Chat 模式直接使用该接口；Responses 模式在流式接口不可用时回退到该接口。

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

字段：

- `text` 必填，去掉首尾空白后不能为空。
- `book` 必填，规则见 `BookInfo`。
- `session_id` 可选，不用于 AI 上下文续接。
- `provider_api` 可选，支持 `responses`、`chat`；缺省使用 `AI_PROVIDER_API`，默认 `responses`。
- `proofread_mode` 可选，支持 `fast`、`thinking`；默认 `fast`。
- `reasoning_enabled` 可选，默认 `false`；Chat 模式下写入请求体 `reasoning.enabled`。
- `context` 可选，用于调用来源等调试信息。

Response:

```json
{
  "issues": []
}
```

`issues` 为空表示未发现明显问题；插件只显示结果，不插入批注或修订。

### `POST /api/proofread/stream`

Responses 模式的审校进度接口。Chat 模式不走 SSE，应使用 `/api/proofread`。

Request 与 `/api/proofread` 相同。

Response 使用 `text/event-stream`：

```text
event: status
data: {"stage":"received","message":"已接收选区文本。"}

event: status
data: {"stage":"calling_ai","message":"正在调用 AI Responses API。"}

event: status
data: {"stage":"normalizing","message":"正在整理结构化审校结果。"}

event: result
data: {"issues":[]}

event: status
data: {"stage":"completed","message":"审校完成。"}
```

失败时返回：

```text
event: error
data: {"message":"AI provider returned HTTP 500"}
```

### `POST /api/proofread/chunked`

同步分块审校接口，用于测试、调试和较小分块任务。产品主链路使用 `/api/proofread/tasks`。

Request 在 `/api/proofread` 基础上增加：

```json
{
  "scope": "document",
  "chunk_size": 5000
}
```

- `scope` 支持 `selection`、`document`；默认 `selection`。
- `chunk_size` 默认 `5000`。

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
      "locator": null,
      "chunk_index": 0,
      "global_start": 0,
      "global_end": 4
    }
  ],
  "error_message": null
}
```

状态值支持 `queued`、`running`、`succeeded`、`partial_succeeded`、`failed`、`cancelled`。

### `POST /api/proofread/tasks`

创建内存异步分块任务。请求与 `/api/proofread/chunked` 相同。

Response 为 `ChunkedProofreadResult`，`task_id` 必填，初始 `status` 为 `queued`。

任务只保存在后端内存中；服务重启后不可恢复。

### `GET /api/proofread/tasks/{task_id}`

查询任务状态、进度、聚合结果和错误信息。任务不存在时返回 404。

### `GET /api/proofread/tasks/{task_id}/events`

订阅任务 SSE。

```http
GET /api/proofread/tasks/{task_id}/events
Accept: text/event-stream
```

事件名包括：

```text
queued
running
chunk_started
heartbeat
chunk_retry_requested
chunk_retrying
chunk_completed
chunk_failed
retry_queued
completed
cancelled
error
```

事件数据包含 `task_id`、`scope`、`status`、`total_chunks`、`completed_chunks`、`failed_chunks`、`issue_count` 和 `message`。chunk 相关事件额外包含 `chunk_index`、`chunk_start`、`chunk_end`、`chunk_len`、`elapsed_seconds`；失败事件包含 `error_message`。

### `DELETE /api/proofread/tasks/{task_id}`

标记取消任务。正在运行的 chunk 完成后停止后续 chunk，最终状态为 `cancelled`。

### `POST /api/proofread/tasks/{task_id}/retry-current`

仅用于运行中的分块任务。当前 chunk 长时间无响应时，前端可请求该接口；后端取消当前 AI 调用并重新审校同一 chunk。任务不存在返回 404，状态不允许时返回 409。

### `POST /api/proofread/tasks/{task_id}/retry-failed`

仅用于已有失败 chunk 的终态任务。后端只重试失败 chunk，成功后移出失败集合并合并结果；仍失败的 chunk 保持失败计数。任务不存在返回 404，状态不允许时返回 409。

## Word 写回规则

- 插件拿到 `issues` 后先展示结果，不立即写回 Word。
- 默认选中全部问题，用户可按严重程度、类别、定位状态和是否有 `replacement` 筛选。
- 已勾选 + 批注模式 + 可定位：在 `original` 对应片段插入逐条批注。
- 已勾选 + 修订模式 + 可定位 + `replacement` 非空：临时开启 `TrackAll`，用 `replacement` 替换 `original`，生成 Word 原生修订。
- 已勾选 + 修订模式 + 可定位 + 无 `replacement`：回退为原位批注。
- 已勾选 + 无 locator 或定位失败：拆成短汇总批注，默认最多写入 10 条，每条按约 1200 字预算。
- 未勾选问题不写回。
- 已成功提交的批注或修订不回滚；某批失败时换 fresh `Word.run` 重试或降级汇总。

## 环境变量

```text
AI_API_KEY=local-omlx-dev-key
AI_PROVIDER_API=responses
OPENAI_API_BASE_URL=http://127.0.0.1:8001/v1
OPENAI_MODEL=Qwen3.6-35B-A3B-4.4bit-msq
AI_REQUEST_TIMEOUT_SECONDS=180
AI_MAX_TOKENS=32768
AI_FAST_MAX_TOKENS=8192
AI_THINKING_MAX_TOKENS=16384
BACKEND_HOST=127.0.0.1
BACKEND_PORT=8000
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
WORD_ADDIN_API_BASE_URL=http://127.0.0.1:8000
```

- `AI_API_KEY` 为空时走 mock fallback。
- `AI_PROVIDER_API` 默认 `responses`。
- `proofread_mode=fast` 使用 `AI_FAST_MAX_TOKENS`；`proofread_mode=thinking` 使用 `AI_THINKING_MAX_TOKENS`。
- `BACKEND_LOG_LEVEL=INFO` 不打印完整请求正文；`DEBUG` 可能打印选区文本、书名、介绍和 AI 输出，仅用于本地调试。

## 验收标准

- `GET /health` 返回 200 和 `{ "status": "ok" }`。
- 空文本、缺少 `book` 或空 `book.title` 返回 422。
- 未配置 `AI_API_KEY` 时返回 mock `issues[]`。
- 配置 `AI_API_KEY` 时按 `provider_api` 调用 Responses 或 Chat；provider 异常返回 502，错误信息不包含 Key 或 Authorization header。
- Responses 请求不携带 `previous_response_id`，同一 `session_id` 多次审校互不续接上下文。
- AI 输出不含 `start/end` 时，后端按 `original` 计算位置；重复 `original` 按 issue 顺序定位不同 occurrence；找不到时返回 `null`。
- 有效 `replacement` 被保留；缺失或空字符串归一为 `null`；纯空白差异 issue 被过滤。
- Responses 流式接口返回阶段进度事件和最终 `result` 事件；Chat 模式不走 SSE。
- 当前选区 `> 7000` 字时创建分块任务；全书正文始终创建分块任务；默认 `chunk_size=5000`。
- 分块任务返回 `global_start/global_end`，并把 `locator.key_start/key_end` 平移到全文坐标。
- 任务 SSE 返回分块进度、`heartbeat`、当前块耗时和失败原因；SSE 不可用时前端轮询任务状态。
- 当前 chunk 超过前端等待阈值后可重试当前分块；终态任务存在失败 chunk 时可重试失败分块。
- 部分 chunk 失败但仍有可用结果时，任务状态为 `partial_succeeded`。
- 点击停止审校会中断前端请求并取消后端异步任务；分块审校保留已收到 issues 供查看和应用。
- Word 中书名为空或空选区时显示错误，不调用审校接口。
- 后端返回非空 `issues[]` 时，插件只展示结果；点击“应用 N 条到 Word”后才写回已选问题。
- 单条“定位”可选中对应原文；重复原文优先通过 `locator.key` 和 key 内 `original` 小范围搜索定位。
- 批注模式不改正文；修订模式生成可接受/拒绝的 Word 修订，并在完成后恢复原修订设置。
- 后端返回空 `issues[]`、请求失败或用户停止时，不插入批注或修订。
- 插件本地保存最近 20 条新 schema 历史，支持清空、导出 JSON、导入 JSON；新历史可再次筛选、勾选、定位和应用，旧历史缺少 locator occurrence 时只读。
