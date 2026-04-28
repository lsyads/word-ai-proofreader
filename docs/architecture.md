# Word AI 审校助手链路说明

本文说明 Word 插件、FastAPI 后端、AI provider 之间的通讯链路、数据格式和 SSE 使用位置。

## 总览

```text
Word 插件任务窗格
  -> POST /api/sessions
  -> 小选区 Responses: POST /api/proofread/stream；小选区 Chat: POST /api/proofread
  -> 长选区/全书: POST /api/proofread/tasks + SSE/轮询
  -> FastAPI 后端
  -> POST {OPENAI_API_BASE_URL}/responses 或 /chat/completions
  -> AI provider
  -> FastAPI 后端按 original 计算 start/end；分块任务额外计算 global_start/global_end
  -> Word 插件先展示结果，用户确认后按应用方式插入原文片段批注或生成 Word 修订
```

当前实现支持 Responses 和 Chat 两种 OpenAI 兼容 API。两种模式都按单轮审校处理：后端不保存 provider 上下文，不读取或写入上一轮 `response.id`，也不会向 `/v1/responses` 发送 `previous_response_id`。

未配置 `AI_API_KEY` 时，后端走 mock 审校结果，不调用 AI provider。

V2 支持统一范围审校：当前选区不超过 5000 字时使用原单段链路；当前选区超过 5000 字或选择“全书正文”时，插件创建后端内存异步任务，后端按约 3000 字顺序分块审校。分块优先在段落或句末边界切分，找不到时向后延伸到下一个边界，不硬切自然句。任务状态只存内存，服务重启后不可恢复。

## Word 插件到后端

### 1. 创建本地 session

插件打开时会先创建一个本地 session。点击“清空当前结果”时，也会调用同一个接口创建新的 session。该 ID 用于兼容插件流程和历史记录，不承担 AI 上下文续接。

Request:

```http
POST /api/sessions
```

Response:

```json
{
  "session_id": "session_8d7f...",
  "created_at": "2026-04-25T04:00:00+00:00"
}
```

插件保存当前 `session_id`，后续审校请求可以带上它。后端不会用该 ID 查找上一轮 AI 响应。

### 2. 普通审校接口

普通接口用于非流式回退，也用于 Chat 模式。插件在 Responses 模式优先使用流式接口，只有流式读取不可用时才回退到这个接口；Chat 模式直接使用这个接口。

Request:

```http
POST /api/proofread
Content-Type: application/json
```

```json
{
  "text": "需要审校的 Word 选区文本",
  "book": {
    "title": "书名",
    "introduction": "可选书籍介绍"
  },
  "session_id": "session_8d7f...",
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

`issues` 为空表示未发现明显问题。插件此时只更新任务窗格，不插入 Word 批注。

`book.title` 必填，去掉首尾空白后不能为空；`book.introduction` 可选，空白会归一为 `null`。后端会把书籍信息加入 prompt 作为背景，但书籍信息不属于待审正文，AI 仍只能对 `<text>` 内的 Word 选区文本返回问题。`provider_api` 可选，支持 `responses` 和 `chat`；不传时使用后端环境变量 `AI_PROVIDER_API`。`proofread_mode` 可选，支持 `fast` 和 `thinking`；默认 `fast`。`reasoning_enabled` 可选，默认 `false`，用于控制 Chat Completions 请求体中的 `reasoning.enabled`，不改变审校 prompt。

AI 原始输出不包含 `start/end/comment`。后端解析 AI 输出后，会按每条 issue 的 `original` 在请求文本中搜索并填充 `start/end`。重复 `original` 按 issue 顺序匹配下一处；找不到时保留该 issue，但返回 `start/end: null`。

`replacement` 是可直接替换 `original` 的正文文本。不能直接替换的问题，例如事实待核、需人工判断、体例疑问，返回 `replacement: null`；空字符串会被后端归一为 `null`。

如果 `original` 与 `replacement` 去掉所有空白后完全一致，说明该 issue 只是加/删/改空白，后端会在返回给 Word 插件前过滤掉。

### 3. 流式审校接口

插件在 Responses 模式使用流式接口展示运行过程；Chat 模式不走 SSE，直接使用普通接口。

Request:

```http
POST /api/proofread/stream
Content-Type: application/json
Accept: text/event-stream
```

```json
{
  "text": "需要审校的 Word 选区文本",
  "book": {
    "title": "书名",
    "introduction": "可选书籍介绍"
  },
  "session_id": "session_8d7f...",
  "provider_api": "responses",
  "proofread_mode": "fast",
  "context": {
    "source": "word-addin"
  }
}
```

Response 使用 SSE:

```text
event: status
data: {"stage":"received","message":"已接收选区文本。"}

event: status
data: {"stage":"calling_ai","message":"正在调用 AI Responses API。"}

event: status
data: {"stage":"normalizing","message":"已收到 AI 输出，正在解析结构化结果。"}

event: result
data: {"issues":[...]}

event: status
data: {"stage":"completed","message":"审校完成。"}
```

失败时返回：

```text
event: error
data: {"message":"AI provider returned HTTP 500"}
```

后端会给 SSE 响应设置：

```http
Cache-Control: no-cache, no-transform
Connection: keep-alive
X-Accel-Buffering: no
```

每个后端 SSE 事件前还有一行 padding 注释，降低小包被代理或 WebView 缓冲的概率。

### 4. 分块审校任务

长选区和全书正文使用异步任务接口。插件先创建任务，再优先连接任务 SSE；如果 Word WebView 或代理不支持进度流，则回退为轮询任务状态。

创建任务：

```http
POST /api/proofread/tasks
Content-Type: application/json
```

```json
{
  "text": "需要审校的长文本或全书正文",
  "book": {
    "title": "书名",
    "introduction": "可选书籍介绍"
  },
  "session_id": "session_8d7f...",
  "provider_api": "responses",
  "proofread_mode": "fast",
  "scope": "document",
  "chunk_size": 3000,
  "context": {
    "source": "word-addin",
    "flow": "chunked-task"
  }
}
```

查询任务：

```http
GET /api/proofread/tasks/{task_id}
```

Response:

```json
{
  "task_id": "task_xxx",
  "scope": "document",
  "status": "succeeded",
  "total_chunks": 3,
  "completed_chunks": 3,
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
      "chunk_index": 1,
      "global_start": 3000,
      "global_end": 3004
    }
  ],
  "error_message": null
}
```

任务 SSE：

```http
GET /api/proofread/tasks/{task_id}/events
Accept: text/event-stream
```

事件名包括 `queued`、`running`、`chunk_started`、`heartbeat`、`chunk_retry_requested`、`chunk_retrying`、`chunk_completed`、`chunk_failed`、`retry_queued`、`completed`、`cancelled`、`error`。事件数据包含当前分块进度、累计问题数和失败块数；chunk 相关事件还包含 chunk 范围、长度和 `elapsed_seconds`。前端收到 `chunk_started` 后本地每秒刷新当前块等待时长；长时间运行的 chunk 会周期性发送 `heartbeat`，用于保活和校准进度。当前 chunk 超过前端阈值后，可调用 `POST /api/proofread/tasks/{task_id}/retry-current` 取消当前 AI 调用并重试同一 chunk。失败事件包含 `error_message`，任务窗格会直接显示失败原因。停止任务使用 `DELETE /api/proofread/tasks/{task_id}`；后端标记取消后，会在当前 chunk 完成后停止后续 chunk。部分 chunk 失败但至少一个 chunk 成功时，最终状态为 `partial_succeeded`；任务结束后可调用 `POST /api/proofread/tasks/{task_id}/retry-failed` 只重试失败 chunk。

## 后端到 AI Provider

### 1. 非流式 Responses 调用

后端普通接口调用 AI provider 的 `/responses`。

Request:

```http
POST {OPENAI_API_BASE_URL}/responses
Authorization: Bearer {AI_API_KEY}
Content-Type: application/json
```

```json
{
  "model": "Qwen3.6-35B-A3B-4.4bit-msq",
  "input": "系统审校要求...\n\n书籍背景信息...\n\n<text>\n需要审校的 Word 选区文本\n</text>",
  "temperature": 0.2,
  "max_output_tokens": 16384,
  "text": {
    "format": {
      "type": "json_object"
    }
  }
}
```

每次审校都是独立请求，不携带 `previous_response_id`。

Provider Response:

```json
{
  "id": "resp_xxx",
  "object": "response",
  "status": "completed",
  "output": [
    {
      "type": "message",
      "content": [
        {
          "type": "output_text",
          "text": "{\"issues\":[]}"
        }
      ]
    }
  ]
}
```

后端从 `output[].content[].text` 提取文本，解析为：

```json
{
  "issues": []
}
```

AI JSON 中的 issue 只需要包含 `id`、`category`、`severity`、`original`、`replacement`、`suggestion`。解析时先尝试标准 JSON；如果模型返回 Markdown 代码块、前后解释、尾随逗号或未转义控制字符，后端会抽取并清理 JSON 后再做 schema 校验。整包解析仍失败时，后端会按单条 issue 抢救可校验条目并跳过坏条目。解析成功后，后端先过滤纯空白差异 issue，再按 `original` 定位并填充 `start/end`。

### 2. Chat Completions 调用

插件选择 `provider_api=chat` 时，后端调用 `/chat/completions`：

```http
POST {OPENAI_API_BASE_URL}/chat/completions
Authorization: Bearer {AI_API_KEY}
Content-Type: application/json
```

```json
{
  "model": "Qwen3.6-35B-A3B-4.4bit-msq",
  "messages": [
    {"role": "system", "content": "系统审校要求..."},
    {"role": "user", "content": "书籍背景信息...\n\n<text>\n需要审校的 Word 选区文本\n</text>"}
  ],
  "temperature": 0.2,
  "max_tokens": 16384,
  "reasoning": {"enabled": false}
}
```

Chat 模式使用标准 Chat Completions request/response：后端请求 `/v1/chat/completions`，从 `choices[0].message.content` 读取精简 issues JSON。`reasoning.enabled` 默认 `false`，用户开启“深度思考”后改为 `true`。后端会照常过滤纯空白差异 issue 并计算 `start/end`，不会写入或读取 `previous_response_id`，也不会把 Chat 结果包装成 SSE。

### 3. 快速/深度审校

`proofread_mode=fast` 使用更短 prompt 和16K 输出上限，只抓明显问题，默认 `AI_FAST_MAX_TOKENS=16384`。`proofread_mode=thinking` 使用更细审要求和32K 输出上限，默认 `AI_THINKING_MAX_TOKENS=32768`。两种模式都要求 AI 不返回 `start/end`。

### 4. 流式 Responses 调用

后端流式接口仅用于 Responses 模式，会调用 provider 的流式 Responses API。

Request 与非流式基本一致，但增加：

```json
{
  "stream": true
}
```

Provider Response 使用 SSE，例如：

```text
event: response.created
data: {"type":"response.created","response":{"id":"resp_xxx","status":"in_progress"}}

event: response.in_progress
data: {"type":"response.in_progress","response":{"id":"resp_xxx","status":"in_progress"}}

event: response.output_text.done
data: {"type":"response.output_text.done","text":"{\"issues\":[]}"}

event: response.completed
data: {"type":"response.completed","response":{"id":"resp_xxx","status":"completed","output":[...]}}
```

后端不会把 provider 原始 SSE 直接透传给 Word 插件，而是转换为插件稳定使用的事件：

```text
provider response.created        -> backend event: status
provider response.in_progress    -> backend event: status
provider response.output_text.done -> backend event: status(normalizing)
provider response.completed      -> backend event: result + status(completed)
provider response.failed         -> backend event: error
```

这样插件不需要理解不同 provider 的原始事件格式，只消费后端固定的 `status`、`result`、`error`。

## SSE 使用位置

当前 SSE 只用于 Responses 模式，有两段：

```text
Word 插件 <-SSE- FastAPI 后端 <-SSE- AI provider
```

第一段是产品 UI 链路：

- 接口：`POST /api/proofread/stream`
- 消费方：Word 插件
- 事件：`status`、`result`、`error`
- 作用：展示运行过程、最终审校结果、错误信息

第二段是 provider 链路，仅在 Responses 流式调用时存在：

- 接口：`POST {OPENAI_API_BASE_URL}/responses`，请求体 `stream: true`
- 消费方：FastAPI 后端
- 事件：`response.created`、`response.in_progress`、`response.output_text.done`、`response.completed` 等
- 作用：接收 AI provider 原生流式输出，并在完成后解析结构化审校结果

## 批注与修订应用规则

插件拿到最终 `issues` 后先渲染结果，不立即写入 Word。任务窗格会默认选中全部问题，并支持按严重程度、类别、定位状态和是否有 `replacement` 筛选；用户可以逐条勾选、批量选择，并点击单条“定位”在 Word 中选中对应原文。用户点击“应用 N 条到 Word”后才按当前应用方式写回已勾选问题：

```text
已勾选 + 批注模式 + start/end 可定位 -> 在 original 对应原文片段插入逐条批注
已勾选 + 修订模式 + start/end 可定位 + replacement 非空 -> 临时开启 TrackAll，用 replacement 替换 original，生成 Word 原生修订
已勾选 + 修订模式 + start/end 可定位 + 无 replacement -> 在 original 对应原文片段插入逐条批注
已勾选 + 定位失败 -> 在“应用 N 条到 Word”时合并为一条简短汇总批注
未勾选 -> 不写回 Word
issues.length = 0 -> 只显示“未发现明显问题”，不插入批注
请求失败或用户停止 -> 不插入批注
```

分块结果会把 `global_start/global_end` 转换成前端应用时使用的 `start/end`。当前选区分块仍在选区范围内搜索；全书分块在 `document.body` 中搜索。插件用全局位置计算重复 `original` 的 occurrence，降低长文中误命中的风险。未定位汇总批注锚定在审校范围起点的第一个非空字符；当前选区使用选区内第一个非空字符，全书使用正文第一个非空字符，找不到非空字符时退回范围起点。

修订模式会读取运行前的 `document.changeTrackingMode`，将其临时设为 `Word.ChangeTrackingMode.trackAll`，替换完成后恢复原设置。全书修订按全局位置倒序应用，减少前面的替换影响后面范围。用户随后可以在 Word 审阅面板中接受或拒绝这些修订。

批注内容以 `replacement`、`suggestion` 为核心，并带上 `category` 和 `severity`。插件历史记录会保存本次 API 类型、审校模式、应用方式、问题数、定位成功数、修订数、未定位数、是否已应用、`selectedIssueIds` 和 `skippedIssueCount`。开发阶段历史记录使用新 schema，不兼容旧历史数据。

## 停止审校和历史记录

插件运行审校时，主按钮会从“AI 审校”切换为“停止审校”。点击停止后，前端显示“正在停止审校，当前分块完成后结束”，使用 `AbortController` 中断当前请求；如果当前是分块任务，还会调用 `DELETE /api/proofread/tasks/{task_id}` 标记后端任务取消，并查询任务快照保留已完成分块返回的 issues。当前 chunk 超过等待阈值后，“重试当前分块”按钮可用；任务结束后仍有失败 chunk 时，“重试失败分块”按钮可用。

插件会在本地 `localStorage` 保存最近 20 条审校历史，用于任务窗格回看，并支持清空、另存为 JSON、导入 JSON。历史记录会显示审校时的书名、范围、分块进度、问题数、已选问题 ID、跳过数量和应用统计；导入时只接受当前开发版 schema。历史记录不承担 AI 上下文续接，后端每次审校都发起独立 AI 请求。

## 调试日志

后端 `BACKEND_LOG_LEVEL=INFO` 时记录请求入口、provider、模式、文本长度、HTTP 状态、AI provider 返回报文、问题数、定位数量、分块失败编号和错误原因，不记录完整请求正文。AI 返回报文可能包含 `original` 原文摘录，用于联调定位。临时设为 `DEBUG` 时，会额外记录后端请求体和 AI provider 请求 payload；这些内容可能包含完整选区文本、书名、介绍和 AI 输出。日志 helper 不记录 API Key、`Authorization` 或 Bearer token。DEBUG 仅用于本地调试，不建议生产开启。
