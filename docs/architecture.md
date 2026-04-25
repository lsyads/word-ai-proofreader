# Word AI 审校助手链路说明

本文说明 Word 插件、FastAPI 后端、AI provider 之间的通讯链路、数据格式和 SSE 使用位置。

## 总览

```text
Word 插件任务窗格
  -> POST /api/sessions
  -> POST /api/proofread/stream
  -> FastAPI 后端
  -> POST {OPENAI_API_BASE_URL}/responses
  -> AI provider
```

当前实现使用 provider 原生 Responses API 维护 AI session。后端只保存轻量映射：

```text
session_id -> last_response_id
```

同一个 `session_id` 的后续审校请求会把上一轮 AI provider 返回的 `response.id` 作为 `previous_response_id` 传给 `/v1/responses`，真正上下文续接由 provider 完成。

未配置 `AI_API_KEY` 时，后端走 mock 审校结果，不调用 AI provider。

## Word 插件到后端

### 1. 创建 AI session

插件打开时会先创建一个 AI 对话。点击“新建对话”时，也会调用同一个接口创建新的 AI session。

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

插件保存当前 `session_id`，后续审校请求都会带上它。

### 2. 普通审校接口

普通接口用于非流式回退。插件优先使用流式接口，只有流式读取不可用时才回退到这个接口。

Request:

```http
POST /api/proofread
Content-Type: application/json
```

```json
{
  "text": "需要审校的 Word 选区文本",
  "session_id": "session_8d7f...",
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

`issues` 为空表示未发现明显问题。插件此时只更新任务窗格，不插入 Word 批注。

### 3. 流式审校接口

插件主路径使用流式接口展示运行过程。

Request:

```http
POST /api/proofread/stream
Content-Type: application/json
Accept: text/event-stream
```

```json
{
  "text": "需要审校的 Word 选区文本",
  "session_id": "session_8d7f...",
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
data: {"stage":"calling_ai","message":"正在调用 AI 原生 Responses session。"}

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
  "input": "系统审校要求...\n\n请审校以下 Word 选区文本：\n需要审校的 Word 选区文本",
  "temperature": 0.2,
  "max_output_tokens": 1200,
  "text": {
    "format": {
      "type": "json_object"
    }
  },
  "previous_response_id": "resp_xxx"
}
```

第一次审校没有 `previous_response_id`。同一个 `session_id` 的第二次及以后审校会带上上一轮 provider 返回的 `response.id`。

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

解析成功后，后端把 provider 的 `id` 写入当前 session 的 `last_response_id`。

### 2. 流式 Responses 调用

后端流式接口会调用 provider 的流式 Responses API。

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

当前有两段 SSE：

```text
Word 插件 <-SSE- FastAPI 后端 <-SSE- AI provider
```

第一段是产品 UI 链路：

- 接口：`POST /api/proofread/stream`
- 消费方：Word 插件
- 事件：`status`、`result`、`error`
- 作用：展示运行过程、最终审校结果、错误信息

第二段是 provider 链路：

- 接口：`POST {OPENAI_API_BASE_URL}/responses`，请求体 `stream: true`
- 消费方：FastAPI 后端
- 事件：`response.created`、`response.in_progress`、`response.output_text.done`、`response.completed` 等
- 作用：接收 AI provider 原生流式输出，并在完成后解析结构化审校结果

## 批注插入规则

插件拿到最终 `issues` 后再决定是否写入 Word：

```text
issues.length > 0 -> 插入一条汇总批注
issues.length = 0 -> 只显示“未发现明显问题”，不插入批注
请求失败或用户停止 -> 不插入批注
```

## 停止审校和历史记录

插件运行审校时，主按钮会从“AI 审校”切换为“停止审校”。点击停止后，前端使用 `AbortController` 中断当前请求。

插件会在本地 `localStorage` 保存最近 20 条审校历史，用于任务窗格回看。历史记录不承担 AI 上下文续接；AI 上下文只由 provider `previous_response_id` 维护。
