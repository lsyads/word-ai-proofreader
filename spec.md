# Word AI 审校助手 API 契约与验收标准

本文是当前 V2.2 Agent 工作台和兼容 V1 底层审校能力的 API 契约来源。其他文档只摘要接口或链接到本文，不重复维护完整 schema。

V2 目标是出版审校 Agent 工作台，可以重新设计 project/session/run/history schema，不要求兼容 V1 本地历史、Agent trace、任务状态、DOCX 结果索引或旧任务快照。当前 V2.2 实现当前选区和 DOCX 审校项目、文档地图、后台 Agent run、审校目标入 prompt、专项 pass、候选问题确认、项目删除、本书规则/项目记忆、approved 写回、DOCX 下载、报告和脱敏 run event trace；插件 UI 默认收敛为开始审校、本次进度、审校建议、写回和下载审校后文件。

## 范围

目标是在 Word 中完成审校闭环：当前选区由插件读取并在用户确认后写回 Word；全书正文通过上传 `.docx` 交给后端抽取、分块、审校并生成带批注或修订+批注的新 Word 文件。

包含：

- V2 当前选区审校项目和全书 `.docx` 审校项目。
- V2 文档地图、审校计划、后台 Agent run、候选问题队列、项目记忆、审校报告和脱敏 run event trace；插件把候选显示为审校建议。
- 编辑确认队列：接受、忽略、暂缓、当前选区已接受建议写回标记、DOCX 已接受建议后端写回和结果下载。
- Responses API、Chat Completions API 和未配置 Key 时的 mock fallback。
- 兼容 V1 的结构化 `issues[]`、后端原文定位、`locator`、分块全局位置、分块任务和 DOCX 后端写回能力，作为 V2 复用的底层服务与直接 API 入口。

不包含：

- 登录、云端历史、跨后端重启恢复任务。
- `.doc` 旧二进制格式、页眉页脚、脚注和尾注扫描。
- token 级模型文本流。
- 当前选区未经人工确认自动改正文。

## 系统约定

- 后端默认地址：`http://127.0.0.1:8000`。
- Word 插件开发地址：`https://localhost:3000/taskpane.html`。
- 本地开发时，插件从同源 `/api/*` 请求，由 Webpack dev server 代理到后端。
- API Key 只存在于后端运行环境，不进入前端代码、manifest、Webpack 配置、构建产物或文档。
- 后端不保存 AI provider 上下文；Responses 和 Chat 每次审校都是独立请求，不发送 `previous_response_id`。

## 数据流

```text
Word 当前选区或 DOCX 文件
  -> word-addin 创建 V2 selection 或 docx project
  -> GET /api/v2/projects/{project_id}/document-map
  -> GET /api/v2/projects/{project_id}/plan
  -> POST /api/v2/projects/{project_id}/runs
  -> backend 分块调用 AI、本地高置信规则、归并和 evaluator
  -> GET /api/v2/projects/{project_id}/candidates
  -> editor 接受、忽略或暂缓建议
  -> 当前选区: Office.js 写回已接受建议后 mark-written
  -> DOCX: POST /api/v2/projects/{project_id}/writeback 后下载审校后文件
  -> 兼容底层入口: /api/proofread、/api/proofread/tasks、/api/proofread/docx/tasks
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
- 全书 `.docx`：始终走 DOCX 文件任务。
- 默认 `chunk_size=5000`，允许范围 `500..10000`。
- 分块优先在目标长度前的段落换行和句末标点附近切分；找不到时向后延伸到下一个边界，不硬切自然句。
- 极端情况下全文无任何边界时，保留剩余文本为一个 chunk。

### 全书 DOCX 规则

- 仅支持 `.docx`；`.doc` 返回 400，提示先另存为 `.docx`。
- 审校范围包含目录可见文本、正文段落、表格文字和常见文本框文字；页眉页脚、脚注、尾注暂不纳入。
- 后端抽取可见文本时同步建立“文本字符范围 -> OOXML 文本节点”映射。AI 仍只返回精简 issue；后端把 issue 绑定到来源 chunk，优先用 chunk 内定位结果，写回前可在该 chunk 内按 `original` 二次精确搜索，再映射回 DOCX 写回，不把 Word 写回 `locator` 交给前端。
- 分块优先级：先按“章”拆分，再按“节”拆分；仍超过 7000 字时，如果可提取目录样式文本，再用目录小标题辅助拆分；仍超过 7000 字时复用当前选区的段落/句末规则。
- 批注模式生成 Word 原生批注；修订+批注模式对有 `replacement` 的问题生成 `w:del/w:ins` 原生修订并在 `replacement` 插入文本上附原因批注，无 `replacement` 或无法安全定位时降级为批注或汇总批注。
- 任务完成后后端按新文件名保存结果，例如 `书稿-AI审校-批注-20260502153000.docx`，插件显示下载入口和保留期限。
- 结果文件保存在后端 `DOCX_OUTPUT_DIR`，默认至少保留 7 天；后端通过 SQLite 索引恢复重启后的下载能力，过期或文件被外部清理时下载返回明确错误。

## API 契约

### `GET /health`

Response:

```json
{
  "status": "ok"
}
```

## V2 Agent 工作台 API

V2 API 以审校项目为核心，V2.2 支持 `selection` 和 `docx` 两类项目。DOCX 项目创建和 V1 DOCX 任务一样使用原始 DOCX bytes 作为请求体，避免 Word WebView multipart 兼容问题。V2.2 存量数据独立保存到 `AGENT_WORKSPACE_DIR`，允许重建 schema，不读取或迁移 V1/V2 旧 history、trace、task 或 DOCX result index。插件 UI 面向内部试点编辑收敛为普通使用路径：选择文本或文件、开始审校、查看建议、接受/忽略、写回或下载；Agent 内部 trace/plan/memory 和调试日志默认放入“排障信息（技术支持）”。

### `POST /api/v2/projects`

创建 V2 DOCX 审校项目，立即建立文档地图和默认审校计划。

Query：

- `filename`：必填，必须以 `.docx` 结尾。
- `book`：必填，URL 编码后的 `BookInfo` JSON。
- `review_goal`：可选，审校目标；缺省为全书出版审校目标。

Request body：原始 `.docx` 二进制。

Response：`V2ProjectResponse`，包含 `project_id/source_type/status/source_filename/text_preview/book/review_goal/run_count/latest_run_id/latest_run_status/latest_run_stage/candidate_count/pending_count/approved_count`。DOCX 项目的 `source_type` 为 `docx`。

### `POST /api/v2/projects/selection`

创建 V2 当前选区审校项目，立即建立轻量文档地图和默认审校计划。

Request:

```json
{
  "text": "当前 Word 选区文本",
  "book": {"title": "书名", "introduction": "可选书籍介绍"},
  "review_goal": "检查当前选区中的出版审校问题。",
  "session_id": "session_xxx"
}
```

Response 同 `V2ProjectResponse`，其中 `source_type` 为 `selection`，`source_filename` 为 `当前选区`，`text_preview` 为选区短预览。

### `GET /api/v2/projects`

返回最近 V2.2 项目列表，用于插件重新打开后恢复工作台。Query `limit` 默认 20，返回 `{"projects": [V2ProjectResponse]}`。

### `GET /api/v2/projects/{project_id}`

查询 V2 项目摘要。项目不存在返回 404。

### `DELETE /api/v2/projects/{project_id}`

删除 V2 项目及关联工作台数据。后端会清理 project、document map、review plan、runs、run events、candidates、report、memory 和该项目的 DOCX output 目录。项目不存在返回 404。

Response:

```json
{
  "project_id": "project_xxx",
  "deleted": true
}
```

### `GET /api/v2/projects/{project_id}/document-map`

返回文档地图摘要：`text_len/block_count/chunk_count/blocks/chunks`。DOCX 项目基于 DOCX 文档模型生成；selection 项目基于选区文本的段落和分块生成。`blocks` 和 `chunks` 只包含短 preview，不返回完整正文。

### `GET /api/v2/projects/{project_id}/plan`

返回 V2.2 审校计划。`steps[]` 包含 `step_id/title/tool_name/status/description/enabled/reason`，用于展示基础审校、术语一致性、体例规则、跨章节一致性、候选归并、evaluator 和人工确认阶段。术语一致性和跨章节一致性阶段入口保留，用于后续接入出版规则或项目记忆；默认不生成术语并用和重复数字这类低置信本地候选。

### 本地 pass 规则

V2 本地 pass 规则以可审计规则表实现，不调用模型、不写入完整正文记忆。默认只保留高置信机械体例规则：`style_consecutive_punctuation` 将连续同类句末标点规范为单个标点，且不处理 `？！` 等混合标点；`style_ascii_comma` 将中文语境英文逗号改为中文逗号；`style_halfwidth_parenthesis` 将中文正文半角括号改为全角括号。默认规则候选必须携带 `pass_name/rule_id/confidence/evidence_kind/global_start/global_end/replacement`，置信度不低于本地高置信门槛，并由归并阶段去重。`terminology_variant_pair` 和 `cross_chapter_numeric_consistency` 的 pass 入口保留，但默认不再产出低置信人工核查候选。

### `POST /api/v2/projects/{project_id}/runs`

启动一次 V2.2 Agent run。请求体包含 `session_id/ai_profile_id/provider_api/proofread_mode/reasoning_enabled/temperature`。接口快速返回 `queued` 的 `V2RunResponse`，后端后台执行 `plan_review/proofread_pass/terminology_pass/style_rule_pass/consistency_pass/merge_candidates/evaluate_candidates`。完成后如有候选问题，状态进入 `waiting_for_approval` 并进入编辑确认队列；如果候选数为 0，状态为 `succeeded`，表示审校完成且暂无需要确认的问题。

### `GET /api/v2/projects/{project_id}/runs/{run_id}`

查询 V2 run 状态、chunk 统计和候选问题数量。

### `GET /api/v2/projects/{project_id}/runs/{run_id}/events`

以 SSE 回放该 run 的事件。事件名包括 `plan_created`、`pass_started`、`pass_completed`、`tool_started`、`tool_completed`、`candidate_found`、`candidate_merged`、`candidate_evaluated`、`memory_updated`、`waiting_for_approval`、`review_completed`、`report_ready` 和 `error`。

### `GET /api/v2/projects/{project_id}/runs/{run_id}/trace`

返回脱敏 run event trace。trace 不保存完整正文、API Key、Authorization 或 Bearer token。

### `GET /api/v2/projects/{project_id}/candidates`

分页查询候选问题队列。Query `page` 默认 1，`page_size` 默认 20、最大 100；可选 `status` 过滤 `pending/approved/rejected/deferred/written`，可选 `pass_name` 过滤专项 pass。响应包含 `project_id/candidates/page/page_size/total/total_pages/has_previous/has_next`。V2.2 候选包含 `pass_name/confidence/evidence_kind/rule_id/replacement/needs_human_review/evaluation_note`，用于区分专项 pass、证据类型、可直接替换文本和 evaluator 复核意见。插件把候选显示为“审校建议”，把 `approved/rejected/written` 显示为“已接受/已忽略/已写入”；`needs_human_review=false` 表示 evaluator 未标记为人工重点判断，候选仍必须由编辑接受后才能写回。

### `POST /api/v2/projects/{project_id}/candidates/decisions`

批量更新候选问题决策。

Request:

```json
{
  "decisions": [
    {"candidate_id": "candidate_xxx", "status": "approved"}
  ]
}
```

`status` 允许 `approved`、`rejected`、`deferred`。V2.2 插件默认把 `approved/rejected` 展示为“接受/忽略”；`deferred` 保留给 API 兼容和后续更清晰的“稍后处理”设计。

编辑决策会派生轻量项目记忆，例如已批准的问题类别；不会把完整正文写入长期记忆。

### `POST /api/v2/projects/{project_id}/candidates/bulk-decisions`

批量更新该项目所有 `pending` 候选问题，忽略分页、状态筛选和 pass 筛选。插件里的“接受全部待处理/忽略全部待处理”调用该接口，并在执行前弹窗提示会处理所有待处理建议，而不是当前页。

Request:

```json
{
  "status": "approved"
}
```

`status` 允许 `approved`、`rejected`、`deferred`。编辑决策会派生轻量项目记忆；不会把完整正文写入长期记忆。

### `GET /api/v2/projects/{project_id}/memory`

返回项目记忆列表。记忆项包含 `memory_id/kind/key/value/source/confidence/created_at/updated_at`，默认保存术语、体例规则、编辑偏好、本书约定或候选摘要，不保存完整正文。

### `POST /api/v2/projects/{project_id}/memory`

手动新增项目记忆。Request 包含 `kind/key/value/source/confidence`，用于把责任编辑确认过的术语、体例或本书约定写入后续 Agent run 上下文。

### `DELETE /api/v2/projects/{project_id}/memory/{memory_id}`

删除错误或过时的项目记忆。删除后返回最新 memory 列表。

### `POST /api/v2/projects/{project_id}/writeback`

DOCX 项目只写回 `approved` 候选问题；没有 approved 问题时返回 409。请求体包含 `application_mode` 和 `fallback_summary_truncate_enabled`。写回后 approved 候选变为 `written`，项目状态变为 `written`。Selection 项目调用该接口返回 409，因为当前选区写回必须由 Word 插件通过 Office.js 完成。

### `POST /api/v2/projects/{project_id}/candidates/mark-written`

当前选区项目由 Word 插件完成 Office.js 写回后，调用该接口把已成功写回的候选标记为 `written`，并刷新项目报告。

Request:

```json
{
  "candidate_ids": ["candidate_xxx"]
}
```

Response 包含 `project_id/updated_count/candidates`。

### `GET /api/v2/projects/{project_id}/report`

返回审校报告，包含问题总数、各状态数量、severity/category/pass 分布和未处理事项。

### `GET /api/v2/projects/{project_id}/download`

下载 V2 写回后的 DOCX。尚未写回或文件丢失时返回 404。

### `POST /api/sessions`

创建本地 session ID，用于插件启动、清空当前结果和历史记录关联；不承担 AI 上下文续接。

Response:

```json
{
  "session_id": "session_xxx",
  "created_at": "2026-04-25T04:00:00+00:00"
}
```

### `GET /api/ai-profiles`

返回后端 `.env` 解析出的 AI 配置档案，供 Word 插件下拉选择。响应不包含 API Key。

Response:

```json
[
  {
    "id": "default",
    "label": "Default AI (.env)",
    "model": "Qwen3.6-35B-A3B-4.4bit-msq",
    "default_api": "responses",
    "supported_apis": ["responses", "chat"],
    "configured": true
  }
]
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
  "ai_profile_id": "default",
  "provider_api": "responses",
  "proofread_mode": "fast",
  "reasoning_enabled": false,
  "temperature": 0.2,
  "context": {
    "source": "word-addin"
  }
}
```

字段：

- `text` 必填，去掉首尾空白后不能为空。
- `book` 必填，规则见 `BookInfo`。
- `session_id` 可选，不用于 AI 上下文续接。
- `ai_profile_id` 可选；缺省使用后端 profile 列表第一项。旧 `.env` 配置会生成 `default` profile。
- `provider_api` 可选，支持 `responses`、`chat`；缺省使用所选 profile 的 `default_api`。
- `proofread_mode` 可选，支持 `fast`、`thinking`；默认 `fast`。
- `reasoning_enabled` 可选，默认 `false`；默认 Chat 供应商写入请求体 `reasoning.enabled`；Xiaomi MiMo profile 写入 `thinking.type`。
- `temperature` 可选，默认 `0.2`，范围 `0` 到 `1.5`；写入 Responses 和 Chat provider 请求体。
- `context` 可选，用于调用来源等调试信息。

Response:

```json
{
  "run_id": "agent_run_xxx",
  "issues": []
}
```

`issues` 为空表示未发现明显问题；插件只显示结果，不插入批注或修订+批注。

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

流式事件的 `data` 会附带本次审校的 `run_id`，用于查询 Agent trace。

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
  "run_id": "agent_run_xxx",
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

Response 为 `ChunkedProofreadResult`，`task_id` 必填，初始 `status` 为 `queued`，并包含本次 Agent 编排的 `run_id`。

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

事件数据包含 `task_id`、`run_id`、`scope`、`status`、`total_chunks`、`completed_chunks`、`failed_chunks`、`issue_count` 和 `message`。chunk 相关事件额外包含 `chunk_index`、`chunk_start`、`chunk_end`、`chunk_len`、`elapsed_seconds`；失败事件包含 `error_message`。

### `DELETE /api/proofread/tasks/{task_id}`

标记取消任务。正在运行的 chunk 完成后停止后续 chunk，最终状态为 `cancelled`。

### `POST /api/proofread/tasks/{task_id}/retry-current`

仅用于运行中的分块任务。当前 chunk 长时间无响应时，前端可请求该接口；后端取消当前 AI 调用并重新审校同一 chunk。任务不存在返回 404，状态不允许时返回 409。

### `POST /api/proofread/tasks/{task_id}/retry-failed`

仅用于已有失败 chunk 的终态任务。后端只重试失败 chunk，成功后移出失败集合并合并结果；仍失败的 chunk 保持失败计数。任务不存在返回 404，状态不允许时返回 409。

### `POST /api/proofread/docx/tasks`

创建全书 DOCX 文件审校任务。请求体是原始 `.docx` 二进制；元数据走 query 参数，避免 Word WebView 对 multipart 的兼容差异。

```http
POST /api/proofread/docx/tasks?filename=书稿.docx&book={...}&provider_api=responses&proofread_mode=fast&reasoning_enabled=false&temperature=0.2&application_mode=comment&fallback_summary_truncate_enabled=true
Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document
```

字段：

- `filename` 必填，必须以 `.docx` 结尾；`.doc` 返回 400。
- `book` 必填，是 URL 编码后的 `BookInfo` JSON。
- `provider_api`、`proofread_mode`、`reasoning_enabled`、`temperature` 与普通审校一致。
- `application_mode` 支持 `comment`、`revision`；`comment` 生成批注版 Word，`revision` 生成修订+批注版 Word。
- `fallback_summary_truncate_enabled` 可选，默认 `true`；为 `false` 时未定位汇总批注不限制总条数，长内容按 1500 字预算拆成多条。

Response:

```json
{
  "task_id": "docx_task_xxx",
  "run_id": "agent_run_xxx",
  "status": "queued",
  "total_chunks": 12,
  "completed_chunks": 0,
  "failed_chunks": 0,
  "issue_count": 0,
  "source_filename": "书稿.docx",
  "application_mode": "comment",
  "output_filename": null,
  "download_url": null,
  "expires_at": null,
  "retention_days": null,
  "error_message": null
}
```

### `GET /api/proofread/docx/tasks/{task_id}`

查询 DOCX 任务状态。终态成功或部分成功时，`output_filename`、`download_url`、`expires_at` 和 `retention_days` 非空。后端重启后，如果结果索引和文件仍未过期，该接口仍可返回终态快照；新生成的持久化结果会保留并返回 `run_id`，旧索引记录可能返回 `run_id: null`。

### `GET /api/proofread/docx/tasks/{task_id}/events`

订阅 DOCX 任务 SSE。事件名与普通分块任务一致，事件数据额外包含 `source_filename`、`output_filename`、`download_url`、`expires_at` 和 `retention_days`。

### `DELETE /api/proofread/docx/tasks/{task_id}`

标记取消 DOCX 任务。正在运行的 chunk 完成或取消后停止后续 chunk。

### `POST /api/proofread/docx/tasks/{task_id}/retry-current`

仅用于运行中的 DOCX 任务。当前 chunk 长时间无响应时，前端可请求后端重新审校同一 chunk。

### `POST /api/proofread/docx/tasks/{task_id}/retry-failed`

仅用于已有失败 chunk 的 DOCX 终态任务。重试完成后重新生成结果文件。

### `GET /api/proofread/docx/tasks/{task_id}/download`

下载后端生成的审校后 `.docx` 文件。结果尚未生成、已过期或文件丢失时返回明确错误；未过期结果即使后端重启也可通过持久化索引下载。

### `GET /api/agent/runs/{run_id}/trace`

查询一次 Agent 审校 run 的可观测 trace。`run_id` 会随普通审校响应、分块任务响应、DOCX 任务响应和任务 SSE 事件返回；Word 插件可用该接口在“运行过程”面板展示 Agent trace 摘要。trace 只记录节点和 chunk 元数据，不记录完整正文、API Key、Authorization 或 Bearer token。若旧 DOCX 结果没有 `run_id`，或 trace SQLite 已被清理，该接口可能无法查询并返回 404。

trace 默认保存到 `backend/var/agent-traces/traces.sqlite3`，可通过 `AGENT_TRACE_DIR` 修改目录。DOCX 下载索引使用独立 SQLite：`backend/var/docx-results/results.sqlite3`，可通过 `DOCX_OUTPUT_DIR` 修改目录。

Response:

```json
{
  "run_id": "agent_run_xxx",
  "flow": "chunked_task",
  "task_id": "task_xxx",
  "status": "succeeded",
  "created_at": "2026-05-16T00:00:00+00:00",
  "updated_at": "2026-05-16T00:00:03+00:00",
  "total_chunks": 2,
  "completed_chunks": 2,
  "failed_chunks": 0,
  "issue_count": 3,
  "error_message": null,
  "metadata": {
    "scope": "document",
    "proofread_mode": "fast"
  },
  "nodes": [
    {
      "node_name": "proofread_chunk",
      "status": "succeeded",
      "started_at": "2026-05-16T00:00:01+00:00",
      "ended_at": "2026-05-16T00:00:02+00:00",
      "elapsed_seconds": 1.0,
      "error_message": null
    }
  ],
  "chunks": [
    {
      "chunk_index": 0,
      "chunk_start": 0,
      "chunk_end": 5000,
      "chunk_len": 5000,
      "status": "succeeded",
      "issue_count": 2,
      "retry_count": 0,
      "error_message": null,
      "started_at": "2026-05-16T00:00:01+00:00",
      "ended_at": "2026-05-16T00:00:02+00:00",
      "elapsed_seconds": 1.0
    }
  ]
}
```

## Word 写回规则

- 当前选区：插件拿到 `issues` 后先展示结果，不立即写回 Word。
- 默认选中全部问题，用户可按严重程度、类别、定位状态和是否有 `replacement` 筛选。
- 已勾选 + 批注模式 + 可定位：在 `original` 对应片段插入逐条批注。
- 已勾选 + 修订+批注 + 可定位 + `replacement` 非空：临时开启 `TrackAll`，用 `replacement` 替换 `original`，生成 Word 原生修订，再把原因批注锚定到插入后的 `replacement` 文本。
- 已勾选 + 修订+批注 + 可定位 + 无 `replacement`：回退为原位批注。
- 已勾选 + 无 locator 或定位失败：拆成短汇总批注，默认最多写入 10 条，每条 1500 字以内；关闭默认截断后不限制总条数，单条超 1500 字继续拆分。
- 未勾选问题不写回。
- 已成功提交的批注或修订不回滚；某批失败时换 fresh `Word.run` 重试或降级汇总。
- 全书 DOCX：后端按 `application_mode` 直接生成新 Word 文件，插件不做逐条勾选和 Office.js 写回。

## 环境变量

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
BACKEND_HOST=127.0.0.1
BACKEND_PORT=8000
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
DOCX_OUTPUT_DIR=var/docx-results
DOCX_RETENTION_DAYS=7
AGENT_TRACE_DIR=var/agent-traces
AGENT_WORKSPACE_DIR=var/agent-workspace
WORD_ADDIN_API_BASE_URL=http://127.0.0.1:8000
```

- `AI_API_KEY` 为空时走 mock fallback。
- 不配置 `AI_PROFILES_JSON` 时，后端根据 `AI_API_KEY`、`AI_PROVIDER_API`、`OPENAI_API_BASE_URL`、`OPENAI_MODEL` 生成 `default` profile。
- `AI_PROFILES_JSON` 可选，用于配置多个 OpenAI 兼容 profile；每项包含 `id`、`label`、`api_base_url`、`api_key_env`、`model`、`default_api`、`supported_apis`。Xiaomi MiMo 示例：`{"id":"xiaomi-mimo","label":"Xiaomi MiMo","api_base_url":"https://api.xiaomimimo.com/v1","api_key_env":"MIMO_API_KEY","model":"mimo-v2.5-pro","default_api":"chat","supported_apis":["chat"]}`。
- `AI_PROVIDER_API` 默认 `responses`，用于旧 `.env` 默认 profile 的 `default_api`。
- `proofread_mode=fast` 使用 `AI_FAST_MAX_TOKENS`；`proofread_mode=thinking` 使用 `AI_THINKING_MAX_TOKENS`。
- `temperature` 是请求级参数，不需要环境变量；兼容底层 API schema 默认 `0.2`，当前 V2.2 插件工作台默认 `0.6`。
- `AGENT_TRACE_DIR` 保存 Agent run trace 的 SQLite 文件，默认 `backend/var/agent-traces`。
- `AGENT_WORKSPACE_DIR` 保存 V2 project/session/run/history schema 的 SQLite 文件和 V2 输出文件，默认 `backend/var/agent-workspace`。
- `BACKEND_LOG_LEVEL=INFO` 不打印完整请求正文；`DEBUG` 可能打印选区文本、书名、介绍和 AI 输出，仅用于本地调试。

## 验收标准

- `GET /health` 返回 200 和 `{ "status": "ok" }`。
- 空文本、缺少 `book` 或空 `book.title` 返回 422。
- 未配置 `AI_API_KEY` 时返回 mock `issues[]`。
- 配置默认 profile 的 `AI_API_KEY`，或多 profile 对应的 `api_key_env` 时，按 `ai_profile_id` 和 `provider_api` 调用 Responses 或 Chat；provider 异常返回 502，错误信息不包含 Key 或 Authorization header。
- `api_base_url=https://api.xiaomimimo.com/v1` 的 Chat profile 按 Xiaomi MiMo OpenAI-compatible Chat Completions 适配：使用 `max_completion_tokens`、`thinking.type`、`response_format={"type":"json_object"}`，不发送 `max_tokens` 或 `reasoning`。
- V2 selection 和 DOCX 项目可以创建文档地图、审校计划和 Agent run；run 完成后候选进入编辑确认队列，候选数为 0 时项目 run 状态为 `succeeded`。
- 普通审校、流式审校、分块任务、DOCX 全书任务和 V2 Agent run 都接受 `temperature`；越界返回 422，合法值会传给 AI provider；V2.2 插件默认提交 `0.6`。
- 普通审校、流式审校、分块任务和 DOCX 全书任务都会返回 `run_id`；新生成的 DOCX 持久化结果在服务重启后仍返回 `run_id`；`GET /api/agent/runs/{run_id}/trace` 可查询节点、chunk、耗时、状态、错误和重试次数，且 trace 不包含完整正文或密钥。
- `/api/ai-profiles` 不返回 Key；profile 不存在或不支持所选 `provider_api` 时返回 400。
- Responses 请求不携带 `previous_response_id`，同一 `session_id` 多次审校互不续接上下文。
- AI 输出不含 `start/end` 时，后端按 `original` 计算位置；重复 `original` 按 issue 顺序定位不同 occurrence；找不到时返回 `null`。
- 有效 `replacement` 被保留；缺失或空字符串归一为 `null`；纯空白差异 issue 被过滤。
- 默认本地规则只产出英文逗号、半角括号、连续同类句末标点三类高置信机械体例候选；候选包含 `replacement`，`evidence_kind="rule"`，且 `needs_human_review=false`。
- `terminology_variant_pair` 和 `cross_chapter_numeric_consistency` 阶段入口保留，但默认不因术语并用或重复数字产出低置信人工核查候选。
- Responses 流式接口返回阶段进度事件和最终 `result` 事件；Chat 模式不走 SSE。
- 当前选区 `> 7000` 字时创建分块任务；全书正文上传 `.docx` 创建 DOCX 文件任务；默认 `chunk_size=5000`。
- 分块任务返回 `global_start/global_end`，并把 `locator.key_start/key_end` 平移到全文坐标。
- 任务 SSE 返回分块进度、`heartbeat`、当前块耗时和失败原因；SSE 不可用时前端轮询任务状态。
- 当前 chunk 超过前端等待阈值后可重试当前分块；终态任务存在失败 chunk 时可重试失败分块。
- DOCX 文件任务支持同样的当前分块重试和失败分块重试；重试成功后重新生成结果文件。
- 部分 chunk 失败但仍有可用结果时，任务状态为 `partial_succeeded`。
- 点击停止审校会中断前端请求并取消后端异步任务；分块审校保留已收到 issues 供查看和应用。
- Word 中书名为空或空选区时显示错误，不调用审校接口。
- 后端返回非空候选或兼容 `issues[]` 时，插件只展示审校建议；编辑接受建议并点击写回按钮后才写回 Word。
- 全书 `.docx` 写回完成后插件显示后端保存的新文件名、保留期限和“下载审校后文件”按钮。
- 单条“定位”可选中对应原文；重复原文优先通过 `locator.key` 和 key 内 `original` 小范围搜索定位。
- 批注模式不改正文；修订+批注模式生成可接受/拒绝的 Word 修订，并把原因批注锚定到插入后的 `replacement` 文本，完成后恢复原修订设置。
- 后端返回空 `issues[]`、请求失败或用户停止时，不插入批注或修订+批注。
- `.docx` 全书任务即使未发现问题，也生成可下载的新文件；请求失败或用户停止时不生成新的可下载结果。
- 插件工作台可从“打开最近审校”列表手动打开历史 V2 审校；本地工作区保存项目、run、候选、报告和结果索引，不把完整正文或原始 DOCX 写入长期记忆。兼容历史功能保留清空、导出 JSON、导入 JSON 能力。
