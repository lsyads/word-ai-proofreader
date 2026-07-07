# Word AI Proofreader API Contract and Acceptance Criteria

Language: English | [简体中文](spec_cn.md)

This document is the contract source for the current V2.2 agent workbench and the retained lower-level direct review APIs. Other docs summarize interfaces or link here; they do not duplicate the complete schema.

V2 is a publishing review agent workbench. It may redesign project/session/run/history schemas and does not migrate V1/V2 old local history, agent traces, task state, DOCX result indexes, or task snapshots. The current V2.2 implementation covers current-selection and DOCX review projects, document maps, background agent runs, review goals in prompt context, a basic language review pass, candidate confirmation, project deletion, book rules/project memory, approved writeback, DOCX download, reports, and redacted run-event traces. The add-in UI defaults to start review, current progress, review suggestions, writeback, and reviewed-file download.

## Scope

The goal is to complete a review loop in Word: the add-in reads the current selection and writes it back only after user confirmation; whole-book body text is uploaded as `.docx`, extracted, chunked, reviewed, and written into a new Word file with comments or revisions plus comments.

Includes:

- V2 current-selection review projects and whole-book `.docx` review projects.
- V2 document map, review plan, background agent run, candidate queue, project memory, review report, and redacted run-event trace. The add-in displays candidates as review suggestions.
- Editor confirmation queue: accept, reject, defer, current-selection accepted-suggestion writeback marking, DOCX accepted-suggestion backend writeback, and result download.
- Responses API, Chat Completions API, and mock fallback when no key is configured.
- Retained structured `issues[]`, backend original-text location, `locator`, chunk global positions, chunked tasks, and DOCX backend writeback capabilities as lower-level services and direct API entry points reused by V2.

Does not include:

- Login, cloud history, or task recovery across backend restarts.
- Legacy `.doc` binary format, headers/footers, footnotes, or endnotes.
- Token-level model text streaming.
- Automatic current-selection body-text modification without human confirmation.

## System Conventions

- Backend default address: `http://127.0.0.1:8000`.
- Word add-in development address: `https://localhost:3000/taskpane.html`.
- During local development, the add-in calls same-origin `/api/*`; the Webpack dev server proxies those calls to the backend.
- API keys live only in the backend runtime environment and never enter frontend code, the manifest, Webpack config, build artifacts, or docs.
- The backend does not store AI provider context. Each Responses or Chat review call is independent and does not send `previous_response_id`.

## Data Flow

```text
Word current selection or DOCX file
  -> word-addin creates V2 selection or docx project
  -> GET /api/v2/projects/{project_id}/document-map
  -> GET /api/v2/projects/{project_id}/plan
  -> POST /api/v2/projects/{project_id}/runs
  -> backend calls AI by chunk, records observable passes, merges and evaluates
  -> GET /api/v2/projects/{project_id}/candidates
  -> editor accepts, rejects, or defers suggestions
  -> current selection: Office.js writes accepted suggestions, then mark-written
  -> DOCX: POST /api/v2/projects/{project_id}/writeback, then download reviewed file
  -> lower-level direct entries: /api/proofread, /api/proofread/tasks, /api/proofread/docx/tasks
```

## Common Models

### `BookInfo`

```json
{
  "title": "Book title",
  "introduction": "Optional book introduction"
}
```

- `title` is required and cannot be empty after trimming.
- `introduction` is optional; blank values normalize to `null`.
- Book information is prompt context only and is not body text to review.

### `ProofreadIssue`

```json
{
  "id": "issue-1",
  "category": "typo",
  "severity": "medium",
  "original": "Original fragment",
  "replacement": "Replacement text",
  "suggestion": "Suggested change explanation",
  "start": 0,
  "end": 4,
  "locator": {
    "key": "Original fragment",
    "key_start": 0,
    "key_end": 4,
    "original_start_in_key": 0,
    "original_end_in_key": 4,
    "strategy": "original",
    "key_occurrence_index": 0
  }
}
```

- `category` is a free string such as `typo`, `grammar`, `style`, or `fact`.
- `severity` must be `low`, `medium`, or `high`.
- Raw AI output only needs `id/category/severity/original/replacement/suggestion`; it does not return `start/end/comment/locator`.
- `replacement` may be `null`; empty strings normalize to `null`. Issues that cannot directly replace body text must return `null`.
- If `original` and `replacement` are identical after whitespace removal, or differ only by punctuation, whitespace, full-width/half-width, or Chinese/English symbol form, the backend filters the issue. `category="punctuation"` is also filtered as mechanical copyediting.
- `start/end` are computed by the backend from `original` in request text; if not found, they are `null`.
- `locator` is the Word precise-writeback location package. If unreliable, it is `null` and the frontend falls back to summary comments.

### Chunking Rules

- The V2 workbench main chain always creates a project first: current selection calls `POST /api/v2/projects/selection`; whole-book `.docx` calls `POST /api/v2/projects`; then `/runs` starts execution.
- V2 projects reuse chunking rules internally. Current-selection text and DOCX extracted text both generate a document map and chunks; the background run reviews chunks in project scope.
- In lower-level direct APIs, `/api/proofread` handles short text, `/api/proofread/tasks` handles explicit chunked text tasks, and `/api/proofread/docx/tasks` handles DOCX file tasks.
- Default `chunk_size=5000`, allowed range `500..10000`.
- Chunking prefers paragraph breaks and sentence-ending punctuation near the target length. If no boundary is found, it extends to the next boundary instead of cutting a natural sentence.
- If the entire text has no usable boundary, the remaining text is kept as one chunk.

### Whole-Book DOCX Rules

- Only `.docx` is supported. `.doc` returns 400 and asks the user to save as `.docx`.
- Review scope includes visible table-of-contents text, body paragraphs, table text, and common text-box text. Headers, footers, footnotes, and endnotes are not included.
- While extracting visible text, the backend builds a mapping from text character ranges to OOXML text nodes. AI still returns compact issues only. The backend binds each issue to its source chunk, prefers chunk-local location, can re-search `original` in the chunk before writeback, and maps the result back to DOCX. DOCX writeback location is not delegated to the frontend.
- Chunk priority: split by chapter, then section; if still over 7000 characters and table-of-contents style text is available, use small TOC headings; if still over 7000 characters, reuse current-selection paragraph/sentence-end rules.
- Comment mode writes Word-native comments. Revision-plus-comment mode writes native `w:del/w:ins` revisions for issues with `replacement` and anchors the reason comment to inserted replacement text. Issues without `replacement` or without safe location fall back to precise comments or summary comments.
- After V2 project writeback, the backend stores a new file such as `Manuscript-AI-review-comments-20260502153000.docx`, and the add-in shows a download entry. V2 writeback results are stored under the project output directory in `AGENT_WORKSPACE_DIR`. The current V2 API does not return retention or expiration fields; if the output file is externally cleaned, download returns a clear error.
- Lower-level `/api/proofread/docx/tasks` results are stored under `DOCX_OUTPUT_DIR`, retained for at least 7 days by default, and restored after backend restart through a SQLite index. Expired or externally removed files return clear download errors.

## API Contract

When changing interfaces, schemas, status values, or environment variables, check:

- Backend models and routes: `backend/app/schemas.py`, `backend/app/main.py`.
- V2 orchestration and persistence: `backend/app/agents/workspace.py`, `backend/app/services/project_store.py`.
- Frontend types and calls: `word-addin/src/taskpane/types.ts`, `word-addin/src/taskpane/api.ts`, and `taskpane.ts` when needed.
- Tests and acceptance: choose minimum checks from [TESTING.md](TESTING.md) and update the acceptance criteria below.

### `GET /health`

Response:

```json
{
  "status": "ok"
}
```

## V2 Agent Workbench API

V2 APIs are project-centered. V2.2 supports `selection` and `docx` project types. DOCX project creation uses raw DOCX bytes in the request body, like the lower-level DOCX task API, to avoid Word WebView multipart compatibility problems. V2.2 data is stored independently under `AGENT_WORKSPACE_DIR`; the schema can be rebuilt and does not read or migrate old history, traces, tasks, or DOCX result indexes. The add-in UI targets a normal editor path: choose text or file, start review, inspect suggestions, accept/reject, write back or download. Internal trace/plan/memory and debug logs are folded into troubleshooting information.

V2.2 keeps historical runs, candidates, and run events in a project, but the editor-facing default workbench view always resolves to the latest run. Project summary counts, default candidate list, review report, bulk decisions, current-selection `mark-written`, and DOCX writeback all apply only to the latest run. Historical run candidates are visible only when `run_id` is explicitly supplied to candidate queries.

### `POST /api/v2/projects`

Creates a V2 DOCX review project, immediately builds a document map, and creates the default review plan.

Query:

- `filename`: required; must end with `.docx`.
- `book`: required URL-encoded `BookInfo` JSON.
- `review_goal`: optional review goal; defaults to whole-book publishing review.

Request body: raw `.docx` bytes.

Response: `V2ProjectResponse` with `project_id/source_type/status/source_filename/text_preview/book/review_goal/created_at/updated_at/run_count/latest_run_id/latest_run_status/latest_run_stage/candidate_count/pending_count/approved_count/output_filename/download_url`. DOCX projects have `source_type="docx"`. Candidate counts default to latest-run counts.

### `POST /api/v2/projects/selection`

Creates a V2 current-selection review project, immediately builds a lightweight document map, and creates the default review plan.

Request:

```json
{
  "text": "Current Word selection text",
  "book": {"title": "Book title", "introduction": "Optional book introduction"},
  "review_goal": "Check publishing review issues in the current selection.",
  "session_id": "session_xxx"
}
```

Response is `V2ProjectResponse`; `source_type="selection"`, `source_filename="Current selection"`, and `text_preview` is a short selection preview.

### `GET /api/v2/projects`

Returns recent V2.2 projects for workbench recovery after reopening the add-in. Query `limit` defaults to 20. Response: `{"projects": [V2ProjectResponse]}`.

### `GET /api/v2/projects/{project_id}`

Returns a V2 project summary. Missing projects return 404. Candidate counts default to latest-run counts; old-run candidates are not included in the summary.

### `DELETE /api/v2/projects/{project_id}`

Deletes a V2 project and related workbench data: project, document map, review plan, runs, run events, candidates, report, memory, and that project's DOCX output directory. Missing projects return 404.

Response:

```json
{
  "project_id": "project_xxx",
  "deleted": true
}
```

### `GET /api/v2/projects/{project_id}/document-map`

Returns document map summary: `text_len/block_count/chunk_count/blocks/chunks`. DOCX projects are based on the DOCX document model; selection projects are based on paragraphs and chunks from selected text. `blocks` and `chunks` contain previews only, not full text.

### `GET /api/v2/projects/{project_id}/plan`

Returns the V2.2 review plan. `steps[]` contains `step_id/title/tool_name/status/description/enabled/reason` for plan review, basic language review, candidate merge, evaluator recheck, and human confirmation.

### Local Pass Rules

The current V2 review plan does not show local non-AI rule passes. By default, local non-AI rules emit no candidates. Mechanical copyediting items such as punctuation, whitespace, full-width/half-width, and Chinese/English symbol conversion belong to the copyediting vendor and are filtered even when AI returns them, so they do not enter the publishing-editor confirmation queue. Candidate fields still keep `pass_name/rule_id/confidence/evidence_kind/global_start/global_end/replacement` for historical data, AI candidates, and future extensions.

### `POST /api/v2/projects/{project_id}/runs`

Starts one V2.2 agent run. Request body contains `session_id/ai_profile_id/provider_api/proofread_mode/reasoning_enabled/temperature`. If the same project has a `queued` or `running` run, returns 409 to avoid mixed candidate state. The route quickly returns a queued `V2RunResponse`; backend background work executes `plan_review/proofread_pass/merge_candidates/evaluate_candidates`. Run creation uses stored document-map `chunk_count` for planning and persists run settings. DOCX parsing and body chunk generation happen in background execution. Runs with candidates end in `waiting_for_approval`; runs with zero candidates end in `succeeded`.

### `GET /api/v2/projects/{project_id}/runs/{run_id}`

Returns V2 run status, chunk counters, and candidate count. While running, if the backend has started the current chunk AI request, response may include `current_timeout`, derived from safe run-event summary fields and excluding body text, prompts, payloads, API keys, Authorization, and Bearer tokens.

The add-in polls this route for up to 60 minutes. In-progress UI derives chunk progress from `total_chunks/completed_chunks/failed_chunks`, using `completed_chunks + failed_chunks` as processed chunks. If automatic polling reaches 60 minutes, the UI stops auto-refresh only; it keeps showing a recoverable waiting state and refresh/continue controls instead of marking the backend failed.

### `POST /api/v2/projects/{project_id}/runs/{run_id}/retry-failed`

Retries failed chunks in the latest run only. It does not create a new run and does not rerun the whole book. Request body reuses `V2RunCreateRequest`; a normal new run persists settings, while old runs without settings can receive current AI settings from the frontend. The run must belong to the project and must be the latest run. Running runs and runs without failed chunks return 409; project/run mismatch returns 404.

Successful retry merges new candidates as `pending` into the same run and preserves existing `pending/approved/rejected/written` states. If output was already written and retry creates new candidates, the project returns to `waiting_for_approval` with `output_stale=true`. Old downloads remain but do not include new suggestions. `failed_chunks/candidate_count` reflect the latest state of the same run after retry.

### `GET /api/v2/projects/{project_id}/runs/{run_id}/events`

Replays that run's events as SSE. Event names include `plan_created`, `pass_started`, `pass_completed`, `tool_started`, `tool_completed`, `candidate_found`, `candidate_merged`, `candidate_evaluated`, `memory_updated`, `waiting_for_approval`, `review_completed`, `retry_queued`, `chunk_retrying`, `writeback_completed`, `report_ready`, and `error`. `tool_started` and retry `chunk_retrying` events may record dynamic timeout summary in `data.timeout`: `timeout_seconds/started_at/deadline_at/estimated_input_tokens/estimated_output_tokens/estimated_total_tokens/token_units/proofread_mode/reasoning_enabled`. Retry events record summary fields such as `chunk_index/retry_count/elapsed_seconds/error_type/message`, never full text or secrets.

### `GET /api/v2/projects/{project_id}/runs/{run_id}/trace`

Returns a redacted run-event trace. It does not store full text, API keys, Authorization, or Bearer tokens.

### `GET /api/v2/projects/{project_id}/candidates`

Paginates the candidate queue. Query `page` defaults to 1, `page_size` defaults to 20 and maxes at 100. Optional filters: `status` (`pending/approved/rejected/deferred/written`), `pass_name`, and explicit `run_id`. Without `run_id`, the latest run is used. Response includes `project_id/run_id/candidates/page/page_size/total/total_pages/has_previous/has_next`; `run_id` is `null` before any run exists. V2.2 candidates include `pass_name/confidence/evidence_kind/rule_id/replacement/needs_human_review/evaluation_note`. The add-in displays them as review suggestions and maps `approved/rejected/written` to accepted/rejected/written. `needs_human_review=false` means evaluator did not flag it for special human judgment; it still must be accepted by an editor before writeback.

### `POST /api/v2/projects/{project_id}/candidates/decisions`

Updates candidate decisions in bulk. By default, only candidates from the latest run can be updated. Historical-run candidates must first be inspected with an explicit candidate query and are not accidentally modified by the default decision route.

Request:

```json
{
  "decisions": [
    {"candidate_id": "candidate_xxx", "status": "approved"}
  ]
}
```

`status` may be `approved`, `rejected`, or `deferred`. The V2.2 add-in shows `approved/rejected` as accept/reject. `deferred` is retained for API compatibility and future clearer "handle later" design. Editor decisions derive lightweight project memory, such as approved issue categories, and never write full text into long-term memory.

### `POST /api/v2/projects/{project_id}/candidates/bulk-decisions`

Updates all `pending` candidates in the latest run, ignoring pagination, status filters, and pass filters. The add-in's "accept all pending" and "reject all pending" actions call this endpoint and first show a dialog explaining that the operation applies to all pending suggestions in the current review, not just the current page. Historical-run pending candidates are not changed by default bulk decisions.

Request:

```json
{
  "status": "approved"
}
```

`status` may be `approved`, `rejected`, or `deferred`. Editor decisions derive lightweight project memory and never write full text into long-term memory.

### `GET /api/v2/projects/{project_id}/memory`

Returns project memory. Items include `memory_id/kind/key/value/source/confidence/created_at/updated_at`. By default memory stores terminology, review conventions, editor preferences, book conventions, or candidate summaries, not full text.

### `POST /api/v2/projects/{project_id}/memory`

Adds project memory manually. Request includes `kind/key/value/source/confidence` and is used to store editor-confirmed terminology, review conventions, or book conventions for later agent runs.

### `DELETE /api/v2/projects/{project_id}/memory/{memory_id}`

Deletes incorrect or obsolete project memory and returns the latest memory list.

### `POST /api/v2/projects/{project_id}/writeback`

For DOCX projects, regenerates a result file from the original DOCX every time and includes all `written + approved` candidates from the latest run. Approved candidates from old runs are not written by default. Returns 409 when no approved issue exists. Request body contains `application_mode`, `fallback_summary_truncate_enabled`, and optional `author`. `author` defaults to `Word Proofreader` and is used as the OOXML comment/revision author; blank values use the default. After writeback, newly included approved candidates become `written`, existing `written` candidates remain, project status becomes `written`, summary includes `output_filename/download_url`, and `output_stale=false` is cleared. Selection projects return 409 because current-selection writeback must be done by the Word add-in through Office.js.

Response: `V2WritebackResponse` with `project_id/output_filename/download_url/comment_count/revision_count/fallback_count/failed_count/written_count/included_count`. `written_count` is newly transitioned candidates; `included_count` is the total number of candidates included in this generated file.

### `POST /api/v2/projects/{project_id}/candidates/mark-written`

For current-selection projects, the Word add-in calls this route after Office.js writeback to mark latest-run candidates that were actually written. The add-in submits only candidate ids from the Office.js summary that received precise comments, precise revisions, or successful summary comments. Candidates that were not written, whose fallback failed, or were omitted by summary truncation remain `approved`.

Request:

```json
{
  "candidate_ids": ["candidate_xxx"]
}
```

Response includes `project_id/updated_count/candidates`.

### `GET /api/v2/projects/{project_id}/report`

Returns the latest-run review report: total issues, counts by status, severity/category/pass distributions, and pending items. Historical-run candidates are not included in the default report.

### `GET /api/v2/projects/{project_id}/download`

Downloads the V2 writeback DOCX. Returns 404 when writeback has not happened or the output file is missing. The current V2 download does not return `expires_at` or `retention_days`; availability depends on project `output_filename` and file existence. If `output_stale=true`, the file remains downloadable but does not include newly retried suggestions that have not been written back again.

### `POST /api/sessions`

Creates a local session id for add-in startup, clearing current results, and history association. It does not continue AI context.

Response:

```json
{
  "session_id": "session_xxx",
  "created_at": "2026-04-25T04:00:00+00:00"
}
```

### `GET /api/ai-profiles`

Returns AI profiles parsed from backend `.env` for the Word add-in selector. Response does not include API keys.

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

Direct review endpoint. Chat mode uses this endpoint directly. Responses mode falls back here when streaming is unavailable.

Request:

```json
{
  "text": "Word selection text to review",
  "book": {
    "title": "Book title",
    "introduction": "Optional book introduction"
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

Fields:

- `text` is required and cannot be empty after trimming.
- `book` is required and follows `BookInfo`.
- `session_id` is optional and is not used for AI context continuation.
- `ai_profile_id` is optional; default is the first backend profile. Old `.env` settings create the `default` profile.
- `provider_api` is optional, supports `responses` and `chat`, and defaults to the selected profile's `default_api`.
- `proofread_mode` is optional, supports `fast` and `thinking`, and defaults to `fast`.
- `reasoning_enabled` is optional and defaults to `false`; default Chat providers receive `reasoning.enabled`, while Xiaomi MiMo profiles receive `thinking.type`.
- `temperature` is optional, defaults to `0.2`, accepts `0..1.5`, and is sent to Responses and Chat providers.
- `context` is optional debug information such as call source.

Response:

```json
{
  "run_id": "agent_run_xxx",
  "issues": []
}
```

Empty `issues` means no obvious issue was found. The add-in displays the result and does not insert comments or revisions.

### `POST /api/proofread/stream`

Responses-mode progress endpoint. Chat mode does not use SSE and should use `/api/proofread`.

Request is the same as `/api/proofread`.

Response uses `text/event-stream`:

```text
event: status
data: {"stage":"received","message":"Selection text received."}

event: status
data: {"stage":"calling_ai","message":"Calling AI Responses API."}

event: status
data: {"stage":"normalizing","message":"Normalizing structured review result."}

event: result
data: {"issues":[]}

event: status
data: {"stage":"completed","message":"Review completed."}
```

Failures return:

```text
event: error
data: {"message":"AI provider returned HTTP 500"}
```

Streaming event `data` includes this review's `run_id` for agent trace lookup.

### `POST /api/proofread/chunked`

Synchronous chunked review endpoint for tests, debugging, and smaller chunked tasks. The product main chain uses `/api/proofread/tasks`.

Request extends `/api/proofread` with:

```json
{
  "scope": "document",
  "chunk_size": 5000
}
```

- `scope` supports `selection` and `document`, defaulting to `selection`.
- `chunk_size` defaults to `5000`.

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
      "original": "Original fragment",
      "replacement": "Replacement text",
      "suggestion": "Suggested change",
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

Statuses: `queued`, `running`, `succeeded`, `partial_succeeded`, `failed`, `cancelled`.

### `POST /api/proofread/tasks`

Creates an in-memory asynchronous chunked task. Request is the same as `/api/proofread/chunked`. Response is `ChunkedProofreadResult`, with required `task_id`, initial `status="queued"`, and this agent orchestration's `run_id`.

Tasks are stored only in backend memory and are not recoverable after service restart.

### `GET /api/proofread/tasks/{task_id}`

Returns task status, progress, aggregated result, and errors. Missing tasks return 404.

### `GET /api/proofread/tasks/{task_id}/events`

Subscribes to task SSE:

```http
GET /api/proofread/tasks/{task_id}/events
Accept: text/event-stream
```

Events:

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

Event data includes `task_id`, `run_id`, `scope`, `status`, `total_chunks`, `completed_chunks`, `failed_chunks`, `issue_count`, and `message`. Chunk events also include `chunk_index`, `chunk_start`, `chunk_end`, `chunk_len`, and `elapsed_seconds`; failure events include `error_message`.

### `DELETE /api/proofread/tasks/{task_id}`

Marks a task cancelled. The current chunk may finish, but later chunks stop and final status becomes `cancelled`.

### `POST /api/proofread/tasks/{task_id}/retry-current`

For running chunked tasks only. If the current chunk has been waiting too long, the frontend can ask the backend to cancel the current AI call and review the same chunk again. Missing tasks return 404; invalid states return 409.

### `POST /api/proofread/tasks/{task_id}/retry-failed`

For terminal tasks with failed chunks only. The backend retries failed chunks, removes successful ones from the failure set, and keeps still-failing chunks counted as failed. Missing tasks return 404; invalid states return 409.

### `POST /api/proofread/docx/tasks`

Creates a whole-book DOCX review task. Request body is raw `.docx` bytes; metadata is query parameters to avoid Word WebView multipart compatibility differences.

```http
POST /api/proofread/docx/tasks?filename=manuscript.docx&book={...}&provider_api=responses&proofread_mode=fast&reasoning_enabled=false&temperature=0.2&application_mode=comment&fallback_summary_truncate_enabled=true&author=Word%20Proofreader
Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document
```

Fields:

- `filename` is required, must end with `.docx`; `.doc` returns 400.
- `book` is required and is URL-encoded `BookInfo` JSON.
- `provider_api`, `proofread_mode`, `reasoning_enabled`, and `temperature` match direct review.
- `application_mode` supports `comment` and `revision`. `comment` writes a comment-version Word file; `revision` writes a revision-plus-comment Word file.
- `fallback_summary_truncate_enabled` is optional and defaults to `true`; when `false`, unlocated summary comments have no total count limit, while long entries are split by a 1500-character budget.
- `author` is optional and defaults to `Word Proofreader`; it is written as the OOXML author for backend comments/revisions. Blank values use the default.

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
  "source_filename": "manuscript.docx",
  "application_mode": "comment",
  "output_filename": null,
  "download_url": null,
  "expires_at": null,
  "retention_days": null,
  "error_message": null
}
```

### `GET /api/proofread/docx/tasks/{task_id}`

Returns DOCX task status. In successful or partially successful terminal states, `output_filename`, `download_url`, `expires_at`, and `retention_days` are non-null. After backend restart, if the result index and file have not expired, this route can still return a terminal snapshot. New persistent results retain and return `run_id`; old index records may return `run_id: null`.

### `GET /api/proofread/docx/tasks/{task_id}/events`

Subscribes to DOCX task SSE. Event names match normal chunked tasks. Event data also includes `source_filename`, `output_filename`, `download_url`, `expires_at`, and `retention_days`.

### `DELETE /api/proofread/docx/tasks/{task_id}`

Marks a DOCX task cancelled. The current chunk may finish or cancel, then later chunks stop.

### `POST /api/proofread/docx/tasks/{task_id}/retry-current`

For running DOCX tasks only. If the current chunk has been waiting too long, the frontend can ask the backend to review the same chunk again.

### `POST /api/proofread/docx/tasks/{task_id}/retry-failed`

For terminal DOCX tasks with failed chunks only. After retry finishes, the backend regenerates the result file.

### `GET /api/proofread/docx/tasks/{task_id}/download`

Downloads the backend-generated reviewed `.docx`. If no result exists, the result expired, or the file is missing, returns a clear error. Non-expired results remain downloadable after backend restart through the persistent index.

### `GET /api/agent/runs/{run_id}/trace`

Returns observable trace for one agent review run. `run_id` is returned from direct review responses, chunked task responses, DOCX task responses, and task SSE events. The Word add-in can display trace summary in the run-process panel. Trace records only nodes and chunk metadata, not full text, API keys, Authorization, or Bearer tokens. If an old DOCX result lacks `run_id` or the trace SQLite was cleaned, this route may return 404.

Trace defaults to `backend/var/agent-traces/traces.sqlite3`, configurable by `AGENT_TRACE_DIR`. DOCX download index uses a separate SQLite file, `backend/var/docx-results/results.sqlite3`, configurable by `DOCX_OUTPUT_DIR`.

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

## Word Writeback Rules

- Current selection: after the add-in receives `issues`, it displays results first and does not immediately write to Word.
- All issues are selected by default; users may filter by severity, category, location status, and whether `replacement` exists.
- Selected + comment mode + locatable: insert one comment at the corresponding `original` fragment.
- Selected + revision-plus-comment mode + locatable + non-empty `replacement`: temporarily enable `TrackAll`, replace `original` with `replacement`, create a native Word revision, and anchor the reason comment to inserted replacement text.
- Whole-book DOCX backend revision writeback: prefer single-Word-run revisions. If `original` spans multiple consecutive runs in one paragraph and those runs share the same parent node, write each original run fragment as `w:del`, insert one `w:ins replacement` at the first position, and anchor the comment to inserted text. Cross-paragraph, non-consecutive text, or complex parent nodes do not get direct revisions and fall back to precise comments or summary comments. `revision_count` counts successful issue-level revisions, not low-level `w:del` nodes.
- Selected + revision-plus-comment + locatable + no `replacement`: fall back to in-place comment.
- Selected + no locator or failed location: split into short summary comments. By default, at most 10 are written and each is within 1500 characters. When truncation is disabled, total count is unlimited but entries over 1500 characters are still split.
- Unselected issues are not written back.
- Successfully inserted comments or revisions are not rolled back. If one batch fails, retry in a fresh `Word.run` or fall back to summary comments.
- Whole-book DOCX: the backend directly generates a new Word file according to `application_mode`; the add-in does not do per-item selection or Office.js writeback. Comment and revision author defaults to `Word Proofreader` and can be overridden by frontend writeback settings.

## Environment Variables

```text
AI_API_KEY=
OPENROUTER_API_KEY=
MIMO_API_KEY=

AI_PROVIDER_API=responses
AI_PROFILES_JSON='[...]'
OPENAI_API_BASE_URL=http://127.0.0.1:8001/v1
OPENAI_MODEL=Qwen3.6-35B-A3B-4.4bit-msq
AI_REQUEST_TIMEOUT_SECONDS=180
AI_REQUEST_TIMEOUT_MIN_SECONDS=60
AI_REQUEST_TIMEOUT_MAX_SECONDS=900
AI_REQUEST_TIMEOUT_BASE_SECONDS=60
AI_FAST_TIMEOUT_SECONDS_PER_1K_TOKENS=45
AI_THINKING_TIMEOUT_SECONDS_PER_1K_TOKENS=75
AI_FAST_MAX_TOKENS=8192
AI_THINKING_MAX_TOKENS=16384
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
DOCX_OUTPUT_DIR=var/docx-results
DOCX_RETENTION_DAYS=7
AGENT_TRACE_DIR=var/agent-traces
AGENT_WORKSPACE_DIR=var/agent-workspace
```

`.env.example` mirrors `.env` format and keeps variables read by backend settings or referenced by `AI_PROFILES_JSON`. Its `AI_PROFILES_JSON` contains `local-omlx`, `hy3-preview`, `mimo-v2.5`, and `mimo-v2.5-pro`; the add-in prefers a profile whose id, model, or label contains `mimo-v2.5-pro`.

- Empty key material for the selected profile's `api_key_env` uses mock fallback in the direct proofreading service.
- Without `AI_PROFILES_JSON`, the backend creates a `default` profile from `AI_API_KEY`, `AI_PROVIDER_API`, `OPENAI_API_BASE_URL`, and `OPENAI_MODEL`. These legacy fallback variables remain supported and are kept in `.env.example` because the code reads them.
- `AI_PROFILES_JSON` is optional and configures multiple OpenAI-compatible profiles. Each item contains `id`, `label`, `api_base_url`, `api_key_env`, `model`, `default_api`, and `supported_apis`. Xiaomi MiMo example: `{"id":"mimo-v2.5-pro","label":"mimo-v2.5-pro","api_base_url":"https://api.xiaomimimo.com/v1","api_key_env":"MIMO_API_KEY","model":"mimo-v2.5-pro","default_api":"chat","supported_apis":["chat"]}`.
- `AI_PROVIDER_API` defaults to `responses` and is used as the legacy `.env` default profile's `default_api`.
- `proofread_mode=fast` uses `AI_FAST_MAX_TOKENS`; `proofread_mode=thinking` uses `AI_THINKING_MAX_TOKENS`.
- AI request timeout is dynamically estimated instead of using fixed `AI_REQUEST_TIMEOUT_SECONDS`; the old field remains in the template for compatibility. The backend estimates full outgoing prompt/message `estimated_input_tokens` with `tiktoken`, adds this request's output-token limit to get `estimated_total_tokens`, calculates `token_units = ceil(max(estimated_total_tokens, 1) / 1000)`, uses `AI_FAST_TIMEOUT_SECONDS_PER_1K_TOKENS` for `fast` and `AI_THINKING_TIMEOUT_SECONDS_PER_1K_TOKENS` for `thinking` or `reasoning_enabled=true`, then computes `raw_timeout = AI_REQUEST_TIMEOUT_BASE_SECONDS + token_units * seconds_per_1k` and clamps it between `AI_REQUEST_TIMEOUT_MIN_SECONDS` and `AI_REQUEST_TIMEOUT_MAX_SECONDS`. The dynamic timeout fields are supported overrides and stay in `.env.example`; defaults are base 60 seconds, minimum 60 seconds, maximum 900 seconds, fast 45 seconds per 1k tokens, and thinking 75 seconds per 1k tokens.
- `temperature` is request-level and does not need an environment variable. The lower-level direct API schema defaults to `0.2`; the current V2.2 add-in workbench defaults to `0.6`.
- `AGENT_TRACE_DIR` stores agent run trace SQLite files, defaulting to `backend/var/agent-traces`.
- `AGENT_WORKSPACE_DIR` stores V2 project/session/run/history SQLite files and V2 output files, defaulting to `backend/var/agent-workspace`.
- `BACKEND_LOG_LEVEL=INFO` does not print full request text. `DEBUG` may print selection text, book title, introduction, and AI output; use it only for local debugging.
- `BACKEND_HOST`, `BACKEND_PORT`, and `WORD_ADDIN_API_BASE_URL` are not read by the current code and are intentionally omitted from `.env.example`. Backend listen address is controlled by uvicorn command-line arguments; the add-in development proxy is controlled by `word-addin/webpack.config.js`.

## Acceptance Criteria

- `GET /health` returns 200 and `{ "status": "ok" }`.
- Empty text, missing `book`, or blank `book.title` returns 422.
- Without key material for the selected profile, the backend returns mock `issues[]` from the direct proofreading service.
- With default-profile `AI_API_KEY` or a configured multi-profile `api_key_env`, the backend calls Responses or Chat according to `ai_profile_id` and `provider_api`; provider errors return 502 without key or Authorization header content.
- `api_base_url=https://api.xiaomimimo.com/v1` Chat profiles use Xiaomi MiMo OpenAI-compatible Chat Completions adaptation: `max_completion_tokens`, `thinking.type`, and `response_format={"type":"json_object"}`, without `max_tokens` or `reasoning`.
- V2 selection and DOCX projects can create document maps, review plans, and agent runs. After run completion, candidates enter editor confirmation; if candidate count is 0, run status is `succeeded`.
- Direct review, streaming review, chunked tasks, DOCX tasks, and V2 agent runs accept `temperature`; out-of-range values return 422, valid values are sent to the provider, and the V2.2 add-in defaults to `0.6`.
- Direct review, streaming review, chunked tasks, and DOCX tasks return `run_id`. New persistent DOCX results still return `run_id` after service restart. `GET /api/agent/runs/{run_id}/trace` returns nodes, chunks, elapsed time, status, errors, and retry counts, and traces exclude full text and secrets.
- `/api/ai-profiles` does not return keys. Unknown profiles or unsupported selected `provider_api` values return 400.
- Responses requests do not carry `previous_response_id`; repeated reviews with the same `session_id` do not continue provider context.
- If AI output lacks `start/end`, the backend calculates position from `original`; duplicate `original` fragments locate different occurrences by issue order; missing text returns `null`.
- Valid `replacement` is preserved. Missing or empty strings normalize to `null`. Pure whitespace differences, pure symbol differences, and `punctuation` issues are filtered.
- By default, local non-AI rules emit no candidates. Mechanical copyediting items such as punctuation, whitespace, full-width/half-width, and Chinese/English symbol conversion do not enter the publishing-editor confirmation queue.
- The current V2 review plan shows only real execution stages. Candidates mainly come from AI review, editor memory, or future extensions.
- Responses streaming returns stage progress events and final `result`. Chat mode does not use SSE.
- In the V2 add-in main chain, current selection and whole-book `.docx` both create V2 projects before starting project runs. Lower-level direct APIs retain `/api/proofread/tasks` and `/api/proofread/docx/tasks`, default `chunk_size=5000`.
- The V2 add-in polls project runs automatically for up to 60 minutes and shows chunk total, completed, failed, current chunk, percentage, and current-chunk dynamic AI timeout countdown. Poll timeout is a recoverable waiting state, not backend failure.
- Chunked tasks return `global_start/global_end` and shift `locator.key_start/key_end` to full-text coordinates.
- Task SSE returns chunk progress, `heartbeat`, current-chunk elapsed time, and failure reason. If SSE is unavailable, the frontend polls task state.
- Lower-level chunked tasks can retry the current chunk after the frontend wait threshold. Terminal tasks with failed chunks can retry failed chunks.
- Lower-level DOCX tasks support the same current-chunk and failed-chunk retry; after successful retry, a new result file is generated.
- If some chunks fail but usable results exist, task status is `partial_succeeded`.
- Lower-level tasks support `DELETE` cancellation. The current V2 add-in does not expose a stop-review button; long polling timeout shows refresh-progress and continue-waiting controls.
- Blank Word book title or empty selection shows an error and does not call review APIs.
- When the backend returns candidates or compatible `issues[]`, the add-in only displays suggestions. It writes back to Word only after the editor accepts suggestions and clicks writeback.
- After whole-book `.docx` V2 project writeback, the add-in shows the backend-generated filename and "Download reviewed file". Current V2 project responses do not show retention time.
- Single-item "locate" selects corresponding original text. Duplicate originals prefer `locator.key` and a narrow `original` search within that key.
- Comment mode does not change body text. Revision-plus-comment mode creates accept/rejectable Word revisions and anchors reason comments to inserted replacement text, then restores original revision settings.
- If backend returns empty `issues[]`, request fails, or there are no accepted candidates, the add-in inserts no comments or revisions.
- Lower-level `.docx` whole-book tasks generate a downloadable new file even when no issue is found. Failed or cancelled requests do not create new downloadable results. V2 DOCX project writeback requires accepted candidates and returns 409 if none are approved.
- The add-in workbench can manually reopen historical V2 reviews from the recent-review list. Local workspace stores projects, runs, candidates, reports, and result indexes, but does not write full text or original DOCX into long-term memory. The current V2 add-in does not provide JSON export, JSON import, or clear-history entry points; it supports deleting one project and clearing troubleshooting logs.
