# Word AI 审校助手架构说明

本文面向开发者和 AI coding agent，说明当前 V2.2 出版审校工作台的代码入口、模块边界和端到端数据流。产品目标和协作约束以 [AGENTS.md](AGENTS.md) 为准；API 契约以 [spec.md](spec.md) 为准。

## 代码入口

- FastAPI 入口：`backend/app/main.py`，负责 HTTP 路由、请求校验、后台任务启动、错误状态映射和下载响应。
- API schema：`backend/app/schemas.py`，定义 V2 project、run、candidate、memory、report、writeback 和保留的底层直接 API 请求/响应模型。
- V2 工作台编排：`backend/app/agents/workspace.py`，负责 project scoped run、分块审校、候选归并、evaluator、失败分块重试、写回编排和 run event 记录。
- V2 project 持久化：`backend/app/services/project_store.py`，负责 SQLite schema、project/run/candidate/memory/report/output 元数据和 latest run 语义。
- DOCX 能力：`backend/app/services/docx.py`、`backend/app/services/docx_service.py`、`backend/app/services/docx_tasks.py`、`backend/app/services/docx_store.py`，负责 DOCX 抽取、定位、OOXML 批注/修订写回、底层 DOCX 任务和下载索引。
- AI 配置和调用：`backend/app/services/ai_profiles.py`、`backend/app/services/ai_client.py`、`backend/app/services/ai_provider_service.py`，负责 profile 解析、Responses/Chat 适配、动态 timeout、mock fallback 和敏感信息隔离。
- Agent trace：`backend/app/agents/trace.py` 和 `backend/app/services/project_store.py` 中的 V2 run events；前者服务底层直接 API 的 Agent trace，后者服务 V2 工作台事件流。
- 文档地图和分块：`backend/app/services/document_map_service.py`、`backend/app/services/chunk_service.py`、`backend/app/services/chunking.py`。
- 前端工作台：`word-addin/src/taskpane/taskpane.ts` 负责 UI 状态和流程编排，`api.ts` 负责后端调用，`word.ts` 负责 Office.js 定位和当前选区写回，`types.ts` 负责前端类型，`debug.ts` 负责排障日志。

## V2 端到端流程

1. Word 插件读取当前选区或选择 `.docx` 文件。
2. 插件调用 `POST /api/v2/projects/selection` 或 `POST /api/v2/projects` 创建审校项目。
3. 后端保存 project，构建 document map，并生成默认 V2.2 审校计划。
4. 插件调用 `POST /api/v2/projects/{project_id}/runs` 启动后台 run。
5. `agents/workspace.py` 按 document chunks 执行 `proofread_pass`，复用已验证的 AI 审校、分块、定位和 DOCX 能力。
6. 后端执行候选归并和 evaluator 复核，保存 candidates、report、memory 摘要和 run events。
7. run 有候选时进入 `waiting_for_approval`，无候选时进入 `succeeded`；部分分块失败但有可用结果时可进入 `partial_succeeded`。
8. 编辑在插件中接受、忽略或暂缓候选；批量处理只作用于 latest run 的待处理候选。
9. 当前选区由插件通过 Office.js 写回已接受候选，然后调用 `mark-written`；DOCX 项目由后端 `writeback` 从原始 DOCX 重新生成审校后文件。
10. DOCX 写回后插件通过 `/download` 获取后端保存的新文件；如果重试失败分块后产生新增候选，旧输出保留但 project 标记 `output_stale=true`。

## 模块边界

- `agents/` 负责编排、计划、工具选择、状态推进、复核、人机确认状态和可观测事件。
- `services/` 负责可复用业务能力，尤其是 DOCX 解析、分块、AI 调用、定位、OOXML 写回、SQLite 持久化和报告生成。
- `word-addin/src/taskpane/` 负责 Office.js、任务窗格 UI、轮询/刷新、候选展示、当前选区写回和 DOCX 下载触发。
- `tools.py` 只做薄封装，不复制大量业务逻辑。
- 不从零重写已验证的 DOCX、定位和写回底层服务；V2 工作台需要调整行为时，应优先在编排层或 service 边界上做小改动。
- 修改 API 字段时，同时检查 `backend/app/schemas.py`、`backend/app/main.py`、`word-addin/src/taskpane/types.ts`、`word-addin/src/taskpane/api.ts` 和 [spec.md](spec.md)。

## 常见改动入口

- V2 run 状态、失败重试、`waiting_for_approval`、`partial_succeeded`：先看 `backend/app/agents/workspace.py`，再看 `backend/app/services/project_store.py` 和 `word-addin/src/taskpane/taskpane.ts`。
- 候选分页、批量接受/忽略、latest run 语义：先看 `project_store.py` 和 `main.py` 的 candidates 路由，再看 `api.ts`、`taskpane.ts`。
- DOCX 写回、author、批注/修订、下载丢失：先看 `docx.py`、`docx_service.py`、`workspace.py` 的 `write_approved`，再看 `project_store.py` 输出元数据和前端下载逻辑。
- AI profile、Responses/Chat 差异、Xiaomi MiMo、mock fallback：先看 `ai_profiles.py`、`ai_client.py`、`ai_provider_service.py`，再看 `.env.example` 和前端 profile 选择。
- trace 脱敏、run events、当前分块 timeout：先看 `workspace.py` 的 run event 写入和 `trace.py`，再看 `taskpane.ts` 的进度展示。
- 前端文案、按钮状态、刷新范围、定位/写回交互：先看 `taskpane.ts`，涉及 Office 文档操作再看 `word.ts`。

## 不作为代码事实来源

以下目录是本地状态、依赖、缓存或构建产物，不应作为实现真相或修改目标：

- `backend/var/`
- `backend/.venv/`
- `.pytest_cache/`
- `backend/.pytest_cache/`
- `word-addin/node_modules/`
- `word-addin/dist/`

需要验证当前行为时，优先读源码、schema、测试和 [spec.md](spec.md)，不要从上述生成物反推契约。
