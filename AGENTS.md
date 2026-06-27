# AGENTS.md

本项目是一个面向出版社责任编辑的 Word 文档审校 Agent 系统。

当前阶段目标是完成 V2 出版审校 Agent 工作台改造：在复用 V1 已验证的 DOCX 解析、分块审校、AI 调用、原文定位、批注/修订写回等核心文档处理能力的基础上，把项目从“可观测的审校工作流”升级为“能规划、调用工具、复核、等待编辑确认并沉淀出版规则的 AI Agent 工作台”。

V2 可以不兼容 V1 存量数据。升级时允许重建或清空本地历史、Agent trace、任务状态、DOCX 结果索引、前端历史 schema 和旧任务快照；但不要无意义重写 V1 已经跑通的底层文档处理 service。

## V2 目标

完成一个可运行、可测试、可展示的出版审校 Agent 工作台，使责任编辑不只是“调用一次 AI 审校”，而是可以围绕一本书建立审校项目，让 Agent 制定计划、分阶段执行、给出证据、复核候选结果，并在编辑确认后写回 Word。术语、本书约定和跨章节一致性能力可沉淀为项目记忆和后续扩展，但当前 V2.2 审校计划不把它们单列为执行步骤。

V2 必须做到：

1. 面向审校项目，而不是一次性请求：
   - 引入 project/session scoped 的 Agent 运行概念。
   - 支持围绕一本书保存审校目标、文档地图、运行状态、待确认问题和审校报告。
   - V2 项目数据结构可以重新设计，不要求兼容 V1 历史记录或旧 trace。

2. 保留并复用 V1 核心文档处理逻辑：
   - 继续复用 DOCX 解析、可见文本抽取、章节/分块、AI 审校、结果归一化、原文定位、OOXML 批注/修订写回等 service。
   - `agents/` 负责 Agent 编排、计划、工具选择和复核；`services/` 负责可复用业务能力。
   - 不从零重写已经验证过的 DOCX、定位和写回能力，除非 V2 工作台的产品目标确实要求调整。

3. 建立真正的 Agent 编排层：
   - 当前代码中，复用型文本/分块 Agent runner 使用 LangGraph 定义可观测节点；V2 项目工作台使用 `agents/workspace.py` 的 project-scoped 显式服务编排，记录计划、pass、工具、候选和报告事件。
   - 不使用黑盒式“一个超级 prompt 跑到底”的实现。
   - 不把所有逻辑塞进单个 FastAPI endpoint。
   - Agent 可以根据审校目标选择工具，但工具边界、输入输出和失败处理必须清晰。

4. 建立出版审校工具注册表：
   - 将文档解析、文档地图构建、分块、AI 审校、问题归并、证据绑定、DOCX 写回等当前能力封装为 tool 或可被 tool 调用的 service function。术语检查、体例检查、跨章节一致性检查可以作为后续工具扩展保留，但不进入当前 V2.2 默认审校计划。
   - tool 输入输出必须有明确类型。
   - 复杂输入使用 Pydantic schema。
   - tool 名称使用 `snake_case`。
   - tool docstring 要清楚说明用途、输入、输出和限制。

5. 引入出版规范记忆和审校项目状态：
   - 记忆优先沉淀术语表、人物/地名/机构名一致性、出版社体例规则、编辑确认过的偏好和本书级别约定。
   - 默认不把完整正文写入长期记忆。
   - 记忆和项目状态必须可审计、可清理、可重建。
   - V2 可以重新设计 SQLite schema 或改用新的持久化结构，不承诺兼容 V1 trace 或 DOCX result index。

6. 强化人机协同与确认队列：
   - Agent 只能生成候选问题、风险说明、证据和建议。
   - 当前选区或全书写回前必须经过编辑确认。
   - 不允许未经人工确认自动静默改正文。
   - 支持批准、拒绝、暂缓、批量处理和生成审校报告。

7. 建立 V2 可观测运行记录：
   - 每次 Agent run 生成 `run_id`，项目级任务可关联 `project_id`。
   - 记录 planner 决策、节点开始/结束、耗时、状态、错误、工具调用摘要、chunk 状态、问题数、失败原因和重试次数。
   - trace 不记录完整正文，不记录 API Key，不记录 Authorization 或 Bearer token。
   - 允许为 V2 重建 trace schema；不要求读取旧 V1 trace。

## 推荐目录结构

V2 可以在 `backend/app/` 下新增或调整以下结构：

```text
backend/app/
├── agents/
│   ├── __init__.py
│   ├── graph.py              # LangGraph 编排入口与 compile
│   ├── state.py              # AgentState / ProjectState / TypedDict / Pydantic 状态定义
│   ├── planner.py            # 审校计划生成、阶段选择、工具选择
│   ├── nodes.py              # graph 节点函数
│   ├── tools.py              # LangChain tools 薄封装
│   ├── service.py            # project/session scoped Agent runner
│   ├── trace.py              # V2 run trace 记录、查询、序列化
│   └── memory.py             # 项目记忆、术语/体例规则、确认偏好
├── services/
│   ├── proofread_service.py
│   ├── docx_service.py
│   ├── chunk_service.py
│   ├── locator_service.py
│   ├── ai_provider_service.py
│   └── report_service.py
└── api/
    └── ...
```

注意：

- `agents/` 负责 Agent 编排、计划、工具选择、复核、人机确认状态和可观测运行。
- `services/` 负责可复用业务能力，尤其是 V1 已验证的文档处理逻辑。
- `tools.py` 只做薄封装，不要复制大量业务逻辑。
- V2 可以重新设计 FastAPI 接口、project/session/run/history schema；正式接口必须在实现前同步 `spec.md`。
- 临时排障记录不要写进长期文档；需要留存时放到 Git 忽略的临时目录。
