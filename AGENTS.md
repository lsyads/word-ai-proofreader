# AGENTS.md

本项目是一个面向出版社责任编辑的 Word 文档审校 Agent 系统。

当前阶段目标不是继续增加零散审校功能，而是完成 V1 Agent 化改造：在保留现有 Word 插件、FastAPI 接口、DOCX 解析、分块审校、SSE 进度、批注/修订写回能力的基础上，引入 LangGraph 作为后端 Agent 编排层，并使用 LangChain 的 tool/model/schema 组件封装现有能力。

## V1 目标

完成一个可运行、可测试、可展示的 LangGraph V1 版本，使项目从“AI 审校插件”升级为“文档审校 Agent 工作流”。

V1 必须做到：

1. 保留现有功能兼容性：
   - 当前选区审校仍可用。
   - 长选区分块任务仍可用。
   - 全书 `.docx` 审校仍可用。
   - 现有 Word 插件调用链路尽量不破坏。
   - 现有 `spec.md` 中已定义接口除非必要，不做破坏性修改。

2. 新增后端 Agent 编排层：
   - 使用 LangGraph `StateGraph` 显式定义审校流程。
   - 不使用黑盒式“一个超级 prompt 跑到底”的实现。
   - 不把所有逻辑塞进单个 FastAPI endpoint。
   - 不优先使用复杂多 Agent 框架，V1 先做单图、多节点、可观测的工作流。

3. 使用 LangChain 工具组件：
   - 将现有文档解析、分块、AI 审校、结果归一化、原文定位、DOCX 写回等能力封装为 tool 或可被 tool 调用的 service function。
   - tool 输入输出必须有明确类型。
   - 复杂输入使用 Pydantic schema。
   - tool 名称使用 `snake_case`。
   - tool docstring 要清楚说明用途、输入、输出和限制。

4. 新增 Agent run trace：
   - 每次 Agent 审校任务生成 `run_id`。
   - 记录每个节点开始、结束、耗时、状态、错误信息。
   - 记录每个 chunk 的状态、问题数、失败原因、重试次数。
   - trace 不记录完整正文，不记录 API Key，不记录 Authorization。
   - trace 可以先存 SQLite，也可以先以现有任务状态结构扩展实现。

5. V1 不追求完整产品化：
   - 暂不做复杂多用户权限。
   - 暂不做完整 MCP Server。
   - 暂不做大型向量知识库。

## 推荐目录结构

在 `backend/app/` 下新增或调整以下结构：

```text
backend/app/
├── agents/
│   ├── __init__.py
│   ├── graph.py          # LangGraph StateGraph 定义与 compile
│   ├── state.py          # AgentState / TypedDict / Pydantic 状态定义
│   ├── nodes.py          # graph 节点函数
│   ├── tools.py          # LangChain tools 封装
│   ├── service.py        # Agent runner service，供 FastAPI 调用
│   └── trace.py          # run trace 记录、查询、序列化
├── services/
│   ├── proofread_service.py
│   ├── docx_service.py
│   ├── chunk_service.py
│   ├── locator_service.py
│   └── ai_provider_service.py
└── api/
    └── ...
```

注意：

agents/ 只负责 Agent 编排。
services/ 负责可复用业务能力。
tools.py 只做薄封装，不要复制大量业务逻辑。
现有函数尽量复用就迁移/引用，实在不行再重写。
