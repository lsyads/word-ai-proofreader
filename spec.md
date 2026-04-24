# Word AI 审校助手 MVP 开发计划与技术方案

## 目标

面向出版社责任编辑，完成第一版最小闭环：在 Word 中选中一段文字，点击“AI 审校”，插件调用 FastAPI 后端，后端返回结构化审校问题，插件把审校结果作为一条汇总批注插入当前选区。

## MVP 范围

- 包含：Word 选区读取、后端审校 API、结构化问题返回、当前选区汇总批注、基础错误提示。
- 不包含：登录、审校历史、全文扫描、流式返回、多模型切换、逐条问题精准定位批注。
- 第一版后端使用 OpenAI 兼容接口；没有 `AI_API_KEY` 时返回 mock 结果，保证本地可联调。

## 技术架构

```text
Word 选区
  -> word-addin 读取选区文本
  -> POST /api/proofread
  -> backend/FastAPI 调用 AI 或 mock service
  -> 返回 issues[]
  -> word-addin 汇总 issues
  -> selection.insertComment(...)
```

### 前端

- 目录：`word-addin/`
- 技术栈：Office.js、TypeScript、Webpack。
- 本地地址：`https://localhost:3000/taskpane.html`。
- 后端地址：开发环境请求同源 `/api/proofread`，由 Webpack dev server 代理到 `http://127.0.0.1:8000`。

### 后端

- 目录：`backend/`
- 技术栈：Python、FastAPI、Pydantic、pydantic-settings、httpx、pytest。
- 本地地址：`http://127.0.0.1:8000`。
- 配置来源：后端运行环境变量；本地可通过 `uvicorn --env-file ../.env` 加载。

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

字段约定：

- `category`：问题类别，第一版允许自由字符串，例如 `typo`、`grammar`、`style`、`fact`。
- `severity`：严重程度，使用 `low`、`medium`、`high`。
- `start`、`end`：相对请求文本的字符偏移；MVP 暂不使用它们做精准定位。
- `issues` 为空表示未发现明显问题。

## 开发任务

1. 搭建 FastAPI 后端骨架，提供 `GET /health` 和 `POST /api/proofread`。
2. 定义 Pydantic schema，校验空文本，固定响应结构。
3. 实现 mock 审校服务，未配置 `AI_API_KEY` 时返回可预测的本地结果。
4. 实现 OpenAI 兼容 AI client，配置 `AI_API_KEY` 后请求 chat completions 并解析 JSON。
5. 改造 Word 插件任务窗格，只保留“AI 审校”正式入口、状态提示和结果展示。
6. 插件内部读取 Word 当前选区，调用后端，成功后插入一条汇总批注。
7. 补充后端测试、插件 lint/build 验证和本地联调说明。

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
AI_API_KEY=
OPENAI_API_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-4o-mini
AI_REQUEST_TIMEOUT_SECONDS=30
AI_MAX_TOKENS=1200
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
```

`AI_API_KEY` 只允许存在于后端运行环境中，不进入 Word 插件代码、manifest、Webpack 构建变量或前端产物。未配置 `AI_API_KEY` 时，后端返回 mock 审校结果；已配置时调用 OpenAI 兼容 Chat Completions API。AI HTTP 错误、非 JSON 返回、schema 不匹配统一转换为后端 502。

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
5. 确认当前选区出现一条包含审校问题的 Word 批注。

## 验收标准

- `GET /health` 返回 200 和 `{ "status": "ok" }`。
- 空文本请求 `POST /api/proofread` 返回 422。
- 未配置 `AI_API_KEY` 时，后端返回 mock `issues[]`。
- 配置 `AI_API_KEY` 时，后端调用真实 AI；AI provider 异常时返回 502，且错误信息不包含 Key 或 Authorization header。
- `npm run lint` 通过。
- `npm run build` 通过。
- Word 中空选区点击“AI 审校”时显示错误，不调用后端。
- 后端不可用时显示错误，不插入空批注。
- 后端返回成功时，插件在当前选区插入一条汇总批注。
