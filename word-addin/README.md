# Word Add-in README

`word-addin/` 是 Word AI 审校助手的 Office 插件前端。当前 V2 提供“当前选区 / 全书正文”审校范围，并支持单独开启“深度思考”以控制后端 Chat `reasoning.enabled`。审校完成后先在任务窗格展示结果，用户点击“应用到 Word”后才写回：批注模式把可定位审校建议作为逐条批注插入对应原文片段；修订模式用 `replacement` 替换原文并生成 Word 原生修订；可定位但无 `replacement` 的建议回退为原位批注；不可定位建议合并为一条范围起点汇总批注。

## 目录结构

```text
word-addin/
├── assets/
├── src/
│   └── taskpane/
│       ├── api.ts
│       ├── history.ts
│       ├── render.ts
│       ├── taskpane.css
│       ├── taskpane.html
│       ├── taskpane.ts
│       ├── types.ts
│       └── word.ts
├── .eslintrc.json
├── babel.config.json
├── manifest.xml
├── package-lock.json
├── package.json
├── tsconfig.json
└── webpack.config.js
```

## 文件说明

- `manifest.xml`
  - Office Add-in 清单文件。
  - 定义插件 ID、名称、图标、权限、宿主应用、任务窗格地址。
  - 当前插件宿主是 Word 文档，权限是 `ReadWriteDocument`。
  - 功能区按钮文案是“AI 审校”，点击后打开 `taskpane.html`。

- `package.json`
  - Node 依赖和脚本入口。
  - 常用脚本：
    - `npm run dev-server`：启动 HTTPS Webpack dev server。
    - `npm run start`：启动调试并旁加载到 Word。
    - `npm run stop`：停止调试会话。
    - `npm run lint`：运行 Office add-in lint。
    - `npm run build`：生产构建。
    - `npm run validate`：联网验证 manifest。

- `webpack.config.js`
  - Webpack 构建配置。
  - 入口包含 `taskpane`。
  - 复制 assets 和 manifest 到构建产物。
  - 本地 dev server 使用 HTTPS，默认端口 `3000`。
  - 将同源 `/api` 请求代理到 `http://127.0.0.1:8000`，避免 Word WebView 从 HTTPS 页面直接访问 HTTP 后端导致 `Load failed`。

- `src/taskpane/taskpane.html`
  - 任务窗格 HTML。
  - 定义标题、状态提示、书名/书籍介绍输入、快速/深度审校切换、深度思考开关、Responses/Chat API 切换、批注/修订模式切换、审校范围切换、“AI 审校”、“应用到 Word”、审校结果展示区域、历史记录管理按钮。
  - 当前 MVP 不再保留独立的 WordApi 检测、读取选区、测试批注按钮。

- `src/taskpane/taskpane.css`
  - 任务窗格样式。
  - 控制页面布局、按钮、状态消息、审校结果列表。

- `src/taskpane/taskpane.ts`
  - 当前 MVP 的核心前端逻辑。
  - `Office.onReady` 后绑定“AI 审校”、“应用到 Word”、“清空当前结果”、历史清空/导出/导入按钮。
  - 点击后内部流程：
    1. 校验书名必填，并把书名和可选介绍保存在 `localStorage`。
    2. 检查 Word 批注 API 能力。
    3. 按审校范围读取当前 Word 选区或正文文本。
    4. 小选区走 `/api/proofread/stream` 或 `/api/proofread`；长选区和全书正文走 `/api/proofread/tasks`，通过 SSE 或轮询展示分块进度。
    5. 审校完成后只在任务窗格展示结果，等待用户点击“应用到 Word”。
    6. 批注模式：对有 `start/end` 的 issue 用 `search(original)` 找到原文片段并插入单条批注；定位失败的问题合并为范围起点汇总批注。
    7. 修订模式：临时将 `document.changeTrackingMode` 设为 `TrackAll`，对可定位且有 `replacement` 的 issue 用 `insertText(..., Replace)` 生成 Word 修订，完成后恢复原设置。
    8. 修订模式中可定位但无 `replacement` 的 issue 回退为原位批注；定位失败的问题合并为范围起点汇总批注。
    9. 将最近 20 条历史保存到 `localStorage`，支持清空、另存为 JSON、导入 JSON；历史记录会显示审校时的书名。

- `assets/`
  - 插件图标和 logo。
  - 被 manifest 和任务窗格引用，并由 Webpack 复制到构建产物。

- `.eslintrc.json`
  - Office add-in lint 配置。

- `babel.config.json`
  - TypeScript/Babel 转译配置。

- `tsconfig.json`
  - TypeScript 编译配置。

- `.vscode/`
  - Office add-in 模板生成的 VS Code 调试配置。

## 前后端通信

任务窗格页面运行在：

```text
https://localhost:3000/taskpane.html
```

前端请求：

```text
POST /api/proofread/stream
POST /api/proofread
```

请求体包含必填 `book.title` 和可选 `book.introduction`，后端会把它们作为 prompt 背景传给 AI；书名为空时前端会直接提示，不读取 Word 选区，也不调用后端审校接口。

Webpack dev server 转发：

```text
https://localhost:3000/api/proofread/stream -> http://127.0.0.1:8000/api/proofread/stream
https://localhost:3000/api/proofread -> http://127.0.0.1:8000/api/proofread
```

注意：

- API Key 只存在后端运行环境。
- 前端源码、manifest、Webpack 构建产物中不应出现 `AI_API_KEY`、`Authorization` 或真实 Key。
- 修改 `webpack.config.js` 后需要重启 dev server。

## 本地运行

安装依赖：

```bash
cd word-addin
npm install
```

启动 dev server：

```bash
npm run dev-server
```

旁加载到 Word：

```bash
npm run start
```

停止调试：

```bash
npm run stop
```

## 联调前置条件

先启动后端：

```bash
cd ../backend
source .venv/bin/activate
uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000 --reload
```

确认后端可用：

```bash
curl --noproxy 127.0.0.1 http://127.0.0.1:8000/health
```

确认 dev server 代理可用：

```bash
curl --noproxy localhost -k -X POST https://localhost:3000/api/proofread \
  -H 'Content-Type: application/json' \
  -d '{"text":"这是一段需要审校的文本。","book":{"title":"测试书名","introduction":"这是一部用于联调的测试图书。"}}'
```

## 测试与构建

```bash
cd word-addin
npm run lint
npm run build
npm run validate
```

说明：

- `npm run validate` 需要访问 Microsoft Office manifest validation service。
- `dist/` 是构建产物，不应提交。
- `node_modules/` 不应提交。
