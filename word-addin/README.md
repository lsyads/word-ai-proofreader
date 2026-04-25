# Word Add-in README

`word-addin/` 是 Word AI 审校助手的 Office 插件前端。当前 MVP 提供一个任务窗格入口：“AI 审校”。用户在 Word 中选中文本后点击按钮，插件调用后端 `/api/proofread/stream`，并按应用方式处理结果：批注模式把可定位审校建议作为逐条批注插入对应原文片段；修订模式用 `replacement` 替换原文并生成 Word 原生修订；不可定位建议回退为当前选区汇总批注。

## 目录结构

```text
word-addin/
├── assets/
├── src/
│   └── taskpane/
│       ├── taskpane.css
│       ├── taskpane.html
│       └── taskpane.ts
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
  - 定义标题、状态提示、快速/深度审校切换、Responses/Chat API 切换、批注/修订模式切换、“AI 审校”按钮、审校结果展示区域、历史记录管理按钮。
  - 当前 MVP 不再保留独立的 WordApi 检测、读取选区、测试批注按钮。

- `src/taskpane/taskpane.css`
  - 任务窗格样式。
  - 控制页面布局、按钮、状态消息、审校结果列表。

- `src/taskpane/taskpane.ts`
  - 当前 MVP 的核心前端逻辑。
  - `Office.onReady` 后绑定“AI 审校”、新建对话、历史清空/导出/导入按钮。
  - 点击后内部流程：
    1. 检查 Word 批注 API 能力。
    2. 读取当前 Word 选区文本。
    3. 带上 `provider_api` 和 `proofread_mode` 请求同源 `/api/proofread/stream`，不可用时回退 `/api/proofread`。
    4. 批注模式：对有 `start/end` 的 issue 用 `selection.search(original)` 找到原文片段并插入单条批注。
    5. 修订模式：临时将 `document.changeTrackingMode` 设为 `TrackAll`，对可定位且有 `replacement` 的 issue 用 `insertText(..., Replace)` 生成 Word 修订，完成后恢复原设置。
    6. 对无法定位或无 `replacement` 的 issue 调用 `selection.insertComment(...)` 插入 fallback 汇总批注。
    7. 在任务窗格展示审校结果、定位状态或错误。
    8. 将最近 20 条历史保存到 `localStorage`，支持清空、另存为 JSON、导入 JSON。

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
  -d '{"text":"这是一段需要审校的文本。"}'
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
