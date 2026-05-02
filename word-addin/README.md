# Word Add-in README

`word-addin/` 是 Word AI 审校助手的 Office.js 前端。它读取 Word 当前选区或全书正文，调用后端审校接口，展示结构化问题，并在用户确认后把已选问题写成 Word 批注或修订。

完整 API 契约见仓库根目录 [spec.md](../spec.md)；通讯链路见 [docs/architecture.md](../docs/architecture.md)。

## 目录结构

```text
word-addin/
├── assets/              # manifest 和任务窗格使用的图标
├── src/taskpane/
│   ├── api.ts           # 后端 API、SSE、任务轮询
│   ├── debug.ts         # 前端调试日志
│   ├── history.ts       # localStorage 历史记录
│   ├── render.ts        # 任务窗格渲染和控件状态
│   ├── taskpane.html    # 任务窗格 HTML
│   ├── taskpane.css     # 任务窗格样式
│   ├── taskpane.ts      # 页面入口和主流程编排
│   ├── types.ts         # 前端类型和常量
│   └── word.ts          # Word API 读取、定位、批注、修订
├── manifest.xml
├── package.json
├── tsconfig.json
└── webpack.config.js
```

## 主要职责

- 插件按钮打开 `https://localhost:3000/taskpane.html`。
- 校验书名，读取当前选区或 `document.body.text`。
- 当前选区 `> 7000` 字或全书正文时走 `/api/proofread/tasks`；默认分块大小与后端保持 `5000`。
- Responses 模式优先走 `/api/proofread/stream`，不可用时回退 `/api/proofread`；Chat 模式直接走 `/api/proofread`。
- 审校结果先展示在任务窗格，不自动写回 Word。
- 支持筛选、逐条勾选、批量选择、单条定位、批注模式和修订模式。
- 本地保存最近 20 条新 schema 历史，支持清空、导出 JSON 和导入 JSON。

## 本地运行

安装依赖：

```bash
cd word-addin
npm install
```

启动 HTTPS dev server：

```bash
npm run dev-server
```

默认地址：

```text
https://localhost:3000/taskpane.html
```

旁加载到 Word：

```bash
npm run start
```

停止旁加载调试：

```bash
npm run stop
```

## 后端代理

Webpack dev server 将同源 `/api` 请求代理到后端：

```text
https://localhost:3000/api/* -> http://127.0.0.1:8000/api/*
```

这样可以避免 Word WebView 从 HTTPS 任务窗格直接请求 HTTP 后端时被拦截。修改 `webpack.config.js` 后需要重启 dev server。

## 测试与构建

```bash
cd word-addin
npm run lint
npm run build
```

Manifest 联网校验：

```bash
npm run validate
```

`npm run validate` 需要访问 Microsoft Office manifest validation service，离线或网络受限时可能失败。

## 开发注意事项

- API 字段和端点以 [../spec.md](../spec.md) 为准。
- 不要在前端源码、manifest、Webpack 配置或构建产物中写入 `AI_API_KEY`、`Authorization` 或真实 Key。
- `dist/`、`node_modules/`、本地证书和调试缓存不应提交。
- 修改 Word 用户流程时，同步更新根目录 [../README.md](../README.md) 的联调流程。
