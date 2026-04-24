# Word AI 审校助手

面向出版社责任编辑的 Word AI 审校助手。当前 MVP 目标是在 Word 中选中一段文字，点击“AI 审校”，调用 FastAPI 后端返回结构化审校问题，并把审校结果作为 Word 批注插入当前选区。

## 项目结构

```text
.
├── backend/      # Python FastAPI 后端服务
├── word-addin/   # Office.js + TypeScript + Webpack 的 Word 插件
├── AGENTS.md     # 项目协作约定
└── spec.md       # MVP 技术方案和开发计划
```

## 当前 MVP 能力

- Word 任务窗格提供一个正式入口：“AI 审校”。
- 插件内部读取当前 Word 选区文本。
- 插件调用后端 `POST /api/proofread`。
- 后端返回结构化 `issues[]`。
- 插件把审校问题汇总成一条 Word 批注，插入当前选区。
- 未配置 `AI_API_KEY` 时，后端返回 mock 审校结果，方便本地联调。

## 环境变量

复制 `.env.example` 后按需填写本地配置。

```bash
cp .env.example .env
```

常用变量：

```text
AI_API_KEY=
OPENAI_API_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-4o-mini
BACKEND_HOST=127.0.0.1
BACKEND_PORT=8000
WORD_ADDIN_API_BASE_URL=http://127.0.0.1:8000
```

当前前端 MVP 在开发环境中请求同源 `/api/proofread`，由 `https://localhost:3000` 的 Webpack dev server 代理到 `http://127.0.0.1:8000`，避免 Word 任务窗格从 HTTPS 页面直接请求 HTTP 后端时被 WebView 拦截。

## 启动后端

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --host 127.0.0.1 --port 8000 --reload
```

健康检查：

```bash
curl --noproxy 127.0.0.1 http://127.0.0.1:8000/health
```

预期返回：

```json
{"status":"ok"}
```

本地 mock 审校接口：

```bash
curl --noproxy 127.0.0.1 -X POST http://127.0.0.1:8000/api/proofread \
  -H 'Content-Type: application/json' \
  -d '{"text":"这是一段需要审校的文本。","context":{"source":"manual-curl"}}'
```

## 启动 Word 插件

```bash
cd word-addin
npm install
npm run dev-server
```

插件开发服务默认运行在：

```text
https://localhost:3000/taskpane.html
```

旁加载到 Word：

```bash
cd word-addin
npm run start
```

如果 3000 端口已经被占用，先确认占用的是否是当前插件 dev server：

```bash
lsof -nP -iTCP:3000 -sTCP:LISTEN
curl --noproxy localhost -k -I https://localhost:3000/taskpane.html
```

## 联调流程

1. 启动后端，确认 `/health` 返回 `{"status":"ok"}`。
2. 启动 `word-addin` dev server。
3. 运行 `npm run start` 旁加载插件到 Word。
4. 在 Word 文档中选中一段文本。
5. 打开任务窗格，点击“AI 审校”。
6. 确认任务窗格显示审校结果。
7. 确认 Word 当前选区出现一条汇总批注。

## 测试与验证

后端测试：

```bash
cd backend
source .venv/bin/activate
python -m pytest -q
```

前端检查：

```bash
cd word-addin
npm run lint
npm run build
npm run validate
```

`npm run validate` 需要访问 Office manifest validation service，离线或网络受限时可能失败。

## 常见问题

### 任务窗格只显示旁加载提示

当前实现已经移除模板中的旁加载提示和 `display: none`。如果仍看到旧页面，通常是 Word WebView 或 dev server 缓存了旧 bundle。关闭任务窗格后重新打开，或重新运行：

```bash
cd word-addin
npm run start
```

### 点击“AI 审校”出现 `Script error.`

优先检查三件事：

1. 后端是否可访问：

   ```bash
   curl --noproxy 127.0.0.1 http://127.0.0.1:8000/health
   ```

2. dev server 是否加载了最新 bundle：

   ```bash
   curl --noproxy localhost -k -s https://localhost:3000/taskpane.js | rg "process.env|API_BASE_URL"
   ```

   当前 bundle 不应再依赖 `process.env`。

3. Word 中是否已经重新打开任务窗格。旧任务窗格可能仍持有旧的 `taskpane.js`。

### 点击“AI 审校”出现 `Load failed`

这通常表示 Word 任务窗格里的 `fetch` 请求被 WebView 网络策略拦截。当前开发环境应通过同源代理请求后端：

```text
https://localhost:3000/api/proofread -> http://127.0.0.1:8000/api/proofread
```

修改 `word-addin/webpack.config.js` 后必须重启 dev server：

```bash
cd word-addin
npm run dev-server
```

确认代理是否生效：

```bash
curl --noproxy localhost -k -X POST https://localhost:3000/api/proofread \
  -H 'Content-Type: application/json' \
  -d '{"text":"这是一段需要审校的文本。"}'
```

### 没有配置真实 AI Key

这是允许的。未配置 `AI_API_KEY` 时，后端会返回一条 mock 审校建议，用于验证 Word 插件到后端再到批注插入的闭环。

## 文档维护约定

开发完成后需要同步更新文档：

- 改接口契约时，更新 `spec.md` 和本文件的 API 说明。
- 改启动命令、端口、环境变量时，更新本文件的运行说明。
- 改 Word 插件用户流程时，更新联调流程和常见问题。
- 改项目协作规则时，更新 `AGENTS.md`。
