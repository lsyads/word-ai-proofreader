# Word AI 审校助手

面向出版社责任编辑的 Word AI 审校助手。

## 项目结构

```text
.
├── backend/      # Python FastAPI 后端服务
├── scripts/      # 本地开发辅助脚本
├── word-addin/   # Office.js + TypeScript + Webpack 的 Word 插件
├── AGENTS.md     # 项目协作约定
└── spec.md       # 技术方案和开发计划
```

## 当前 MVP 能力

- Word 任务窗格提供一个正式入口：“AI 审校”。
- 插件打开时会创建一个本地 session；点击“新建对话”会创建新的本地 session。后端不使用 session 续接 AI 上下文。
- 插件内部读取当前 Word 选区文本。
- 插件可切换“快速审校/深度审校”和 `Responses/Chat` API。
- 插件可切换“批注模式/修订模式”；默认批注模式，避免默认改正文。
- 插件调用后端 `POST /api/proofread`。
- Responses 模式下，插件优先调用后端 `POST /api/proofread/stream`，并在任务窗格“运行过程”区域展示阶段进度；流式不可用时自动回退 `POST /api/proofread`。Chat 模式直接调用 `POST /api/proofread`，由后端使用标准 Chat Completions。
- AI 原始输出只包含精简 `issues[]`，不返回 `start/end`；其中 `replacement` 是可直接替换正文的新文本。后端会过滤纯空白差异 issue，并按 `original` 在选区文本中搜索、计算位置。
- 批注模式下，插件对可定位问题逐条在对应原文片段插入批注；修订模式下，插件临时开启 Word 修订跟踪，用 `replacement` 替换可定位原文并生成原生修订；不可定位或无 `replacement` 的问题回退为当前选区汇总批注。
- 插件支持停止当前审校，并在本地保存最近 20 条审校历史用于回看、清空、另存为 JSON 和导入 JSON。
- 未配置 `AI_API_KEY` 时，后端返回 mock 审校结果，方便本地联调。

完整通讯链路和数据格式见 [docs/architecture.md](docs/architecture.md)。

## 环境变量

复制 `.env.example` 后按需填写本地配置。

```bash
cp .env.example .env
```

常用变量：

```text
AI_API_KEY=local-omlx-dev-key
AI_PROVIDER_API=responses
OPENAI_API_BASE_URL=http://127.0.0.1:8001/v1
OPENAI_MODEL=Qwen3.6-35B-A3B-4.4bit-msq
AI_REQUEST_TIMEOUT_SECONDS=180
AI_MAX_TOKENS=32768
AI_FAST_MAX_TOKENS=16384
AI_THINKING_MAX_TOKENS=32768
BACKEND_HOST=127.0.0.1
BACKEND_PORT=8000
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
WORD_ADDIN_API_BASE_URL=http://127.0.0.1:8000
```

当前前端 MVP 在开发环境中先请求同源 `/api/sessions` 创建本地 session。Responses 模式优先请求同源 `/api/proofread/stream`；Chat 模式请求同源 `/api/proofread`。这些请求由 `https://localhost:3000` 的 Webpack dev server 代理到 `http://127.0.0.1:8000`，避免 Word 任务窗格从 HTTPS 页面直接请求 HTTP 后端时被 WebView 拦截。Responses 流式读取不可用时，插件会自动回退到同源 `/api/proofread`。

API Key 只配置在后端运行环境中。本地开发使用根目录 `.env`；生产环境使用部署平台提供的 Secret 或 Environment Variables。不要把真实 Key 写入 `manifest.xml`、`taskpane.ts`、Webpack 配置、前端构建产物或文档。`local-omlx-dev-key` 只用于本机 oMLX 开发服务鉴权，不是真实云端密钥。配置真实 AI 时，后端默认使用 `/v1/responses` 单轮审校，不发送 `previous_response_id`；插件也可以切换到 `/v1/chat/completions` 单轮审校。快速审校使用 `AI_FAST_MAX_TOKENS`，深度审校使用 `AI_THINKING_MAX_TOKENS`。

后端默认 `BACKEND_LOG_LEVEL=INFO`，会打印请求模式、文本长度、provider 状态码、问题数和定位数量等调试信息，不打印 API Key 或选中文本全文。需要更细定位过程时可临时设为 `DEBUG`。

## 启动 oMLX 本地 AI 服务

本地真实 AI 联调推荐先用 oMLX 启动 OpenAI 兼容服务。脚本默认读取 `/Users/wulala/AI/models`，监听 `8001` 端口，避免和 FastAPI 后端 `8000` 冲突。

```bash
./scripts/start-omlx.sh
```

脚本默认会在启动前把 `Qwen3.6-35B-A3B-4.4bit-msq` 写入 `~/.omlx/model_settings.json`，设置为 default + pinned。oMLX 启动时会预加载 pinned 模型，因此正常情况下不需要再到管理页手动加载。

可按需覆盖默认配置：

```bash
OMLX_MODEL_DIR=/Users/wulala/AI/models \
OMLX_BASE_PATH="$HOME/.omlx" \
OMLX_PORT=8001 \
OMLX_API_KEY=local-omlx-dev-key \
OMLX_PRELOAD_MODEL=Qwen3.6-35B-A3B-4.4bit-msq \
./scripts/start-omlx.sh
```

如只想启动服务、不改 oMLX 的模型设置，可使用：

```bash
OMLX_CONFIGURE_MODEL_SETTINGS=0 ./scripts/start-omlx.sh
```

模型服务检查：

```bash
curl --noproxy 127.0.0.1 http://127.0.0.1:8001/v1/models \
  -H 'Authorization: Bearer local-omlx-dev-key'
```

默认后端联调模型使用 `Qwen3.6-35B-A3B-4.4bit-msq`。如果 `/v1/models` 返回的模型 ID 和目录名不同，请把 `.env` 中的 `OPENAI_MODEL` 改成返回的模型 ID。

## 启动后端

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000 --reload
```

健康检查：

```bash
curl --noproxy 127.0.0.1 http://127.0.0.1:8000/health
```

预期返回：

```json
{"status":"ok"}
```

审校接口 smoke test。未配置 `AI_API_KEY` 时返回 mock；使用上面的 oMLX 配置时调用本地真实 AI：

```bash
SESSION_ID="$(curl --noproxy 127.0.0.1 -sS -X POST http://127.0.0.1:8000/api/sessions | python3 -c 'import json,sys; print(json.load(sys.stdin)["session_id"])')"
curl --noproxy 127.0.0.1 -X POST http://127.0.0.1:8000/api/proofread \
  -H 'Content-Type: application/json' \
  -d "{\"text\":\"这是一段需要审校的文本。\",\"session_id\":\"${SESSION_ID}\",\"provider_api\":\"responses\",\"proofread_mode\":\"fast\",\"context\":{\"source\":\"manual-curl\"}}"
```

流式审校接口 smoke test。预期会依次看到 `status`、`result` 等 SSE 事件：

```bash
curl --no-buffer --noproxy 127.0.0.1 -X POST http://127.0.0.1:8000/api/proofread/stream \
  -H 'Content-Type: application/json' \
  -H 'Accept: text/event-stream' \
  -d "{\"text\":\"这是一段需要审校的文本。\",\"session_id\":\"${SESSION_ID}\",\"provider_api\":\"responses\",\"proofread_mode\":\"fast\",\"context\":{\"source\":\"manual-curl\"}}"
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

1. 启动 oMLX 本地 AI 服务，确认 `http://127.0.0.1:8001/v1/models` 可访问。
2. 启动后端，确认 `/health` 返回 `{"status":"ok"}`。
3. 启动 `word-addin` dev server。
4. 运行 `npm run start` 旁加载插件到 Word。
5. 在 Word 文档中选中一段文本。
6. 打开任务窗格，按需选择“快速审校/深度审校”、`Responses/Chat` 和“批注模式/修订模式”。
7. 点击“AI 审校”。
8. 确认任务窗格“运行过程”区域先逐条显示阶段进度，再显示审校结果。
9. 如果选择批注模式且后端返回非空 `issues[]`，确认可定位问题批注在对应原文片段上；不可定位问题会作为 fallback 汇总批注插在当前选区。
10. 如果选择修订模式，确认有 `replacement` 且可定位的问题会在 Word 审阅面板中显示为可接受/拒绝的修订，运行后原修订跟踪设置会恢复。
11. 如果后端返回空 `issues[]`，确认任务窗格提示未发现明显问题，且 Word 中不新增批注或修订。
12. 点击“新建对话”，确认任务窗格清空当前结果，并创建新的本地 session。
13. 审校运行中点击“停止审校”，确认请求停止、不会插入批注或修订，并在历史记录中保存为已停止。
14. 使用历史记录“清空、另存为、导入”，确认本地历史可管理。

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

## 文档维护约定

开发完成后需要同步更新文档：

- 改接口契约时，更新 `spec.md` 和本文件的 API 说明。
- 改启动命令、端口、环境变量时，更新本文件的运行说明。
- 改 Word 插件用户流程时，更新联调流程和常见问题。
- 改项目协作规则时，更新 `AGENTS.md`。
