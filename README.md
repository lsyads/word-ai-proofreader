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

## 当前 V2 能力

- Word 任务窗格提供一个正式入口：“AI 审校”。
- 插件打开时会创建一个本地 session；点击“清空当前结果”会清空任务窗格中的当前审校结果，并创建新的本地 session。后端不使用 session 续接 AI 上下文。
- 插件可选择审校范围：“当前选区”或“全书正文”。全书正文以 `document.body.text` 为范围，暂不包含页眉页脚、脚注、文本框等非正文内容。
- 当前选区不超过 5000 字时沿用单段审校；当前选区超过 5000 字或选择全书正文时，会自动按约 3000 字分块审校。
- 插件提供书籍信息输入：书名必填，介绍可选，并保存在本地用于同一本书连续审校；书籍信息会作为 prompt 背景传给后端。
- 插件可切换“快速审校/深度审校”、`Responses/Chat` API，并可单独开启“深度思考”；深度思考默认关闭，Chat 模式下对应后端 `reasoning.enabled`。
- 插件可切换“批注模式/修订模式”；默认批注模式，避免默认改正文。
- 插件调用后端 `POST /api/proofread`。
- Responses 模式下，插件优先调用后端 `POST /api/proofread/stream`，并在任务窗格“运行过程”区域展示阶段进度；流式不可用时自动回退 `POST /api/proofread`。Chat 模式直接调用 `POST /api/proofread`，由后端使用标准 Chat Completions。
- 分块审校使用 `POST /api/proofread/tasks` 创建内存异步任务，优先通过 `GET /api/proofread/tasks/{task_id}/events` 获取 SSE 进度；进度流不可用时回退 `GET /api/proofread/tasks/{task_id}` 轮询。
- AI 原始输出只包含精简 `issues[]`，不返回 `start/end`；其中 `replacement` 是可直接替换正文的新文本。后端会过滤纯空白差异 issue，并按 `original` 在选区文本中搜索、计算位置。
- 分块结果会把每条问题转换成全文全局位置 `global_start/global_end`；插件用全局位置计算重复原文的 occurrence，避免长文或全书中误命中。
- 审校完成后先展示结果，不立即写回 Word；任务窗格支持按严重程度、类别、定位状态和是否可直接替换筛选问题，并可逐条勾选、全选、全不选、只选高/中风险或只选可直接替换项。
- 可定位问题支持在应用前点击“定位”选中 Word 原文；点击“应用到 Word”时，插件只把已勾选的问题按当前“批注模式/修订模式”插入批注或生成修订。
- 可定位问题会逐条写回对应原文片段；修订模式下可定位且有 `replacement` 的问题生成 Word 修订，可定位但无 `replacement` 的问题回退为原位批注；已勾选但未定位的问题会在“应用到 Word”时合并为一条范围起点汇总批注，未勾选问题不会写回 Word。
- 插件支持停止当前审校，并在本地保存最近 20 条新 schema 审校历史用于回看、清空、另存为 JSON 和导入 JSON；历史记录保存已选问题 ID 和跳过数量，开发阶段不兼容旧历史数据。
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

当前前端 V2 在开发环境中先请求同源 `/api/sessions` 创建本地 session。Responses 模式优先请求同源 `/api/proofread/stream`；Chat 模式请求同源 `/api/proofread`。这些请求由 `https://localhost:3000` 的 Webpack dev server 代理到 `http://127.0.0.1:8000`，避免 Word 任务窗格从 HTTPS 页面直接请求 HTTP 后端时被 WebView 拦截。Responses 流式读取不可用时，插件会自动回退到同源 `/api/proofread`。

V2 分块审校同样走同源 `/api/proofread/tasks`、`/api/proofread/tasks/{task_id}` 和 `/api/proofread/tasks/{task_id}/events`，由 dev server 代理到 FastAPI。异步任务只保存在后端内存中，服务重启后任务状态和结果会丢失。前端收到 `chunk_started` 后会本地每秒刷新当前块耗时，后端低频 `heartbeat` 事件用于保活和校准进度；部分 chunk 失败但有结果时，任务状态为 `partial_succeeded`，后端 INFO 级别日志和前端进度区都会显示失败 chunk 的编号、范围、耗时和错误信息。

API Key 只配置在后端运行环境中。本地开发使用根目录 `.env`；生产环境使用部署平台提供的 Secret 或 Environment Variables。不要把真实 Key 写入 `manifest.xml`、`taskpane.ts`、Webpack 配置、前端构建产物或文档。`local-omlx-dev-key` 只用于本机 oMLX 开发服务鉴权，不是真实云端密钥。配置真实 AI 时，后端默认使用 `/v1/responses` 单轮审校，不发送 `previous_response_id`；插件也可以切换到 `/v1/chat/completions` 单轮审校。快速审校使用 `AI_FAST_MAX_TOKENS`，深度审校使用 `AI_THINKING_MAX_TOKENS`。

后端默认 `BACKEND_LOG_LEVEL=INFO`，会打印请求模式、文本长度、provider 状态码、AI provider 返回报文、问题数、定位数量、分块失败编号和错误原因等信息，不打印 API Key 或完整请求正文。AI 返回报文可能包含 `original` 原文摘录，便于联调定位；模型返回 Markdown 代码块、前后解释、尾随逗号或未转义控制字符时，后端会先清理再做结构校验。临时设为 `DEBUG` 时会额外打印 `/api/proofread`、`/api/proofread/stream` 和 AI provider 请求报文；DEBUG 日志会包含选区文本、书名、介绍和 AI 返回内容，但仍不会打印 API Key、`Authorization` 或 Bearer token，不建议生产开启。

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
  -d "{\"text\":\"这是一段需要审校的文本。\",\"book\":{\"title\":\"测试书名\",\"introduction\":\"这是一部用于联调的测试图书。\"},\"session_id\":\"${SESSION_ID}\",\"provider_api\":\"responses\",\"proofread_mode\":\"fast\",\"context\":{\"source\":\"manual-curl\"}}"
```

流式审校接口 smoke test。预期会依次看到 `status`、`result` 等 SSE 事件：

```bash
curl --no-buffer --noproxy 127.0.0.1 -X POST http://127.0.0.1:8000/api/proofread/stream \
  -H 'Content-Type: application/json' \
  -H 'Accept: text/event-stream' \
  -d "{\"text\":\"这是一段需要审校的文本。\",\"book\":{\"title\":\"测试书名\",\"introduction\":\"这是一部用于联调的测试图书。\"},\"session_id\":\"${SESSION_ID}\",\"provider_api\":\"responses\",\"proofread_mode\":\"fast\",\"context\":{\"source\":\"manual-curl\"}}"
```

分块任务 smoke test。创建任务后用返回的 `task_id` 查询状态，或连接任务 SSE：

```bash
TASK_ID="$(curl --noproxy 127.0.0.1 -sS -X POST http://127.0.0.1:8000/api/proofread/tasks \
  -H 'Content-Type: application/json' \
  -d "{\"text\":\"$(printf '这是一段需要分块审校的文本。%.0s' {1..400})\",\"book\":{\"title\":\"测试书名\",\"introduction\":\"这是一部用于联调的测试图书。\"},\"session_id\":\"${SESSION_ID}\",\"provider_api\":\"responses\",\"proofread_mode\":\"fast\",\"scope\":\"document\",\"chunk_size\":3000,\"context\":{\"source\":\"manual-curl\"}}" \
  | python3 -c 'import json,sys; print(json.load(sys.stdin)["task_id"])')"
curl --noproxy 127.0.0.1 http://127.0.0.1:8000/api/proofread/tasks/${TASK_ID}
curl --no-buffer --noproxy 127.0.0.1 http://127.0.0.1:8000/api/proofread/tasks/${TASK_ID}/events
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
5. 在 Word 文档中选中一段文本，或准备使用“全书正文”范围。
6. 打开任务窗格，填写书名，按需填写书籍介绍，并选择“当前选区/全书正文”、“快速审校/深度审校”、`Responses/Chat` 和“批注模式/修订模式”。
7. 点击“AI 审校”。
8. 确认任务窗格“运行过程”区域先逐条显示阶段进度；长选区或全书会显示分块进度、失败块数和累计问题数。
9. 审校完成后，确认任务窗格只展示结果，不会立即新增 Word 批注或修订。
10. 使用筛选器、复选框和批量选择按钮调整待应用问题；点击单条“定位”，确认 Word 选中对应原文片段。
11. 点击“应用 N 条到 Word”，确认只有已勾选问题写回；批注模式下可定位问题批注在对应原文片段上，修订模式下有 `replacement` 且可定位的问题会在 Word 审阅面板中显示为可接受/拒绝的修订。
12. 已勾选的未定位问题会先显示在任务窗格；点击“应用 N 条到 Word”后，统一插入一条锚定在审校范围起点第一个非空字符的汇总批注；未勾选问题不写回 Word。
13. 如果后端返回空 `issues[]`，确认任务窗格提示未发现明显问题，且 Word 中不新增批注或修订。
14. 点击“清空当前结果”，确认任务窗格清空当前结果，并创建新的本地 session。
15. 审校运行中点击“停止审校”，确认提示“当前分块完成后结束”，请求停止、不会插入批注或修订，并在历史记录中保存为已停止。
16. 使用历史记录“清空、另存为、导入”，确认当前开发版 schema 历史可管理；导入旧 schema 历史会提示格式不兼容。

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
