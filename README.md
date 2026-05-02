# Word AI 审校助手

面向出版社责任编辑的 Word AI 审校助手。编辑可以在 Word 中选中一段正文做局部审校，也可以选择“全书正文”并上传 `.docx` 文件，由 FastAPI 后端完成全书抽取、分块审校、批注/修订写回，并生成新的 Word 文件。

## 项目结构

```text
.
├── backend/      # Python FastAPI 后端服务
├── docs/         # 架构和补充文档
├── scripts/      # 本地开发辅助脚本
├── word-addin/   # Office.js + TypeScript + Webpack Word 插件
├── AGENTS.md     # 协作约定
└── spec.md       # API 契约和验收标准
```

## 当前能力

- 审校范围：当前选区或全书正文。当前选区由插件读取并回写；全书正文上传 `.docx`，后端处理目录可见文本、正文、表格和常见文本框文字。本版不支持 `.doc`，请先另存为 `.docx`。
- 分块规则：当前选区 `> 7000` 字时走后端内存异步任务；全书 `.docx` 先按章拆分，再按节拆分；仍超过 7000 字时再用可提取的目录小标题辅助拆分，最后按现有段落/句末规则切分。
- AI API：支持 OpenAI 兼容 Responses API 和 Chat Completions；未配置 `AI_API_KEY` 时返回 mock 结果，方便本地联调。
- 书籍信息：插件要求填写书名，介绍可选；后端把书籍信息作为 prompt 背景，但只审校传入正文。
- 结果处理：后端把 AI 输出转换为结构化 `issues[]`，过滤纯空白差异，并按 `original` 计算 `start/end/locator`。
- Word 写回：当前选区审校完成后先展示结果，编辑筛选、勾选并确认后由插件写回批注或修订；全书 `.docx` 由后端直接生成带批注或修订的新文件，插件展示新文件名和下载入口。
- 历史记录：插件在本地保存最近 20 条新 schema 历史，支持清空、导出 JSON、导入 JSON；历史不保存完整正文。

完整 API 契约见 [spec.md](spec.md)，通讯链路见 [docs/architecture.md](docs/architecture.md)。

## 环境变量

复制模板后按需填写：

```bash
cp .env.example .env
```

常用配置：

```text
AI_API_KEY=local-omlx-dev-key
AI_PROVIDER_API=responses
OPENAI_API_BASE_URL=http://127.0.0.1:8001/v1
OPENAI_MODEL=Qwen3.6-35B-A3B-4.4bit-msq
AI_REQUEST_TIMEOUT_SECONDS=180
AI_MAX_TOKENS=32768
AI_FAST_MAX_TOKENS=8192
AI_THINKING_MAX_TOKENS=16384
BACKEND_HOST=127.0.0.1
BACKEND_PORT=8000
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
WORD_ADDIN_API_BASE_URL=http://127.0.0.1:8000
```

API Key 只配置在后端运行环境中。不要把真实 Key 写入 `manifest.xml`、前端源码、Webpack 配置、构建产物或文档。

## 启动 oMLX 本地 AI 服务

本地真实 AI 联调可用 oMLX 启动 OpenAI 兼容服务。脚本默认读取 `/Users/wulala/AI/models`，监听 `8001` 端口。

```bash
./scripts/start-omlx.sh
```

模型服务检查：

```bash
curl --noproxy 127.0.0.1 http://127.0.0.1:8001/v1/models \
  -H 'Authorization: Bearer local-omlx-dev-key'
```

如 `/v1/models` 返回的模型 ID 和 `.env` 不一致，请更新 `OPENAI_MODEL`。

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

## 启动 Word 插件

Windows 编辑电脑从零安装和本地试点部署见 [DEPLOYMENT.md](DEPLOYMENT.md)。

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

## 联调流程

1. 启动 oMLX 或配置远程 OpenAI 兼容 API。
2. 启动后端，确认 `/health` 返回 `{"status":"ok"}`。
3. 启动 `word-addin` dev server。
4. 运行 `npm run start` 旁加载插件到 Word。
5. 当前选区审校：在 Word 文档中选中正文；全书审校：准备一个 `.docx` 文件。
6. 打开任务窗格，填写书名，按需选择审校范围、审校模式、API 模式和应用方式；全书模式需选择 `.docx` 文件。
7. 点击“AI 审校”。当前选区会先展示问题；全书 `.docx` 会展示分块进度并在完成后显示新文件名。
8. 当前选区可筛选、勾选、定位并点击“应用 N 条到 Word”；全书 `.docx` 点击“下载审校后 Word”获取后端生成的新文件。
9. 验证停止审校、重试当前分块、重试失败分块、历史导出和历史导入等常用流程。

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
```

Manifest 联网校验：

```bash
cd word-addin
npm run validate
```

`npm run validate` 需要访问 Microsoft Office manifest validation service，离线或网络受限时可能失败。

## 文档维护约定

- 改接口契约时，更新 [spec.md](spec.md)。
- 改通讯链路、SSE、Word 写回流程时，更新 [docs/architecture.md](docs/architecture.md)。
- 改启动方式、端口、环境变量或联调流程时，更新本文件。
- 改 Windows 试点部署流程时，更新 [DEPLOYMENT.md](DEPLOYMENT.md)。
- 临时排障记录放在 `docs/tmp/`，不要写进长期 README。
