# Backend README

`backend/` 是 Word AI 审校助手的 FastAPI 服务。它接收 Word 插件传来的当前选区文本，或接收全书 `.docx` 文件；调用 OpenAI 兼容 AI provider 或 mock fallback，返回结构化审校问题，并负责文本定位、DOCX 批注/修订写回和结果文件保存。

完整 API 契约见仓库根目录 [spec.md](../spec.md)；通讯链路见 [docs/architecture.md](../docs/architecture.md)。

## 目录结构

```text
backend/
├── app/
│   ├── main.py          # FastAPI 入口、路由、SSE 响应
│   ├── schemas.py       # Pydantic 请求/响应模型
│   ├── settings.py      # 环境变量配置
│   └── services/
│       ├── ai_client.py # Responses / Chat Completions client
│       ├── chunking.py  # 分块、聚合、全局位置转换
│       ├── docx.py      # DOCX 抽取、分块、OOXML 批注/修订写回
│       ├── docx_store.py # DOCX 结果文件稳定存储和过期清理
│       ├── docx_tasks.py # DOCX 全书审校任务和下载文件管理
│       ├── proofread.py # 审校编排、mock fallback、定位
│       ├── sessions.py  # 本地 session ID
│       └── tasks.py     # 内存异步分块任务
├── tests/
└── requirements.txt
```

## 主要职责

- 提供 `GET /health`、审校接口、本地 session 接口、同步分块调试接口、异步分块任务接口和 DOCX 全书审校任务接口。
- 校验请求：`text` 非空，`book.title` 必填且非空。
- 在未配置 `AI_API_KEY` 时返回 mock issue，保证本地可联调。
- 在配置 `AI_API_KEY` 时按 `provider_api` 调用 OpenAI 兼容 Responses API 或 Chat Completions。
- 把 AI 精简输出转换为结构化 `issues[]`，过滤纯空白差异，计算 `start/end/locator`。
- 当前选区 `> 7000` 字时按默认 `chunk_size=5000` 分块；分块结果额外返回 `global_start/global_end`。
- 全书 `.docx` 任务仅支持 `.docx`，不支持旧二进制 `.doc`；后端抽取目录可见文本、正文、表格和常见文本框文字，先按章、节拆分，仍超过 7000 字时再按可提取的目录小标题辅助拆分，最后生成新的 `.docx` 结果文件。
- DOCX 结果文件保存到 `DOCX_OUTPUT_DIR`，SQLite 索引记录输出文件、任务统计和过期时间；`DOCX_RETENTION_DAYS` 默认并强制最少为 7 天，后端重启后未过期结果仍可下载。
- 通过内存任务提供分块进度、SSE、取消、当前分块重试和失败分块重试。

## 本地运行

从仓库根目录复制环境变量模板：

```bash
cp .env.example .env
```

启动后端：

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

## 常用环境变量

```text
AI_API_KEY=local-omlx-dev-key
AI_PROVIDER_API=responses
OPENAI_API_BASE_URL=http://127.0.0.1:8001/v1
OPENAI_MODEL=Qwen3.6-35B-A3B-4.4bit-msq
AI_REQUEST_TIMEOUT_SECONDS=180
AI_FAST_MAX_TOKENS=8192
AI_THINKING_MAX_TOKENS=16384
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
```

说明：

- `AI_API_KEY` 为空时使用 mock fallback。
- `AI_PROVIDER_API` 支持 `responses`、`chat`，也可由请求体临时覆盖。
- `local-omlx-dev-key` 只用于本机 oMLX 开发服务鉴权，不是真实云端密钥。
- 真实 API Key 不要提交到 Git，不要写入前端源码或文档。

## 测试

```bash
cd backend
source .venv/bin/activate
python -m pytest -q
```

测试覆盖重点：

- API 校验和错误映射。
- mock fallback。
- Responses / Chat 请求 payload。
- AI 输出清理、schema 校验和错误处理。
- `original` 定位、重复片段定位、`locator`、纯空白差异过滤。
- 分块规则、全局位置转换、异步任务和 SSE 事件。
- DOCX 文字抽取、章节分块、批注/修订写回、结果文件下载、稳定保留和过期清理。

## 开发注意事项

- API 契约变更必须先同步 [../spec.md](../spec.md)。
- 运行方式、端口或环境变量变更必须同步根目录 [../README.md](../README.md)。
- 临时排障记录放在 `docs/tmp/`，不要写进长期 README。
- 后端日志不得打印 API Key、Authorization header 或 Bearer token。
