# Word AI 审校助手 Windows 本地试点部署手册

本文面向技术支持人员，用于在 5 位编辑的 Windows 电脑上从零部署本地试点版。每台编辑电脑都会启动本项目的 FastAPI 后端和 Word 插件开发服务，Word 通过旁加载插件清单使用“AI 审校”按钮；AI 能力由远程 OpenAI 兼容 API 提供。

## 1. 适用范围

- 操作系统：Windows 10 或 Windows 11。
- Office：Microsoft Word 桌面版。
- 部署方式：每台编辑电脑本地运行 `backend` 和 `word-addin`。
- 插件地址：`https://localhost:3000/taskpane.html`。
- 后端地址：`http://127.0.0.1:8000`。
- AI 服务：远程 OpenAI 兼容接口，由 `.env` 中的 `OPENAI_API_BASE_URL`、`OPENAI_MODEL` 和 `AI_API_KEY` 配置。

本试点默认不要求每台电脑安装 oMLX 或本地大模型。如果后续改为本地模型试点，需要另行补充模型安装、硬件要求和模型服务启动步骤。

## 2. 部署架构

```text
Word 桌面版
  -> 旁加载 word-addin/manifest.xml
  -> 打开 https://localhost:3000/taskpane.html
  -> Webpack dev server 代理 /api 到 http://127.0.0.1:8000
  -> FastAPI backend 调用远程 OpenAI 兼容 API
```

注意：

- API Key 只写入每台电脑本地根目录 `.env`。
- 不要把真实 API Key 写入 `manifest.xml`、前端源码、Webpack 配置、截图、聊天记录或文档。
- 真实稿件试点时建议先复制 Word 文档副本，默认使用“批注模式”。

## 3. 前置软件

在每台编辑电脑上安装：

- Microsoft Word 桌面版。
- Git for Windows。
- Python 3.11 或 Python 3.12，安装时勾选“Add python.exe to PATH”。
- Node.js LTS，安装时允许加入 PATH。
- PowerShell。Windows 自带 PowerShell 可用。
- 能访问远程 AI API 地址的网络。

安装完成后，重新打开 PowerShell，执行以下命令确认可用：

```powershell
git --version
python --version
node --version
npm --version
```

如果提示找不到命令，先重启 PowerShell；仍失败时，检查 Git、Python、Node.js 是否已加入系统 PATH。

## 4. 获取代码

建议把项目放在固定目录，例如 `C:\Pilot\word-ai-proofreader`。

方式一：从 Git 仓库克隆：

```powershell
mkdir C:\Pilot
cd C:\Pilot
git clone <repo-url> word-ai-proofreader
cd C:\Pilot\word-ai-proofreader
```

方式二：由技术支持复制完整项目目录到每台电脑，然后进入项目根目录：

```powershell
cd C:\Pilot\word-ai-proofreader
```

确认根目录能看到这些文件和目录：

```powershell
dir
```

预期至少包含：

```text
backend
word-addin
README.md
spec.md
.env.example
```

## 5. 配置本机 `.env`

在项目根目录复制环境变量模板：

```powershell
Copy-Item .env.example .env
notepad .env
```

把 `.env` 调整为试点配置。以下示例中的 `AI_API_KEY`、`OPENAI_API_BASE_URL` 和 `OPENAI_MODEL` 必须替换为试点实际值：

```text
AI_API_KEY=请替换为试点专用Key
AI_PROVIDER_API=responses
OPENAI_API_BASE_URL=https://请替换为远程兼容API地址/v1
OPENAI_MODEL=请替换为试点模型ID
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

配置要求：

- `AI_API_KEY` 使用试点专用 Key，不使用个人生产 Key。
- `OPENAI_API_BASE_URL` 应指向远程 OpenAI 兼容服务的 `/v1` 根路径。
- `OPENAI_MODEL` 必须与远程服务提供的模型 ID 一致。
- `BACKEND_CORS_ORIGINS` 保持 `https://localhost:3000,http://localhost:3000`。
- 不要提交、转发或截图包含真实 Key 的 `.env`。

## 6. 启动后端

打开第一个 PowerShell 窗口：

```powershell
cd C:\Pilot\word-ai-proofreader\backend
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
uvicorn app.main:app --env-file ..\.env --host 127.0.0.1 --port 8000 --reload
```

如果激活虚拟环境时报执行策略错误，执行一次：

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

然后重新执行：

```powershell
.\.venv\Scripts\Activate.ps1
```

保持这个 PowerShell 窗口打开。后端正常启动后，会持续显示 `Uvicorn running on http://127.0.0.1:8000`。

另开一个 PowerShell 窗口检查健康状态：

```powershell
curl.exe http://127.0.0.1:8000/health
```

预期返回：

```json
{"status":"ok"}
```

## 7. 启动 Word 插件服务

打开第二个 PowerShell 窗口：

```powershell
cd C:\Pilot\word-ai-proofreader\word-addin
npm install
npx office-addin-dev-certs install
npm run dev-server
```

安装证书时如果 Windows 弹出安全确认，请选择信任。保持这个 PowerShell 窗口打开。

另开一个 PowerShell 窗口检查插件页面：

```powershell
curl.exe -k -I https://localhost:3000/taskpane.html
```

预期能看到 HTTP 状态 `200`。如果端口尚未完全启动，等待几秒后重试。

## 8. 旁加载 Word 插件

打开第三个 PowerShell 窗口：

```powershell
cd C:\Pilot\word-ai-proofreader\word-addin
npm run start
```

命令会使用 `word-addin\manifest.xml` 旁加载插件并打开 Word。确认 Word 功能区出现“AI 审校”按钮。

如果 Word 已经打开但没有出现按钮：

1. 关闭所有 Word 窗口。
2. 在第三个 PowerShell 窗口执行：

   ```powershell
   npm run stop
   npm run start
   ```

3. 重新检查 Word 功能区。

## 9. 验证审校流程

准备一个 Word 文档副本，按以下步骤验证：

1. 打开 Word 文档，选中一段 500 到 3000 字的正文。
2. 点击功能区“AI 审校”。
3. 在任务窗格填写“书名”，书籍介绍可选。
4. 审校范围选择“当前选区”。
5. 审校模式先选择“快速审校”。
6. API 模式先选择 `Responses`。
7. 应用方式先选择“批注模式”。
8. 点击“AI 审校”。
9. 审校完成后，确认任务窗格先展示结果，不会自动写入 Word。
10. 点击单条问题的“定位”，确认 Word 能选中对应原文。
11. 勾选需要应用的问题，点击“应用 N 条到 Word”。
12. 确认 Word 只对已勾选问题插入批注。

验收通过标准：

- 任务窗格能打开。
- 后端 `/health` 返回 `{"status":"ok"}`。
- 当前选区审校能返回结果或明确提示未发现明显问题。
- 可定位问题点击“定位”后能选中原文。
- 批注模式不会直接改正文。
- 点击“应用到 Word”后只写入已勾选问题。

## 10. 日常启动和停止

每天试点开始时：

1. 启动后端：

   ```powershell
   cd C:\Pilot\word-ai-proofreader\backend
   .\.venv\Scripts\Activate.ps1
   uvicorn app.main:app --env-file ..\.env --host 127.0.0.1 --port 8000 --reload
   ```

2. 启动插件服务：

   ```powershell
   cd C:\Pilot\word-ai-proofreader\word-addin
   npm run dev-server
   ```

3. 旁加载 Word 插件：

   ```powershell
   cd C:\Pilot\word-ai-proofreader\word-addin
   npm run start
   ```

每天试点结束时：

- 在后端和 dev server 的 PowerShell 窗口按 `Ctrl+C` 停止服务。
- 如需停止 Word 调试旁加载，执行：

  ```powershell
  cd C:\Pilot\word-ai-proofreader\word-addin
  npm run stop
  ```

## 11. 常见问题

### 11.1 端口 8000 或 3000 被占用

检查占用进程：

```powershell
netstat -ano | findstr :8000
netstat -ano | findstr :3000
```

如果确认是旧的本项目进程，可在任务管理器中结束对应 PID。不要随意结束不认识的系统或办公软件进程。

### 11.2 PowerShell 不允许激活虚拟环境

错误通常类似“running scripts is disabled on this system”。执行：

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

然后重新打开 PowerShell 或重新激活：

```powershell
.\.venv\Scripts\Activate.ps1
```

### 11.3 Python、Node 或 npm 找不到

先重新打开 PowerShell，再执行：

```powershell
python --version
node --version
npm --version
```

仍失败时，重新安装对应软件，并确认安装时加入 PATH。

### 11.4 证书不受信任或任务窗格打不开

在 `word-addin` 目录重新安装 Office 开发证书：

```powershell
cd C:\Pilot\word-ai-proofreader\word-addin
npx office-addin-dev-certs install
```

然后重启 `npm run dev-server`，关闭并重新打开 Word。

### 11.5 Word 没有出现“AI 审校”按钮

执行：

```powershell
cd C:\Pilot\word-ai-proofreader\word-addin
npm run stop
npm run start
```

如果仍然没有出现，关闭所有 Word 窗口后重试。

### 11.6 任务窗格提示 `Load failed`

按顺序检查：

```powershell
curl.exe http://127.0.0.1:8000/health
curl.exe -k -I https://localhost:3000/taskpane.html
```

如果后端不可用，回到第 6 节重启后端。如果插件页面不可用，回到第 7 节重启 dev server。

### 11.7 后端返回 502 或审校失败

优先检查第一个 PowerShell 窗口中的后端日志。常见原因：

- `AI_API_KEY` 未替换或无权限。
- `OPENAI_API_BASE_URL` 写错。
- `OPENAI_MODEL` 与远程服务模型 ID 不一致。
- 当前网络无法访问远程 AI API。
- 远程 AI 服务返回非 JSON 或响应超时。

修正 `.env` 后，按 `Ctrl+C` 停止后端，再重新启动后端。

### 11.8 远程模型不可达

如果试点 AI 服务提供 `/v1/models`，可以检查：

```powershell
curl.exe https://请替换为远程兼容API地址/v1/models -H "Authorization: Bearer 请替换为试点专用Key"
```

如果该命令失败，先确认网络、代理、API 地址和 Key；不要把包含真实 Key 的命令截图发到群聊。

## 12. 安全和试点约束

- 每台电脑的 `.env` 只保存在本机，不提交到 Git。
- 不把真实 API Key 写入前端文件、清单文件、构建产物或文档。
- 不在群聊、截图、录屏和反馈表中暴露 API Key。
- 真实未出版稿件先使用副本试点。
- 默认使用“批注模式”，不要默认用“修订模式”直接替换正文。
- 审校结果必须由编辑人工确认，插件只作为辅助工具。
- 后端日志默认使用 `INFO`，不要在真实稿件试点中长期使用 `DEBUG`。

## 13. 试点反馈记录建议

每位编辑每天记录 3 到 5 次真实使用反馈即可。建议记录：

- 电脑编号或编辑编号。
- 文档类型和大致字数。
- 审校范围：当前选区或全书正文。
- 审校模式：快速审校或深度审校。
- API 模式：Responses 或 Chat。
- 是否成功返回结果。
- 定位是否准确。
- 批注是否写入正确位置。
- 有价值建议数量。
- 明显误报数量。
- 错误提示或截图，截图前遮挡敏感正文和 API Key。
