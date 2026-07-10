# Word AI 审校助手 Windows 本地试点部署手册

语言：[English](DEPLOYMENT.md) | 简体中文

本文面向技术支持人员，用于在 5 位编辑的 Windows 电脑上从零部署本地试点版。每台编辑电脑都会启动本项目的 FastAPI 后端和 Word 插件开发服务，Word 通过旁加载插件清单打开 DocPilot V2 出版审校工作台；AI 能力由远程 OpenAI 兼容 API 提供。

相关文档：项目入口见 [README_cn.md](README_cn.md)，代码结构见 [ARCHITECTURE_cn.md](ARCHITECTURE_cn.md)，测试矩阵见 [TESTING_cn.md](TESTING_cn.md)，API 契约见 [spec_cn.md](spec_cn.md)。本文只维护 Windows 本地试点部署流程。

## 1. 适用范围

- 操作系统：Windows 10 或 Windows 11。
- Office：Microsoft Word 桌面版。
- 部署方式：每台编辑电脑本地运行 `backend` 和 `word-addin`。
- 插件地址：`https://localhost:3000/taskpane.html`。
- 后端地址：`http://127.0.0.1:8000`。
- AI 服务：远程 OpenAI 兼容接口。默认模板优先使用 `deepseek-v4-pro` profile，并从 `DEEPSEEK_API_KEY` 读取 Key。如果试点改用单一自定义远程服务，再置空 `AI_PROFILES_JSON`，由 `.env` 中的 `OPENAI_API_BASE_URL`、`OPENAI_MODEL` 和 `AI_API_KEY` 生成 `Default AI (.env)` profile。

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
- 真实稿件试点时建议先复制 Word 文档副本；如果试点优先保守，请在任务窗格高级设置中把写回模式改为“批注模式”后再写回。

## 3. 前置软件

在每台编辑电脑上安装：

- Microsoft Word 桌面版。
- Git for Windows。
- Python 3.11 或 Python 3.12，安装时勾选“Add python.exe to PATH”。
- Node.js 22.15.0 或更新版本，安装时允许加入 PATH。
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
git clone https://github.com/lsyads/word-ai-proofreader.git word-ai-proofreader
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

把 `.env` 调整为试点配置。当前 `.env.example` 默认包含 `AI_PROFILES_JSON` 多 profile 模板，其中本地 `local-omlx` profile 读取 `AI_API_KEY`，远程 profile 读取 `OPENROUTER_API_KEY`、`MIMO_API_KEY` 或 `DEEPSEEK_API_KEY`。后端会优先按 profile 的 `api_key_env` 读取 Key，插件会优先匹配 id、model 或 label 包含 `deepseek-v4-pro` 的 profile。当前模板中匹配的是 `deepseek-v4-pro` profile，Key 来自 `DEEPSEEK_API_KEY`。

默认 DeepSeek 试点保留 `AI_PROFILES_JSON='[...]'`，只填写：

```text
DEEPSEEK_API_KEY=
```

Windows 远程试点如果改用一个自定义 OpenAI 兼容服务，建议把 `AI_PROFILES_JSON='[...]'` 整段改为 `AI_PROFILES_JSON=`，再使用下面的 legacy `.env` profile。

自定义单服务试点时，以下示例中的 `AI_API_KEY`、`OPENAI_API_BASE_URL` 和 `OPENAI_MODEL` 必须替换为试点实际值：

```text
AI_API_KEY=
AI_PROVIDER_API=responses
AI_PROFILES_JSON=
OPENAI_API_BASE_URL=https://请替换为远程兼容API地址/v1
OPENAI_MODEL=请替换为试点模型ID
AI_REQUEST_TIMEOUT_SECONDS=180
AI_FAST_MAX_TOKENS=16384
AI_THINKING_MAX_TOKENS=32768
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
```

配置要求：

- 默认 DeepSeek profile 使用试点专用的 `DEEPSEEK_API_KEY`，不要使用个人生产 Key。
- 自定义单服务 profile 使用试点专用的 `AI_API_KEY`，`OPENAI_API_BASE_URL` 指向远程 OpenAI 兼容服务的 `/v1` 根路径，`OPENAI_MODEL` 必须与远程服务提供的模型 ID 一致。
- `AI_PROFILES_JSON` 为空时，后端使用 `AI_API_KEY/AI_PROVIDER_API/OPENAI_API_BASE_URL/OPENAI_MODEL` 生成 `Default AI (.env)` profile。
- 如果试点要保留多 profile，必须为插件实际选中的 profile 配置对应的 `api_key_env`。当前模板中 `deepseek-v4-pro` 读取 `DEEPSEEK_API_KEY`。
- `BACKEND_CORS_ORIGINS` 保持 `https://localhost:3000,http://localhost:3000`。
- `BACKEND_HOST`、`BACKEND_PORT`、`WORD_ADDIN_API_BASE_URL` 当前不被代码读取，因此刻意不放进 `.env.example`；后端地址由第 6 节 uvicorn 命令决定，插件代理地址由 `word-addin\webpack.config.js` 决定。
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
2. 点击功能区“AI 审校”，打开 DocPilot V2 任务窗格。
3. 在任务窗格选择“当前选区”，填写“书名”；书籍背景和审校重点可在“补充信息”里填写。
4. 试点优先保守时，展开“更多设置（试点支持）”，把“写回模式”改为“批注模式”。
5. 点击“开始审校”。
6. 审校完成后，确认任务窗格显示审校建议；没有建议时应显示“未发现需要处理的建议”。
7. 点击单条建议的“定位原文”，确认 Word 能选中对应原文。
8. 逐条“接受”或“忽略”建议；也可用“接受全部待处理/忽略全部待处理”处理所有待处理建议，并确认弹窗说明作用范围。
9. 点击“写回已接受建议”，确认 Word 只写回已接受建议。
10. 切换到“全书 DOCX”来源时，选择 `.docx` 文件后重复审校；接受建议并点击写回后，再点击“下载审校后文件”。

验收通过标准：

- 任务窗格能打开。
- 后端 `/health` 返回 `{"status":"ok"}`。
- 当前选区或 DOCX 审校能返回审校建议，或明确提示未发现需要处理的建议。
- 可定位建议点击“定位原文”后能选中原文。
- 批注模式不会直接改正文。
- 点击“写回已接受建议”后只写入已接受建议。

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
- 未置空 `AI_PROFILES_JSON`，插件选中了 `deepseek-v4-pro`，但对应的 `DEEPSEEK_API_KEY` 或其他 `api_key_env` 为空。
- 当前网络无法访问远程 AI API。
- 远程 AI 服务返回非 JSON 或响应超时。

修正 `.env` 后，按 `Ctrl+C` 停止后端，再重新启动后端。

### 11.8 远程模型不可达

如果试点 AI 服务提供 `/v1/models`，可以检查：

```powershell
$env:PILOT_AI_API_KEY = "在这里粘贴试点专用Key"
curl.exe https://请替换为远程兼容API地址/v1/models -H "Authorization: Bearer $env:PILOT_AI_API_KEY"
```

如果该命令失败，先确认网络、代理、API 地址和 Key；不要把包含真实 Key 的命令截图发到群聊。

## 12. 安全和试点约束

- 每台电脑的 `.env` 只保存在本机，不提交到 Git。
- 不把真实 API Key 写入前端文件、清单文件、构建产物或文档。
- 不在群聊、截图、录屏和反馈表中暴露 API Key。
- 真实未出版稿件先使用副本试点。
- 试点优先保守时使用“批注模式”；使用“修订+批注”前，必须确认编辑理解 Word 修订可接受或拒绝。
- 审校结果必须由编辑人工确认，插件只作为辅助工具。
- 后端日志默认使用 `INFO`，不要在真实稿件试点中长期使用 `DEBUG`。

## 13. 试点反馈记录建议

每位编辑每天记录 3 到 5 次真实使用反馈即可。建议记录：

- 电脑编号或编辑编号。
- 文档类型和大致字数。
- 审校来源：当前选区或全书 DOCX。
- 审校模式：快速审校或深度审校，默认深度审校。
- API 模式：Responses 或 Chat。
- 是否成功返回结果。
- 定位是否准确。
- 批注或修订是否写入正确位置。
- 有价值建议数量。
- 明显误报数量。
- 错误提示或截图，截图前遮挡敏感正文和 API Key。
