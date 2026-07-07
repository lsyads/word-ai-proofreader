# Word AI Proofreader Windows Local Pilot Deployment Guide

Language: English | [简体中文](DEPLOYMENT_cn.md)

This guide is for technical support staff deploying a local pilot from scratch on five Windows editor machines. Each editor machine runs the FastAPI backend and the Word add-in development service locally. Word opens the DocPilot V2 publishing review workbench by sideloading the add-in manifest. AI capability is provided by a remote OpenAI-compatible API.

Related docs: project entry point [README.md](README.md), code structure [ARCHITECTURE.md](ARCHITECTURE.md), test matrix [TESTING.md](TESTING.md), and API contract [spec.md](spec.md). This document only maintains the Windows local pilot deployment flow.

## 1. Scope

- OS: Windows 10 or Windows 11.
- Office: Microsoft Word desktop.
- Deployment style: each editor machine runs `backend` and `word-addin` locally.
- Add-in URL: `https://localhost:3000/taskpane.html`.
- Backend URL: `http://127.0.0.1:8000`.
- AI service: remote OpenAI-compatible endpoint. For the pilot, leave `AI_PROFILES_JSON` empty by default and let `.env` values `OPENAI_API_BASE_URL`, `OPENAI_MODEL`, and `AI_API_KEY` create the `Default AI (.env)` profile.

This pilot does not require every machine to install oMLX or a local model by default. If a later pilot uses local models, add separate model installation, hardware, and model-service startup instructions.

## 2. Deployment Architecture

```text
Word desktop
  -> sideload word-addin/manifest.xml
  -> open https://localhost:3000/taskpane.html
  -> Webpack dev server proxies /api to http://127.0.0.1:8000
  -> FastAPI backend calls a remote OpenAI-compatible API
```

Notes:

- API keys are written only to the local root `.env` on each machine.
- Do not put real API keys in `manifest.xml`, frontend source, Webpack config, screenshots, chat logs, or documentation.
- For real manuscript pilots, copy the Word document first. If the pilot should be conservative, choose comment mode in advanced writeback settings before writing back.

## 3. Prerequisites

Install on each editor machine:

- Microsoft Word desktop.
- Git for Windows.
- Python 3.11 or Python 3.12, with "Add python.exe to PATH" checked.
- Node.js 22.15.0 or newer, with PATH enabled.
- PowerShell, which is built into Windows.
- Network access to the remote AI API endpoint.

After installation, reopen PowerShell and check:

```powershell
git --version
python --version
node --version
npm --version
```

If a command is not found, restart PowerShell first. If it still fails, check whether Git, Python, and Node.js were added to the system PATH.

## 4. Get the Code

Use a stable directory such as `C:\Pilot\word-ai-proofreader`.

Option 1, clone from Git:

```powershell
mkdir C:\Pilot
cd C:\Pilot
git clone https://github.com/lsyads/word-ai-proofreader.git word-ai-proofreader
cd C:\Pilot\word-ai-proofreader
```

Option 2, technical support copies the full project directory to each machine, then enters the project root:

```powershell
cd C:\Pilot\word-ai-proofreader
```

Confirm the root directory contains at least:

```text
backend
word-addin
README.md
spec.md
.env.example
```

## 5. Configure Local `.env`

Copy the environment template in the project root:

```powershell
Copy-Item .env.example .env
notepad .env
```

Adjust `.env` for the pilot. `.env.example` contains a multi-profile `AI_PROFILES_JSON` template, including a local `local-omlx` profile that reads `AI_API_KEY` and remote profiles that read `OPENROUTER_API_KEY` or `MIMO_API_KEY`. The backend reads keys according to each profile's `api_key_env`, and the add-in prefers a profile whose id, model, or label contains `mimo-v2.5-pro`. In the current template that profile is `mimo-v2.5-pro`, with key from `MIMO_API_KEY`. If the Windows remote pilot uses one remote OpenAI-compatible service, set `AI_PROFILES_JSON=` and use the legacy `.env` profile below.

Replace `AI_API_KEY`, `OPENAI_API_BASE_URL`, and `OPENAI_MODEL` with pilot values:

```text
AI_API_KEY=
AI_PROVIDER_API=responses
AI_PROFILES_JSON=
OPENAI_API_BASE_URL=https://pilot-ai.example.com/v1
OPENAI_MODEL=pilot-model-id
AI_REQUEST_TIMEOUT_SECONDS=180
AI_FAST_MAX_TOKENS=16384
AI_THINKING_MAX_TOKENS=32768
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
```

Configuration requirements:

- Use a pilot-specific `AI_API_KEY`, not a personal production key.
- `OPENAI_API_BASE_URL` should point to the remote OpenAI-compatible `/v1` root.
- `OPENAI_MODEL` must match a model id exposed by the remote service.
- When `AI_PROFILES_JSON` is empty, the backend creates `Default AI (.env)` from `AI_API_KEY/AI_PROVIDER_API/OPENAI_API_BASE_URL/OPENAI_MODEL`.
- If multiple profiles are retained, configure the `api_key_env` required by the profile the add-in will actually select. In the current template, `mimo-v2.5-pro` reads from `MIMO_API_KEY`.
- Keep `BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000`.
- `BACKEND_HOST`, `BACKEND_PORT`, and `WORD_ADDIN_API_BASE_URL` are intentionally omitted from `.env.example` because the current code does not read them. The backend address is controlled by the uvicorn command in section 6; the add-in proxy address is controlled by `word-addin\webpack.config.js`.
- Do not commit, forward, or screenshot a `.env` that contains a real key.

## 6. Start the Backend

Open the first PowerShell window:

```powershell
cd C:\Pilot\word-ai-proofreader\backend
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
uvicorn app.main:app --env-file ..\.env --host 127.0.0.1 --port 8000 --reload
```

If virtual-environment activation fails because script execution is disabled, run once:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Then activate again:

```powershell
.\.venv\Scripts\Activate.ps1
```

Keep this PowerShell window open. A healthy backend keeps printing `Uvicorn running on http://127.0.0.1:8000`.

Open another PowerShell window to check health:

```powershell
curl.exe http://127.0.0.1:8000/health
```

Expected response:

```json
{"status":"ok"}
```

## 7. Start the Word Add-in Service

Open the second PowerShell window:

```powershell
cd C:\Pilot\word-ai-proofreader\word-addin
npm install
npx office-addin-dev-certs install
npm run dev-server
```

If Windows prompts to trust the certificate, accept it. Keep this PowerShell window open.

Open another PowerShell window to check the add-in page:

```powershell
curl.exe -k -I https://localhost:3000/taskpane.html
```

Expected result: HTTP status `200`. If the port is still starting, wait a few seconds and retry.

## 8. Sideload the Word Add-in

Open the third PowerShell window:

```powershell
cd C:\Pilot\word-ai-proofreader\word-addin
npm run start
```

The command sideloads `word-addin\manifest.xml` and opens Word. Confirm that the Word ribbon shows the Chinese `AI 审校` button.

If Word is open but the button does not appear:

1. Close all Word windows.
2. In the third PowerShell window, run:

   ```powershell
   npm run stop
   npm run start
   ```

3. Check the Word ribbon again.

## 9. Validate the Review Flow

Prepare a copy of a Word document and validate:

1. Open the Word document and select 500 to 3000 characters of body text.
2. Click the ribbon `AI 审校` button to open the task pane.
3. Choose `当前选区` and fill in `书名`. Book background and review focus can be filled in under `补充信息（可选）`.
4. For a conservative pilot, expand `更多设置（试点支持）` and change `写回模式` to `批注模式`.
5. Click `开始审校`.
6. After review, confirm the task pane shows review suggestions. If there are no suggestions, it should show `审校完成，未发现需要处理的建议。` or an equivalent no-suggestions message.
7. Click `定位原文` for one suggestion and confirm Word selects the corresponding text.
8. Accept or reject suggestions one by one, or use `接受全部待处理` / `忽略全部待处理` and confirm the dialog explains its scope.
9. Click `写回 N 条已接受建议` and confirm Word writes only accepted suggestions.
10. Switch to `全书 DOCX`, choose a `.docx` file, repeat review, accept suggestions, write back, and click `下载审校后文件`.

Acceptance checks:

- The task pane opens.
- Backend `/health` returns `{"status":"ok"}`.
- Current-selection or DOCX review returns suggestions, or clearly reports no suggestions.
- `定位原文` selects the original text for locatable suggestions.
- `批注模式` does not directly change body text.
- `写回 N 条已接受建议` writes only accepted suggestions.

## 10. Daily Start and Stop

At the start of each pilot day:

1. Start the backend:

   ```powershell
   cd C:\Pilot\word-ai-proofreader\backend
   .\.venv\Scripts\Activate.ps1
   uvicorn app.main:app --env-file ..\.env --host 127.0.0.1 --port 8000 --reload
   ```

2. Start the add-in service:

   ```powershell
   cd C:\Pilot\word-ai-proofreader\word-addin
   npm run dev-server
   ```

3. Sideload the Word add-in:

   ```powershell
   cd C:\Pilot\word-ai-proofreader\word-addin
   npm run start
   ```

At the end of each pilot day:

- Press `Ctrl+C` in the backend and dev-server PowerShell windows.
- To stop Word debugging sideload state:

  ```powershell
  cd C:\Pilot\word-ai-proofreader\word-addin
  npm run stop
  ```

## 11. Common Problems

### 11.1 Port 8000 or 3000 Is Already Used

Check the process:

```powershell
netstat -ano | findstr :8000
netstat -ano | findstr :3000
```

If the process is an old instance of this project, end that PID in Task Manager. Do not kill unknown system or office-software processes casually.

### 11.2 PowerShell Blocks Virtual Environment Activation

The error usually says scripts are disabled. Run:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Then reopen PowerShell or activate again:

```powershell
.\.venv\Scripts\Activate.ps1
```

### 11.3 Python, Node, or npm Is Not Found

Reopen PowerShell and run:

```powershell
python --version
node --version
npm --version
```

If the commands still fail, reinstall the relevant software and confirm PATH is enabled during installation.

### 11.4 Certificate Is Not Trusted or the Task Pane Does Not Open

Reinstall Office development certificates from `word-addin`:

```powershell
cd C:\Pilot\word-ai-proofreader\word-addin
npx office-addin-dev-certs install
```

Restart `npm run dev-server`, close Word, and reopen it.

### 11.5 Word Does Not Show the `AI 审校` Button

Run:

```powershell
cd C:\Pilot\word-ai-proofreader\word-addin
npm run stop
npm run start
```

If it still does not appear, close all Word windows and retry.

### 11.6 The Task Pane Shows `Load failed`

Check in order:

```powershell
curl.exe http://127.0.0.1:8000/health
curl.exe -k -I https://localhost:3000/taskpane.html
```

If the backend is unavailable, restart it from section 6. If the add-in page is unavailable, restart the dev server from section 7.

### 11.7 Backend Returns 502 or Review Fails

Check backend logs in the first PowerShell window first. Common causes:

- `AI_API_KEY` was not replaced or lacks permission.
- `OPENAI_API_BASE_URL` is wrong.
- `OPENAI_MODEL` does not match the remote service model id.
- `AI_PROFILES_JSON` was not cleared and the add-in selected `mimo-v2.5-pro`, but `MIMO_API_KEY` or another required `api_key_env` is empty.
- The network cannot reach the remote AI API.
- The remote AI service returned non-JSON or timed out.

After fixing `.env`, press `Ctrl+C` to stop the backend and start it again.

### 11.8 Remote Model Is Unreachable

If the pilot AI service exposes `/v1/models`, check:

```powershell
$env:PILOT_AI_API_KEY = "paste-pilot-key-here"
curl.exe https://pilot-ai.example.com/v1/models -H "Authorization: Bearer $env:PILOT_AI_API_KEY"
```

If it fails, check network, proxy, API address, and key. Do not post screenshots containing real keys in group chats.

## 12. Security and Pilot Constraints

- Each machine's `.env` stays local and is not committed to Git.
- Do not write real API keys into frontend files, manifest files, build artifacts, or documentation.
- Do not expose API keys in group chats, screenshots, recordings, or feedback forms.
- Use copies of real unpublished manuscripts for pilots.
- For conservative pilots, use comment mode. Before using revisions plus comments, confirm the editor understands Word revisions can be accepted or rejected.
- Review results must be confirmed by an editor. The add-in is an assistive tool.
- Backend logs default to `INFO`. Do not keep `DEBUG` enabled for real manuscript pilots.

## 13. Pilot Feedback Notes

Each editor can record 3 to 5 real-use feedback entries per day. Suggested fields:

- Machine id or editor id.
- Document type and approximate word count.
- Review source: current selection or whole-book DOCX.
- Review mode: fast or deep; the default is deep review.
- API mode: Responses or Chat.
- Whether results were returned successfully.
- Whether location was accurate.
- Whether comments or revisions were written to the correct position.
- Number of valuable suggestions.
- Number of clear false positives.
- Error messages or screenshots, with sensitive manuscript text and API keys covered before sharing.
