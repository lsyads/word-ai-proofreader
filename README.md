# Word AI Proofreader

Language: English | [简体中文](README_cn.md)

Word AI Proofreader is a Word document review agent for publishing editors. The current main flow is the V2.2 publishing review workbench: an editor selects text in Word or uploads a `.docx` manuscript, starts a review, inspects suggestions, accepts or rejects them, and then writes approved changes back to Word.

V2.2 is project-based. A project is built around one book and stores the document map, review plan, run status, candidate suggestions, editor decisions, and report. V2 does not migrate old local project data, old traces, old task snapshots, or old result indexes. The lower-level direct APIs remain in the codebase so the validated document-processing capabilities can still be reused for development and debugging.

## Project Structure

```text
.
├── backend/          # Python FastAPI backend
├── scripts/          # Local development helper scripts
├── word-addin/       # Office.js + TypeScript + Webpack Word add-in
├── AGENTS.md         # AI coding collaboration rules
├── ARCHITECTURE.md   # Code entry points, module boundaries, and V2 data flow
├── DEPLOYMENT.md     # Windows local pilot deployment guide
├── TESTING.md        # Test commands and minimum validation matrix
├── spec.md           # API contract and acceptance criteria
└── *_cn.md           # Chinese counterparts of the public Markdown docs
```

## Current Capabilities (V2.2 Workbench)

- Review sources: current Word selection and whole-book DOCX. The add-in reads the current selection and writes it back through Office.js. Whole-book DOCX files are uploaded to the backend, which processes visible table-of-contents text, body text, tables, and common text-box content. `.doc` is not supported; save it as `.docx` first.
- Primary workflow: the task pane keeps the normal path focused on review source, book title, start review, current progress, review suggestions, writeback, and download. Recent reviews and troubleshooting details are folded by default.
- AI API: OpenAI-compatible Responses API and Chat Completions are supported. The add-in defaults to deep review, temperature `0.6`, revision-plus-comment writeback, and prefers a profile whose id, model, or label contains `mimo-v2.5-pro`. These settings live under "More settings (pilot support)". The backend direct API schema still defaults temperature to `0.2`. If no key is configured, the backend returns mock results for local integration.
- Book information: the add-in requires a book title and accepts an optional introduction. The backend uses book information and the review goal as prompt context, but it does not store the full manuscript in long-term memory.
- Publishing-editor boundary: the implemented plan shows five real execution stages: plan review, basic language review, candidate merge, evaluator recheck, and editor confirmation. Terminology, book conventions, and cross-chapter consistency are not separate V2.2 plan steps. Mechanical copyediting items such as punctuation, whitespace, full-width/half-width conversion, and Chinese/English symbol conversion belong to the copyediting vendor by default, so AI and local rules do not let those items enter the publishing-editor confirmation queue.
- Suggestions: the agent only generates candidate issues, risk notes, evidence, and recommendations. Editors accept, reject, or bulk-process pending suggestions. No suggestion is written back until an editor accepts it.
- Word writeback: accepted current-selection suggestions are located and written back by the add-in as comments or revisions plus comments. DOCX projects write back only accepted suggestions through the backend and expose a download link for the reviewed file.
- Observability: the V2 workbench stores the current review, document map, review plan, run-event trace, suggestions, project memory, and report. These technical details are folded into "Troubleshooting information (support)". Traces do not record full text, API keys, Authorization headers, or Bearer tokens.
- Result retention: V2 DOCX writeback results are stored in the project output directory under `AGENT_WORKSPACE_DIR`. The current V2 project API does not return retention days or expiration time; a download is available while the output file still exists. Lower-level DOCX task result indexes are still controlled by `DOCX_OUTPUT_DIR` and `DOCX_RETENTION_DAYS`.

Documentation:

- [AGENTS.md](AGENTS.md): AI coding collaboration rules, V2 goals, and boundaries.
- [ARCHITECTURE.md](ARCHITECTURE.md): code entry points, module boundaries, V2 workbench data flow, and common change points.
- [spec.md](spec.md): API contract, status semantics, environment variables, and acceptance criteria.
- [TESTING.md](TESTING.md): backend, frontend, manifest, and Word manual integration test matrix.
- [DEPLOYMENT.md](DEPLOYMENT.md): Windows editor-machine local pilot deployment guide.
- [SECURITY.md](SECURITY.md): vulnerability reporting, secret handling, and manuscript privacy notes.
- [CONTRIBUTING.md](CONTRIBUTING.md): contribution workflow and doc-sync rules.

## V2 Workbench

V2.2 has converged on a project-based review loop: choose the current selection or whole-book DOCX, fill in the book title, start review, inspect suggestions, accept or reject them, and write back or download the reviewed file. Current-selection location and writeback are handled by Office.js. DOCX writeback is handled by the backend. The add-in UI no longer exposes the old standalone entry points.

Product goals and AI-coding boundaries are in [AGENTS.md](AGENTS.md). Code entry points and the end-to-end flow are in [ARCHITECTURE.md](ARCHITECTURE.md). The full API contract and status semantics are in [spec.md](spec.md).

## Environment Variables

Copy the template and fill in local values:

```bash
cp .env.example .env
```

The template keeps the same shape as `.env` and includes the variables read by the backend. Common backend settings:

```text
AI_API_KEY=
OPENROUTER_API_KEY=
MIMO_API_KEY=

AI_PROVIDER_API=responses
AI_PROFILES_JSON='[...]'
OPENAI_API_BASE_URL=http://127.0.0.1:8001/v1
OPENAI_MODEL=Qwen3.6-35B-A3B-4.4bit-msq
AI_REQUEST_TIMEOUT_SECONDS=180
AI_REQUEST_TIMEOUT_MIN_SECONDS=60
AI_REQUEST_TIMEOUT_MAX_SECONDS=900
AI_REQUEST_TIMEOUT_BASE_SECONDS=60
AI_FAST_TIMEOUT_SECONDS_PER_1K_TOKENS=45
AI_THINKING_TIMEOUT_SECONDS_PER_1K_TOKENS=75
AI_FAST_MAX_TOKENS=8192
AI_THINKING_MAX_TOKENS=16384
BACKEND_LOG_LEVEL=INFO
BACKEND_CORS_ORIGINS=https://localhost:3000,http://localhost:3000
DOCX_OUTPUT_DIR=var/docx-results
DOCX_RETENTION_DAYS=7
AGENT_TRACE_DIR=var/agent-traces
AGENT_WORKSPACE_DIR=var/agent-workspace
```

API keys belong only in the backend runtime environment. Do not put real keys in `manifest.xml`, frontend source, Webpack config, build artifacts, screenshots, chat logs, or documentation.
All `*_API_KEY` entries in `.env.example` are intentionally blank; fill them only in your local `.env`.

`.env.example` currently includes `local-omlx`, `hy3-preview`, `mimo-v2.5`, and `mimo-v2.5-pro` profiles in `AI_PROFILES_JSON`. The local profile reads `AI_API_KEY` and points to `http://127.0.0.1:8001/v1`. The add-in prefers a profile whose id, model, or label contains `mimo-v2.5-pro`; the template's matching remote profile is `mimo-v2.5-pro`, with key material read from `MIMO_API_KEY`.

The backend still supports a legacy single-profile fallback when `AI_PROFILES_JSON` is removed or left empty: `AI_API_KEY`, `AI_PROVIDER_API`, `OPENAI_API_BASE_URL`, and `OPENAI_MODEL` create `Default AI (.env)`. Those fields stay in `.env.example` because the code reads them. `BACKEND_HOST`, `BACKEND_PORT`, and `WORD_ADDIN_API_BASE_URL` are omitted because the current code does not read them; backend listen address is controlled by the `uvicorn ... --host/--port` command, and the add-in development proxy is configured in `word-addin/webpack.config.js`. See [spec.md](spec.md) for full environment-variable rules and [DEPLOYMENT.md](DEPLOYMENT.md) for the Windows pilot configuration.

The backend returns only profile id, label, model, and supported API modes to the add-in. It never returns API keys.

Default SQLite locations:

- Agent trace: `backend/var/agent-traces/traces.sqlite3`, optionally controlled by `AGENT_TRACE_DIR`.
- V2 workbench: `backend/var/agent-workspace/projects.sqlite3`, optionally controlled by `AGENT_WORKSPACE_DIR`.
- DOCX download index: `backend/var/docx-results/results.sqlite3`, controlled by `DOCX_OUTPUT_DIR`; it records metadata needed to recover downloads and new task `run_id` values.

These directories live under `backend/var/` by default and are not committed to Git.

## Start a Local oMLX AI Service

For local real-AI integration, oMLX can expose an OpenAI-compatible service. The helper script defaults to `$HOME/AI/models`, `$HOME/.omlx`, no API key, and port `8001`; override `OMLX_MODEL_DIR`, `OMLX_BASE_PATH`, `OMLX_PORT`, `OMLX_API_KEY`, or `OMLX_PRELOAD_MODEL` as needed. If you enable bearer auth for the local service, set the same local-only value in `OMLX_API_KEY` and `.env` `AI_API_KEY`.

```bash
OMLX_MODEL_DIR="$HOME/AI/models" ./scripts/start-omlx.sh
```

Model-service check:

```bash
curl_headers=()
if [ -n "${OMLX_API_KEY:-}" ]; then
  curl_headers=(-H "Authorization: Bearer ${OMLX_API_KEY}")
fi
curl --noproxy 127.0.0.1 http://127.0.0.1:8001/v1/models \
  "${curl_headers[@]}"
```

If `/v1/models` returns a model id that differs from `.env`, update the matching `AI_PROFILES_JSON` profile's `model`. If you intentionally use the legacy single-profile fallback, update `OPENAI_MODEL` instead.

## Start the Backend

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000 --reload
```

Health check:

```bash
curl --noproxy 127.0.0.1 http://127.0.0.1:8000/health
```

Expected response:

```json
{"status":"ok"}
```

## Start the Word Add-in

For a Windows editor-machine setup from scratch, see [DEPLOYMENT.md](DEPLOYMENT.md).

```bash
cd word-addin
npm install
npm run dev-server
```

The development service runs at:

```text
https://localhost:3000/taskpane.html
```

Sideload into Word:

```bash
cd word-addin
npm run start
```

## Integration Flow

1. Start oMLX or configure a remote OpenAI-compatible API.
2. Start the backend and confirm `/health` returns `{"status":"ok"}`.
3. Start the `word-addin` dev server.
4. Run `npm run start` to sideload the add-in into Word.
5. Current-selection review: select body text in Word. Whole-book review: prepare a `.docx` file.
6. Open the task pane, choose "Current selection" or "Whole-book DOCX", and fill in the book title. Whole-book mode requires selecting a `.docx` file.
7. Expand "Additional information" or "More settings (pilot support)" as needed to adjust review goal, model profile, review mode, temperature, and writeback mode.
8. Click "Start review". The add-in refreshes progress automatically and shows suggestions when the run finishes.
9. Review original text, replacement, explanation, and evidence in the suggestions area. Accept, reject, or bulk-process pending suggestions.
10. For current selections, click "Write back accepted suggestions" to let the add-in write to Word. For DOCX projects, writeback generates a backend output file, then click "Download reviewed file".
11. Validate refresh progress, continue-waiting behavior for long tasks, recent review reopening, troubleshooting details, DOCX download, and report summary.

## Tests and Validation

Common checks are listed below. See [TESTING.md](TESTING.md) for the full validation matrix.

Backend:

```bash
cd backend
source .venv/bin/activate
python -m pytest -q
```

Frontend:

```bash
cd word-addin
npm run lint
npm run build
```

Manifest validation:

```bash
cd word-addin
npm run validate
```

`npm run validate` needs access to the Microsoft Office manifest validation service and can fail when offline or network-restricted.

## Documentation Maintenance

- API contract changes require updates to [spec.md](spec.md).
- Code entry point, module boundary, or V2 data-flow changes require updates to [ARCHITECTURE.md](ARCHITECTURE.md).
- Test command, validation matrix, or manual integration changes require updates to [TESTING.md](TESTING.md).
- Startup, port, environment-variable, or integration-flow changes require updates to this file.
- V2 Agent workbench goal, architecture boundary, or collaboration-rule changes require updates to [AGENTS.md](AGENTS.md).
- Windows pilot deployment changes require updates to [DEPLOYMENT.md](DEPLOYMENT.md).
- Public release, security, or contribution process changes require updates to [SECURITY.md](SECURITY.md) or [CONTRIBUTING.md](CONTRIBUTING.md).
- Every Markdown documentation change must keep the English canonical file and the matching `_cn.md` file content-equivalent.
- Do not put temporary troubleshooting notes in long-lived docs. If they must be kept, put them in a Git-ignored temporary directory.
