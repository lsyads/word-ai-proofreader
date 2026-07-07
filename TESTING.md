# Word AI Proofreader Testing

Language: English | [简体中文](TESTING_cn.md)

This document collects validation commands for development and AI coding work. It is the primary minimum validation matrix. Product acceptance criteria are in [spec.md](spec.md), and Windows pilot deployment validation is in [DEPLOYMENT.md](DEPLOYMENT.md).

## Quick Checks

Backend tests, offline:

```bash
cd backend
source .venv/bin/activate
python -m pytest -q
```

Frontend static checks and build, offline:

```bash
cd word-addin
npm run lint
npm run build
```

Python dependency audit, requires network and `pip-audit` locally; otherwise use the GitHub Actions `CI` workflow:

```bash
cd backend
python -m pip install pip-audit
python -m pip_audit -r requirements.txt --strict
```

Manifest validation, requires access to the Microsoft Office manifest validation service:

```bash
cd word-addin
npm run validate
```

Local health check, requires the backend to be running:

```bash
curl --noproxy 127.0.0.1 http://127.0.0.1:8000/health
```

Word manual integration, requires Word desktop, the Office add-in dev server, and a usable backend:

```bash
cd word-addin
npm run dev-server
npm run start
```

## Minimum Validation Matrix

| Change type | Minimum checks |
| --- | --- |
| API/schema/status fields | `cd backend && python -m pytest -q`; check [spec.md](spec.md), `word-addin/src/taskpane/types.ts`, and `api.ts` |
| V2 Agent run, status progression, failed retry | `backend/tests/test_v2_workspace.py`, `backend/tests/test_agents.py`; run the full backend suite when needed |
| Candidate queue, pagination, bulk decisions, latest-run semantics | `backend/tests/test_v2_workspace.py`; frontend changes also require `npm run lint` and `npm run build` |
| DOCX extraction, location, comments/revisions, download | `backend/tests/test_docx.py`, `backend/tests/test_v2_workspace.py`; frontend download changes also require frontend checks |
| AI profiles, Responses/Chat, mock fallback, timeout | `backend/tests/test_ai_profiles.py`, `backend/tests/test_ai_client.py`, `backend/tests/test_api.py` |
| Trace redaction, run events, current-chunk timeout | `backend/tests/test_agents.py`, `backend/tests/test_v2_workspace.py`; frontend progress-display changes also require frontend checks |
| Frontend UI, button states, refresh, location, writeback | `cd word-addin && npm run lint && npm run build`; key Office.js behavior requires Word manual integration |
| Frontend dependency or lockfile changes | `cd word-addin && npm ci && npm audit --registry=https://registry.npmjs.org && npm run lint && npm run build`; audit needs the official npm registry because some mirrors do not implement the audit endpoint |
| Backend dependency changes | GitHub Actions `CI` workflow `Python dependency audit`, or local `cd backend && python -m pip_audit -r requirements.txt --strict` when `pip-audit` is installed |
| Environment variables, startup, deployment docs | `backend/tests/test_settings.py`, `backend/tests/test_env_example.py`; check `.env.example`, [README.md](README.md), and [DEPLOYMENT.md](DEPLOYMENT.md) |
| Markdown documentation | Check English/cn pairing, language switches, matching internal links, and stale template or personal strings |
| Public release gate | Run GitHub Actions `CI` and `Secret Scan` manually with `workflow_dispatch`; release only after backend tests, Python audit, frontend audit/lint/build, Gitleaks, and TruffleHog are green |

Run a single backend test file like this:

```bash
cd backend
source .venv/bin/activate
python -m pytest -q tests/test_v2_workspace.py
```

## Check Types

- Offline: backend pytest, frontend lint, frontend build.
- Requires network: `npm run validate`, `pip-audit`, real remote AI API connectivity checks, GitHub Actions dependency downloads, and full-history secret-scan actions.
- Requires npm security API: `cd word-addin && npm audit --registry=https://registry.npmjs.org`.
- Requires GitHub Actions for the public-release record when local tools are unavailable: `CI` and `Secret Scan` should be manually dispatched and green before making the repository public.
- Requires Word desktop: add-in sideloading, original-text location, current-selection writeback, and manual confirmation of comments/revisions.
- Requires a real AI key: real Responses/Chat calls, remote model timeout behavior, and provider compatibility checks.
- Does not require a real AI key: mock fallback, schema validation, candidate state, project storage, and DOCX writeback unit tests.

## Manual Integration Checklist

1. Start the backend and confirm `/health` returns `{"status":"ok"}`.
2. Start the `word-addin` dev server and sideload the add-in into Word.
3. Current-selection review: select text, fill in `书名`, click `开始审校`, then confirm candidate display, `定位原文`, accept/reject, and writeback.
4. DOCX review: choose a `.docx` file, click `开始审校`, accept candidates, run backend writeback, and click `下载审校后文件`.
5. Long-task scenario: confirm chunk progress, current-chunk timeout countdown, `刷新进度`, and `继续等待` controls do not report waiting as failure.
6. Failed-chunk scenario: confirm `retry-failed` only retries failed chunks, new candidates enter the same run, and previously written output is marked `output_stale=true` when necessary.

## Documentation Checks

This project does not have a dedicated Markdown linter. At minimum, manually check:

- Relative links open from the repository root.
- Commands match `word-addin/package.json` and `backend/requirements.txt`.
- New API fields, status values, or environment variables are synchronized across [spec.md](spec.md), [README.md](README.md), and related test documentation.
- The English canonical Markdown file and matching `_cn.md` file stay content-equivalent.
- English docs link to English docs, while Chinese docs link to `_cn.md` counterparts, except language-switch links.
