# Word AI Proofreader Architecture

Language: English | [简体中文](ARCHITECTURE_cn.md)

This document is for developers and AI coding agents. It explains the current V2.2 publishing review workbench entry points, module boundaries, and end-to-end data flow. Product goals and collaboration constraints are defined in [AGENTS.md](AGENTS.md). The API contract is defined in [spec.md](spec.md).

## Code Entry Points

- FastAPI entry point: `backend/app/main.py`, responsible for HTTP routes, request validation, background task startup, error-status mapping, and download responses.
- API schemas: `backend/app/schemas.py`, defining V2 project, run, candidate, memory, report, writeback, and retained lower-level direct API request/response models.
- V2 workbench orchestration: `backend/app/agents/workspace.py`, responsible for project-scoped runs, chunked review, candidate merging, evaluator rechecks, failed-chunk retry, writeback orchestration, and run-event recording.
- V2 project persistence: `backend/app/services/project_store.py`, responsible for SQLite schema, project/run/candidate/memory/report/output metadata, and latest-run semantics.
- DOCX capabilities: `backend/app/services/docx.py`, `backend/app/services/docx_service.py`, `backend/app/services/docx_tasks.py`, and `backend/app/services/docx_store.py`, responsible for DOCX extraction, location, OOXML comments/revisions, lower-level DOCX tasks, and download indexes.
- AI configuration and calls: `backend/app/services/ai_profiles.py`, `backend/app/services/ai_client.py`, and `backend/app/services/ai_provider_service.py`, responsible for profile parsing, Responses/Chat adaptation, dynamic timeouts, mock fallback, and secret isolation.
- Agent trace: `backend/app/agents/trace.py` and V2 run events in `backend/app/services/project_store.py`; the former serves lower-level direct API traces, and the latter serves V2 workbench event streams.
- Document maps and chunks: `backend/app/services/document_map_service.py`, `backend/app/services/chunk_service.py`, and `backend/app/services/chunking.py`.
- Frontend workbench: `word-addin/src/taskpane/taskpane.ts` owns UI state and flow orchestration, `api.ts` owns backend calls, `word.ts` owns Office.js location and current-selection writeback, `types.ts` owns frontend types, and `debug.ts` owns troubleshooting logs.

## V2 End-to-End Flow

1. The Word add-in reads the current selection or accepts a selected `.docx` file.
2. The add-in calls `POST /api/v2/projects/selection` or `POST /api/v2/projects` to create a review project.
3. The backend saves the project, builds a document map, and creates the default V2.2 review plan.
4. The add-in calls `POST /api/v2/projects/{project_id}/runs` to start a background run.
5. `agents/workspace.py` executes `proofread_pass` over document chunks and reuses validated AI review, chunking, location, and DOCX capabilities.
6. The backend merges candidates, runs evaluator rechecks, and stores candidates, report, memory summary, and run events.
7. A run with candidates enters `waiting_for_approval`; a run with zero candidates enters `succeeded`. A run with some failed chunks and usable results may enter `partial_succeeded`.
8. The editor accepts, rejects, or defers candidates in the add-in. Bulk operations only affect latest-run pending candidates.
9. Current-selection projects are written back by Office.js and then call `mark-written`; DOCX projects call backend `writeback`, which regenerates a reviewed file from the original DOCX.
10. After DOCX writeback, the add-in downloads the new file through `/download`. If retrying failed chunks later creates new candidates, old output remains but the project is marked `output_stale=true`.

## Module Boundaries

- `agents/` owns orchestration, planning, tool selection, state progression, rechecks, human-confirmation state, and observable events.
- `services/` owns reusable business capabilities, especially DOCX parsing, chunking, AI calls, location, OOXML writeback, SQLite persistence, and report generation.
- `word-addin/src/taskpane/` owns Office.js integration, task-pane UI, polling/refresh, candidate display, current-selection writeback, and DOCX download triggering.
- `tools.py` must stay a thin wrapper and must not duplicate large business logic.
- Do not rewrite validated DOCX, location, and writeback low-level services from scratch. When V2 workbench behavior must change, prefer small changes at orchestration or service boundaries.
- When API fields change, check `backend/app/schemas.py`, `backend/app/main.py`, `word-addin/src/taskpane/types.ts`, `word-addin/src/taskpane/api.ts`, and [spec.md](spec.md).

## Common Change Points

- V2 run status, failed retries, `waiting_for_approval`, `partial_succeeded`: start with `backend/app/agents/workspace.py`, then inspect `backend/app/services/project_store.py` and `word-addin/src/taskpane/taskpane.ts`.
- Candidate pagination, bulk accept/reject, latest-run semantics: start with `project_store.py` and candidate routes in `main.py`, then inspect `api.ts` and `taskpane.ts`.
- DOCX writeback, author, comments/revisions, missing downloads: start with `docx.py`, `docx_service.py`, and `workspace.py` `write_approved`, then inspect output metadata in `project_store.py` and frontend download logic.
- AI profiles, Responses/Chat differences, Xiaomi MiMo, mock fallback: start with `ai_profiles.py`, `ai_client.py`, and `ai_provider_service.py`, then inspect `.env.example` and frontend profile selection.
- Trace redaction, run events, current-chunk timeout: start with run-event writes in `workspace.py` and `trace.py`, then inspect progress rendering in `taskpane.ts`.
- Frontend text, button state, refresh scope, location/writeback interaction: start with `taskpane.ts`; if Office document operations are involved, inspect `word.ts`.

## Not Sources of Code Truth

The following directories contain local state, dependencies, caches, or build output. Do not use them as implementation truth or edit targets:

- `backend/var/`
- `backend/.venv/`
- `.pytest_cache/`
- `backend/.pytest_cache/`
- `word-addin/node_modules/`
- `word-addin/dist/`

To verify current behavior, prefer source code, schemas, tests, and [spec.md](spec.md). Do not infer contracts from generated or local-state directories.
