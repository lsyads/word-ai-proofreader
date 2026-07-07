# Contributing

Language: English | [简体中文](CONTRIBUTING_cn.md)

Thanks for helping improve Word AI Proofreader. This project is a publishing-editor workbench, so contributions should preserve the product boundary: the agent produces suggestions and evidence, while editors confirm before any writeback.

## Development Setup

Backend:

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m pytest -q
```

Frontend:

```bash
cd word-addin
npm install
npm run lint
npm run build
```

Use `.env.example` as a template. Put real keys only in a local `.env`, never in tracked files.

## Contribution Rules

- Read [AGENTS.md](AGENTS.md) before making code changes.
- Keep `agents/` responsible for orchestration and human-confirmation state, and `services/` responsible for reusable business capabilities.
- Do not bypass editor confirmation before writeback.
- Do not store full manuscript text in long-term memory by default.
- Keep traces and logs free of full text, API keys, Authorization headers, and Bearer tokens.
- Avoid rewriting validated DOCX/location/writeback services unless the V2 workbench goal requires it.

## Documentation Sync

Documentation is bilingual:

- English `.md` files are canonical.
- Matching `_cn.md` files must stay content-equivalent.
- English docs link to English docs.
- Chinese docs link to `_cn.md` counterparts, except the language-switch links.

When changing an API, schema, status value, or environment variable, update [spec.md](spec.md) and [spec_cn.md](spec_cn.md). When changing entry points, module boundaries, or data flow, update [ARCHITECTURE.md](ARCHITECTURE.md) and [ARCHITECTURE_cn.md](ARCHITECTURE_cn.md). When changing test commands or validation expectations, update [TESTING.md](TESTING.md) and [TESTING_cn.md](TESTING_cn.md).

## Pull Request Checklist

- Explain the user-facing or developer-facing behavior changed.
- Mention which tests were run.
- Keep unrelated refactors out of the pull request.
- Include doc updates for changed contracts, commands, or product boundaries.
- Run a secret scan before opening a public PR that touches configuration, deployment docs, logs, or history.
- Do not include real manuscripts, API keys, `.env`, local SQLite workspaces, dependency folders, or build output.

## Manual Validation

Some behavior cannot be fully validated by unit tests:

- Office.js location and current-selection writeback require Word desktop.
- `npm run validate` requires the Microsoft Office manifest validation service.
- Real AI provider behavior requires a configured key and compatible endpoint.

Call out any skipped manual validation in the PR.
