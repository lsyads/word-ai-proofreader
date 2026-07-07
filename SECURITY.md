# Security Policy

Language: English | [简体中文](SECURITY_cn.md)

## Supported Status

This project is preparing for public open source release. Security reports are accepted for the current default branch and the latest tagged release once releases are published. Older local pilot snapshots, generated files, ignored caches, and untracked deployment copies are not supported.

## Reporting a Vulnerability

Please do not open a public issue for a vulnerability that exposes secrets, unpublished manuscript text, or a working exploit. Report it privately through the repository security advisory channel if available, or contact the maintainers through the project repository owner.

Include:

- Affected commit, branch, or release.
- Reproduction steps.
- Impact and who can trigger it.
- Whether API keys, Authorization headers, Bearer tokens, manuscript text, DOCX files, traces, or SQLite workspace files are involved.

Do not include real API keys, full unpublished manuscript text, or private documents in the report. Use redacted examples or synthetic documents.

## Secrets and Credentials

API keys must stay in backend runtime environment files such as `.env`. Never commit real keys, tokens, certificates, or private deployment configs. The repository `.gitignore` excludes `.env`, local certificates, `backend/var/`, build output, dependency folders, and caches, but contributors should still run a secret scan before opening pull requests that touch configuration or history.

Before making the repository public, manually run the GitHub Actions `Secret Scan` workflow and require the full-history Gitleaks and TruffleHog jobs to pass. Local Gitleaks or TruffleHog runs are useful but optional when the Actions workflow provides the release record.

The backend and agent traces are designed not to record full text, API keys, Authorization headers, or Bearer tokens. If you find a path that logs or persists those values, treat it as a security issue.

## Manuscript Privacy

This tool handles Word manuscript content. In real AI mode, selected text or DOCX chunks are sent to the configured AI provider. Contributors and deployers must explain that data flow to users and should use manuscript copies during pilots.

Project memory should store structured publishing facts such as terminology, book conventions, and editor preferences. It must not store full manuscript text by default.

## Safe Defaults

- Human confirmation is required before current-selection or whole-book writeback.
- `批注模式` is the conservative writeback mode for pilots.
- Backend `INFO` logs should not print full request text. Use `DEBUG` only for local debugging with non-sensitive documents.
- Do not attach real `.docx` manuscripts, SQLite workspaces, `.env` files, or provider responses to public issues.
