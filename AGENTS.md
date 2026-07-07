# AGENTS.md

Language: English | [简体中文](AGENTS_cn.md)

This project is a Word document review agent system for publishing editors.

The current main flow is the V2.2 publishing review agent workbench. It reuses validated low-level document-processing capabilities, including DOCX parsing, chunked review, AI calls, original-text location, and comment/revision writeback. On top of those capabilities, it builds a project around one book so the agent can create a plan, call tools, recheck candidates, wait for editor confirmation, and accumulate publishing rules.

V2 does not need to be compatible with V1 existing data, old traces, old task snapshots, or old result indexes. Do not sacrifice current implementation clarity for old-data migration. The code may keep lower-level direct APIs for reuse of validated capabilities and development debugging.

## Required Reading Order for AI Coding

1. Read this file first to confirm V2 product goals, collaboration boundaries, and forbidden changes.
2. For implementation entry points, module responsibilities, or data flow, read `ARCHITECTURE.md`.
3. For APIs, schemas, status values, environment variables, or acceptance criteria, read `spec.md`.
4. For validation-command selection after changes, read `TESTING.md`.
5. For Windows pilot deployment flow changes, read `DEPLOYMENT.md`.

## V2 Product Boundary

The goal is a runnable, testable, demonstrable publishing review agent workbench. A publishing editor should do more than "call AI once": they should create a review project around a book, let the agent plan and execute in stages, inspect evidence, recheck candidate results, and write back to Word only after editor confirmation. Terminology, book conventions, and cross-chapter consistency can be accumulated as project memory and future extensions, but the current V2.2 review plan does not show them as separate execution steps.

V2 coding must follow these rules:

- Work around review projects, not one-off requests. Save the review goal, document map, run state, pending issues, and review report for a book.
- `agents/` owns orchestration, planning, tool selection, rechecking, human-confirmation state, and observable runs. `services/` owns reusable business capabilities.
- Do not rewrite validated DOCX, location, or writeback low-level services from scratch unless the V2 workbench product goal truly requires it.
- Do not implement a black-box "one super prompt does everything" flow, and do not put all logic into a single FastAPI endpoint.
- Tool inputs and outputs must be clearly typed. Use Pydantic schemas for complex inputs. Tool names use `snake_case`. `tools.py` is a thin wrapper layer only.
- Memory should prioritize terminology, person/place/organization name consistency, publisher style rules, editor-confirmed preferences, and book-level conventions. Do not store full manuscript text in long-term memory by default.
- The agent may generate candidate issues, risk notes, evidence, and suggestions only. Any current-selection or whole-book writeback must go through editor confirmation first. Silent automatic body-text modification is not allowed.
- Every agent run produces a `run_id`. Trace records decisions, nodes, tool-call summaries, chunk state, elapsed time, errors, and retry counts, but not full text, API keys, Authorization headers, or Bearer tokens.
- Formal API, status-value, environment-variable, or schema changes must update `spec.md`. Choose validation commands according to `TESTING.md`.
- Do not put temporary troubleshooting notes in long-lived docs. If they must be retained, put them in a Git-ignored temporary directory.
- Markdown documentation is bilingual for open source: the English `.md` file is canonical, and the matching `_cn.md` file must stay content-equivalent.
