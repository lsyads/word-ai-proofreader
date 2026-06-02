from __future__ import annotations

import json
import shutil
import sqlite3
import uuid
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from app.schemas import (
    BookInfo,
    V2CandidateIssue,
    V2DocumentMapResponse,
    V2ProjectResponse,
    V2ReviewPlanResponse,
    V2ReviewReportResponse,
    V2RunEventResponse,
    V2RunResponse,
)
from app.settings import Settings, get_settings


class V2ProjectNotFound(KeyError):
    """Raised when a V2 review project cannot be found."""


class V2RunNotFound(KeyError):
    """Raised when a V2 review run cannot be found."""


class V2CandidateNotFound(KeyError):
    """Raised when a V2 candidate issue cannot be found."""


@dataclass(frozen=True)
class StoredProject:
    project_id: str
    source_type: str
    status: str
    source_filename: str
    text_preview: str | None
    book: BookInfo
    review_goal: str
    source_bytes: bytes
    created_at: str
    updated_at: str
    output_filename: str | None
    output_relative_path: str | None


def workspace_dir(settings: Settings | None = None) -> Path:
    resolved_settings = settings or get_settings()
    return resolved_settings.agent_workspace_dir


def project_output_path(project_id: str, filename: str, settings: Settings | None = None) -> Path:
    root = workspace_dir(settings).resolve()
    path = (root / "outputs" / project_id / filename).resolve()
    if root != path and root not in path.parents:
        raise V2ProjectNotFound("V2 output path is outside workspace.")
    return path


def create_project(
    *,
    source_type: str = "docx",
    source_filename: str,
    source_bytes: bytes,
    book: BookInfo,
    review_goal: str,
    text_preview: str | None = None,
    settings: Settings | None = None,
) -> StoredProject:
    _ensure_schema(settings)
    project_id = f"project_{uuid.uuid4().hex}"
    now = _now_iso()
    with _connect(settings) as connection:
        connection.execute(
            """
            INSERT INTO v2_projects (
                project_id, source_type, status, source_filename, text_preview, book_json,
                review_goal, source_bytes, created_at, updated_at, output_filename,
                output_relative_path
            )
            VALUES (?, ?, 'created', ?, ?, ?, ?, ?, ?, ?, NULL, NULL)
            """,
            (
                project_id,
                source_type,
                source_filename,
                text_preview,
                json.dumps(book.model_dump(), ensure_ascii=False),
                review_goal,
                source_bytes,
                now,
                now,
            ),
        )
        connection.commit()
    return require_project(project_id, settings)


def require_project(project_id: str, settings: Settings | None = None) -> StoredProject:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        row = connection.execute(
            """
            SELECT project_id, source_type, status, source_filename, text_preview,
                   book_json, review_goal, source_bytes, created_at, updated_at,
                   output_filename, output_relative_path
            FROM v2_projects
            WHERE project_id = ?
            """,
            (project_id,),
        ).fetchone()
    if not row:
        raise V2ProjectNotFound(project_id)
    return _row_to_project(row)


def update_project_status(project_id: str, status: str, settings: Settings | None = None) -> None:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        connection.execute(
            "UPDATE v2_projects SET status = ?, updated_at = ? WHERE project_id = ?",
            (status, _now_iso(), project_id),
        )
        connection.commit()


def save_project_output(
    project_id: str,
    *,
    output_filename: str,
    output_path: Path,
    settings: Settings | None = None,
) -> None:
    _ensure_schema(settings)
    root = workspace_dir(settings).resolve()
    relative_path = str(output_path.resolve().relative_to(root))
    with _connect(settings) as connection:
        connection.execute(
            """
            UPDATE v2_projects
            SET status = 'written', output_filename = ?, output_relative_path = ?, updated_at = ?
            WHERE project_id = ?
            """,
            (output_filename, relative_path, _now_iso(), project_id),
        )
        connection.commit()


def save_document_map(document_map: V2DocumentMapResponse, settings: Settings | None = None) -> None:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        connection.execute(
            """
            INSERT INTO v2_document_maps (project_id, map_json, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(project_id) DO UPDATE SET
                map_json = excluded.map_json,
                updated_at = excluded.updated_at
            """,
            (document_map.project_id, document_map.model_dump_json(), _now_iso()),
        )
        connection.commit()


def get_document_map(project_id: str, settings: Settings | None = None) -> V2DocumentMapResponse:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        row = connection.execute("SELECT map_json FROM v2_document_maps WHERE project_id = ?", (project_id,)).fetchone()
    if not row:
        raise V2ProjectNotFound(project_id)
    return V2DocumentMapResponse.model_validate_json(row[0])


def save_review_plan(plan: V2ReviewPlanResponse, settings: Settings | None = None) -> None:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        connection.execute(
            """
            INSERT INTO v2_review_plans (project_id, plan_json, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(project_id) DO UPDATE SET
                plan_json = excluded.plan_json,
                updated_at = excluded.updated_at
            """,
            (plan.project_id, plan.model_dump_json(), _now_iso()),
        )
        connection.commit()


def get_review_plan(project_id: str, settings: Settings | None = None) -> V2ReviewPlanResponse:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        row = connection.execute("SELECT plan_json FROM v2_review_plans WHERE project_id = ?", (project_id,)).fetchone()
    if not row:
        raise V2ProjectNotFound(project_id)
    return V2ReviewPlanResponse.model_validate_json(row[0])


def create_run(project_id: str, *, total_chunks: int, settings: Settings | None = None) -> V2RunResponse:
    _ensure_schema(settings)
    require_project(project_id, settings)
    run_id = f"v2_run_{uuid.uuid4().hex}"
    now = _now_iso()
    with _connect(settings) as connection:
        connection.execute(
            """
            INSERT INTO v2_runs (
                run_id, project_id, status, stage, total_chunks, completed_chunks,
                failed_chunks, candidate_count, error_message, created_at, updated_at
            )
            VALUES (?, ?, 'queued', 'queued', ?, 0, 0, 0, NULL, ?, ?)
            """,
            (run_id, project_id, total_chunks, now, now),
        )
        connection.commit()
    return require_run(project_id, run_id, settings)


def require_run(project_id: str, run_id: str, settings: Settings | None = None) -> V2RunResponse:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        row = connection.execute(
            """
            SELECT project_id, run_id, status, stage, total_chunks, completed_chunks,
                   failed_chunks, candidate_count, error_message, created_at, updated_at
            FROM v2_runs
            WHERE project_id = ? AND run_id = ?
            """,
            (project_id, run_id),
        ).fetchone()
    if not row:
        raise V2RunNotFound(run_id)
    return V2RunResponse(
        project_id=row[0],
        run_id=row[1],
        status=row[2],
        stage=row[3],
        total_chunks=row[4],
        completed_chunks=row[5],
        failed_chunks=row[6],
        candidate_count=row[7],
        error_message=row[8],
        created_at=row[9],
        updated_at=row[10],
    )


def update_run(
    project_id: str,
    run_id: str,
    *,
    status: str | None = None,
    stage: str | None = None,
    completed_chunks: int | None = None,
    failed_chunks: int | None = None,
    candidate_count: int | None = None,
    error_message: str | None = None,
    settings: Settings | None = None,
) -> None:
    _ensure_schema(settings)
    fields = ["updated_at = ?"]
    values: list[Any] = [_now_iso()]
    for column, value in (
        ("status", status),
        ("stage", stage),
        ("completed_chunks", completed_chunks),
        ("failed_chunks", failed_chunks),
        ("candidate_count", candidate_count),
    ):
        if value is not None:
            fields.append(f"{column} = ?")
            values.append(value)
    if error_message is not None:
        fields.append("error_message = ?")
        values.append(error_message)
    elif status in {"running", "waiting_for_approval", "succeeded", "partial_succeeded"}:
        fields.append("error_message = NULL")
    values.extend([project_id, run_id])
    with _connect(settings) as connection:
        connection.execute(
            f"UPDATE v2_runs SET {', '.join(fields)} WHERE project_id = ? AND run_id = ?",
            values,
        )
        connection.commit()


def add_run_event(
    project_id: str,
    run_id: str,
    event: str,
    data: dict[str, Any],
    settings: Settings | None = None,
) -> None:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        connection.execute(
            """
            INSERT INTO v2_run_events (project_id, run_id, event, data_json, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (project_id, run_id, event, json.dumps(data, ensure_ascii=False, default=str), _now_iso()),
        )
        connection.commit()


def list_run_events(project_id: str, run_id: str, settings: Settings | None = None) -> list[V2RunEventResponse]:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        rows = connection.execute(
            """
            SELECT event, data_json, created_at
            FROM v2_run_events
            WHERE project_id = ? AND run_id = ?
            ORDER BY id
            """,
            (project_id, run_id),
        ).fetchall()
    return [
        V2RunEventResponse(event=row[0], data=json.loads(row[1]), created_at=row[2])
        for row in rows
    ]


def save_candidates(candidates: list[V2CandidateIssue], settings: Settings | None = None) -> None:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        for candidate in candidates:
            connection.execute(
                """
                INSERT INTO v2_candidate_issues (
                    candidate_id, project_id, run_id, status, issue_json, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(candidate_id) DO UPDATE SET
                    status = excluded.status,
                    issue_json = excluded.issue_json,
                    updated_at = excluded.updated_at
                """,
                (
                    candidate.candidate_id,
                    candidate.project_id,
                    candidate.run_id,
                    candidate.status,
                    candidate.model_dump_json(),
                    candidate.created_at,
                    candidate.updated_at,
                ),
            )
        connection.commit()


def list_candidates(project_id: str, settings: Settings | None = None) -> list[V2CandidateIssue]:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        rows = connection.execute(
            """
            SELECT issue_json
            FROM v2_candidate_issues
            WHERE project_id = ?
            ORDER BY created_at, candidate_id
            """,
            (project_id,),
        ).fetchall()
    return [V2CandidateIssue.model_validate_json(row[0]) for row in rows]


def update_candidate_statuses(
    project_id: str,
    decisions: dict[str, str],
    settings: Settings | None = None,
) -> int:
    _ensure_schema(settings)
    candidates = {candidate.candidate_id: candidate for candidate in list_candidates(project_id, settings)}
    missing = set(decisions) - set(candidates)
    if missing:
        raise V2CandidateNotFound(next(iter(missing)))

    now = _now_iso()
    updated = 0
    with _connect(settings) as connection:
        for candidate_id, status in decisions.items():
            candidate = candidates[candidate_id].model_copy(update={"status": status, "updated_at": now})
            connection.execute(
                """
                UPDATE v2_candidate_issues
                SET status = ?, issue_json = ?, updated_at = ?
                WHERE project_id = ? AND candidate_id = ?
                """,
                (status, candidate.model_dump_json(), now, project_id, candidate_id),
            )
            updated += 1
        connection.commit()
    return updated


def mark_candidates_written(project_id: str, candidate_ids: list[str], settings: Settings | None = None) -> int:
    return update_candidate_statuses(project_id, {candidate_id: "written" for candidate_id in candidate_ids}, settings)


def save_report(report: V2ReviewReportResponse, settings: Settings | None = None) -> None:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        connection.execute(
            """
            INSERT INTO v2_reports (project_id, report_json, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(project_id) DO UPDATE SET
                report_json = excluded.report_json,
                updated_at = excluded.updated_at
            """,
            (report.project_id, report.model_dump_json(), _now_iso()),
        )
        connection.commit()


def get_report(project_id: str, settings: Settings | None = None) -> V2ReviewReportResponse | None:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        row = connection.execute("SELECT report_json FROM v2_reports WHERE project_id = ?", (project_id,)).fetchone()
    return V2ReviewReportResponse.model_validate_json(row[0]) if row else None


def project_response(project_id: str, settings: Settings | None = None) -> V2ProjectResponse:
    project = require_project(project_id, settings)
    candidates = list_candidates(project_id, settings)
    run_count = _run_count(project_id, settings)
    return V2ProjectResponse(
        project_id=project.project_id,
        source_type=project.source_type,
        status=project.status,
        source_filename=project.source_filename,
        text_preview=project.text_preview,
        book=project.book,
        review_goal=project.review_goal,
        created_at=project.created_at,
        updated_at=project.updated_at,
        run_count=run_count,
        candidate_count=len(candidates),
        pending_count=sum(1 for candidate in candidates if candidate.status == "pending"),
        approved_count=sum(1 for candidate in candidates if candidate.status == "approved"),
        output_filename=project.output_filename,
        download_url=f"/api/v2/projects/{project.project_id}/download" if project.output_filename else None,
    )


def output_download_path(project_id: str, settings: Settings | None = None) -> tuple[Path, str]:
    project = require_project(project_id, settings)
    if not project.output_filename or not project.output_relative_path:
        raise V2ProjectNotFound(project_id)
    root = workspace_dir(settings).resolve()
    path = (root / project.output_relative_path).resolve()
    if root != path and root not in path.parents:
        raise V2ProjectNotFound(project_id)
    if not path.exists():
        raise V2ProjectNotFound(project_id)
    return path, project.output_filename


def clear_store_for_tests(settings: Settings | None = None) -> None:
    root = workspace_dir(settings)
    if root.exists():
        shutil.rmtree(root)


def _connect(settings: Settings | None = None) -> sqlite3.Connection:
    root = workspace_dir(settings)
    root.mkdir(parents=True, exist_ok=True)
    return sqlite3.connect(root / "projects.sqlite3")


def _ensure_schema(settings: Settings | None = None) -> None:
    with _connect(settings) as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS v2_projects (
                project_id TEXT PRIMARY KEY,
                source_type TEXT NOT NULL DEFAULT 'docx',
                status TEXT NOT NULL,
                source_filename TEXT NOT NULL,
                text_preview TEXT,
                book_json TEXT NOT NULL,
                review_goal TEXT NOT NULL,
                source_bytes BLOB NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                output_filename TEXT,
                output_relative_path TEXT
            )
            """
        )
        columns = {row[1] for row in connection.execute("PRAGMA table_info(v2_projects)").fetchall()}
        if "source_type" not in columns:
            connection.execute("ALTER TABLE v2_projects ADD COLUMN source_type TEXT NOT NULL DEFAULT 'docx'")
        if "text_preview" not in columns:
            connection.execute("ALTER TABLE v2_projects ADD COLUMN text_preview TEXT")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS v2_document_maps (
                project_id TEXT PRIMARY KEY,
                map_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS v2_review_plans (
                project_id TEXT PRIMARY KEY,
                plan_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS v2_runs (
                run_id TEXT PRIMARY KEY,
                project_id TEXT NOT NULL,
                status TEXT NOT NULL,
                stage TEXT NOT NULL,
                total_chunks INTEGER NOT NULL,
                completed_chunks INTEGER NOT NULL,
                failed_chunks INTEGER NOT NULL,
                candidate_count INTEGER NOT NULL,
                error_message TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS v2_run_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id TEXT NOT NULL,
                run_id TEXT NOT NULL,
                event TEXT NOT NULL,
                data_json TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS v2_candidate_issues (
                candidate_id TEXT PRIMARY KEY,
                project_id TEXT NOT NULL,
                run_id TEXT NOT NULL,
                status TEXT NOT NULL,
                issue_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS v2_reports (
                project_id TEXT PRIMARY KEY,
                report_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        connection.commit()


def _row_to_project(row: tuple) -> StoredProject:
    return StoredProject(
        project_id=row[0],
        source_type=row[1],
        status=row[2],
        source_filename=row[3],
        text_preview=row[4],
        book=BookInfo.model_validate(json.loads(row[5])),
        review_goal=row[6],
        source_bytes=row[7],
        created_at=row[8],
        updated_at=row[9],
        output_filename=row[10],
        output_relative_path=row[11],
    )


def _run_count(project_id: str, settings: Settings | None = None) -> int:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        row = connection.execute("SELECT COUNT(*) FROM v2_runs WHERE project_id = ?", (project_id,)).fetchone()
    return int(row[0]) if row else 0


def _now_iso() -> str:
    return datetime.now(UTC).isoformat()
