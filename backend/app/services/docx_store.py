from __future__ import annotations

import shutil
import sqlite3
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from pathlib import Path

from app.schemas import ApplicationMode, ChunkedTaskStatus
from app.settings import Settings, get_settings


@dataclass(frozen=True)
class StoredDocxResult:
    task_id: str
    source_filename: str
    output_filename: str
    application_mode: ApplicationMode
    status: ChunkedTaskStatus
    total_chunks: int
    completed_chunks: int
    failed_chunks: int
    issue_count: int
    relative_path: str
    created_at: str
    updated_at: str
    expires_at: str


class DocxResultExpired(RuntimeError):
    """Raised when a stored DOCX result has passed its retention window."""


class DocxResultFileMissing(RuntimeError):
    """Raised when metadata exists but the DOCX result file is gone."""


def output_dir(settings: Settings | None = None) -> Path:
    resolved_settings = settings or get_settings()
    return resolved_settings.docx_output_dir


def retention_days(settings: Settings | None = None) -> int:
    resolved_settings = settings or get_settings()
    return resolved_settings.docx_retention_days


def result_path(relative_path: str, settings: Settings | None = None) -> Path:
    root = output_dir(settings).resolve()
    path = (root / relative_path).resolve()
    if root != path and root not in path.parents:
        raise DocxResultFileMissing("DOCX result path is outside the configured output directory.")
    return path


def build_expires_at(settings: Settings | None = None, now: datetime | None = None) -> str:
    base = now or datetime.now(UTC)
    return (base + timedelta(days=retention_days(settings))).isoformat()


def save_result(
    *,
    task_id: str,
    source_filename: str,
    output_filename: str,
    application_mode: ApplicationMode,
    status: ChunkedTaskStatus,
    total_chunks: int,
    completed_chunks: int,
    failed_chunks: int,
    issue_count: int,
    relative_path: str,
    created_at: str,
    updated_at: str,
    expires_at: str,
    settings: Settings | None = None,
) -> None:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        connection.execute(
            """
            INSERT INTO docx_results (
                task_id, source_filename, output_filename, application_mode, status,
                total_chunks, completed_chunks, failed_chunks, issue_count,
                relative_path, created_at, updated_at, expires_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(task_id) DO UPDATE SET
                source_filename = excluded.source_filename,
                output_filename = excluded.output_filename,
                application_mode = excluded.application_mode,
                status = excluded.status,
                total_chunks = excluded.total_chunks,
                completed_chunks = excluded.completed_chunks,
                failed_chunks = excluded.failed_chunks,
                issue_count = excluded.issue_count,
                relative_path = excluded.relative_path,
                updated_at = excluded.updated_at,
                expires_at = excluded.expires_at
            """,
            (
                task_id,
                source_filename,
                output_filename,
                application_mode,
                status,
                total_chunks,
                completed_chunks,
                failed_chunks,
                issue_count,
                relative_path,
                created_at,
                updated_at,
                expires_at,
            ),
        )
        connection.commit()


def get_result(task_id: str, settings: Settings | None = None) -> StoredDocxResult | None:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        row = connection.execute(
            """
            SELECT task_id, source_filename, output_filename, application_mode, status,
                   total_chunks, completed_chunks, failed_chunks, issue_count,
                   relative_path, created_at, updated_at, expires_at
            FROM docx_results
            WHERE task_id = ?
            """,
            (task_id,),
        ).fetchone()

    return _row_to_result(row) if row else None


def resolve_download(task_id: str, settings: Settings | None = None) -> tuple[Path, StoredDocxResult]:
    stored = get_result(task_id, settings)
    if not stored:
        raise KeyError(task_id)

    if _parse_iso(stored.expires_at) <= datetime.now(UTC):
        cleanup_expired(settings)
        raise DocxResultExpired("DOCX proofread result has expired. Please proofread the file again.")

    path = result_path(stored.relative_path, settings)
    if not path.exists():
        raise DocxResultFileMissing("DOCX proofread result file is missing. Please proofread the file again.")

    return path, stored


def cleanup_expired(settings: Settings | None = None, now: datetime | None = None) -> int:
    _ensure_schema(settings)
    cutoff = (now or datetime.now(UTC)).isoformat()
    with _connect(settings) as connection:
        rows = connection.execute(
            "SELECT task_id, relative_path FROM docx_results WHERE expires_at <= ?",
            (cutoff,),
        ).fetchall()
        connection.execute("DELETE FROM docx_results WHERE expires_at <= ?", (cutoff,))
        connection.commit()

    for task_id, relative_path in rows:
        _delete_result_files(task_id, relative_path, settings)

    return len(rows)


def clear_store_for_tests(settings: Settings | None = None) -> None:
    root = output_dir(settings)
    if root.exists():
        shutil.rmtree(root)


def _connect(settings: Settings | None = None) -> sqlite3.Connection:
    root = output_dir(settings)
    root.mkdir(parents=True, exist_ok=True)
    return sqlite3.connect(root / "results.sqlite3")


def _ensure_schema(settings: Settings | None = None) -> None:
    with _connect(settings) as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS docx_results (
                task_id TEXT PRIMARY KEY,
                source_filename TEXT NOT NULL,
                output_filename TEXT NOT NULL,
                application_mode TEXT NOT NULL,
                status TEXT NOT NULL,
                total_chunks INTEGER NOT NULL,
                completed_chunks INTEGER NOT NULL,
                failed_chunks INTEGER NOT NULL,
                issue_count INTEGER NOT NULL,
                relative_path TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                expires_at TEXT NOT NULL
            )
            """
        )
        connection.commit()


def _row_to_result(row: tuple) -> StoredDocxResult:
    return StoredDocxResult(
        task_id=row[0],
        source_filename=row[1],
        output_filename=row[2],
        application_mode=row[3],
        status=row[4],
        total_chunks=row[5],
        completed_chunks=row[6],
        failed_chunks=row[7],
        issue_count=row[8],
        relative_path=row[9],
        created_at=row[10],
        updated_at=row[11],
        expires_at=row[12],
    )


def _delete_result_files(task_id: str, relative_path: str, settings: Settings | None = None) -> None:
    root = output_dir(settings).resolve()
    path = result_path(relative_path, settings)
    task_dir = root / task_id
    if path.exists():
        path.unlink()
    if task_dir.exists() and task_dir.is_dir():
        shutil.rmtree(task_dir, ignore_errors=True)


def _parse_iso(value: str) -> datetime:
    parsed = datetime.fromisoformat(value)
    return parsed if parsed.tzinfo else parsed.replace(tzinfo=UTC)
