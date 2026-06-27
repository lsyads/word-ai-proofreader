from __future__ import annotations

import json
import re
import sqlite3
import time
import uuid
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from app.agents.state import AgentChunkTrace, AgentFlow, AgentNodeTrace, AgentRunTrace
from app.schemas import ChunkedTaskStatus
from app.settings import Settings, get_settings

_SECRET_PATTERNS = (
    re.compile(r"(Authorization\s*[:=]\s*Bearer\s+)[^\s,;}]+", re.IGNORECASE),
    re.compile(r"(Bearer\s+)[A-Za-z0-9._\-]+", re.IGNORECASE),
    re.compile(r"(api[_-]?key\s*[:=]\s*)[^\s,;}]+", re.IGNORECASE),
)
_MAX_ERROR_LENGTH = 1000


class AgentTraceNotFound(KeyError):
    """Raised when an agent run trace cannot be found."""


def create_run(
    flow: AgentFlow,
    *,
    run_id: str | None = None,
    task_id: str | None = None,
    status: ChunkedTaskStatus | str = "queued",
    total_chunks: int = 0,
    metadata: dict[str, Any] | None = None,
    settings: Settings | None = None,
) -> str:
    resolved_run_id = run_id or f"agent_run_{uuid.uuid4().hex}"
    now = _now_iso()
    _ensure_schema(settings)
    with _connect(settings) as connection:
        connection.execute(
            """
            INSERT INTO agent_runs (
                run_id, flow, task_id, status, created_at, updated_at,
                total_chunks, completed_chunks, failed_chunks, issue_count,
                error_message, metadata_json
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, 0, 0, 0, NULL, ?)
            ON CONFLICT(run_id) DO UPDATE SET
                flow = excluded.flow,
                task_id = excluded.task_id,
                status = excluded.status,
                updated_at = excluded.updated_at,
                total_chunks = excluded.total_chunks,
                metadata_json = excluded.metadata_json
            """,
            (
                resolved_run_id,
                flow,
                task_id,
                status,
                now,
                now,
                total_chunks,
                json.dumps(_sanitize_metadata(metadata or {}), ensure_ascii=False),
            ),
        )
        connection.commit()
    return resolved_run_id


def update_run(
    run_id: str,
    *,
    status: ChunkedTaskStatus | str | None = None,
    total_chunks: int | None = None,
    completed_chunks: int | None = None,
    failed_chunks: int | None = None,
    issue_count: int | None = None,
    error_message: str | None = None,
    settings: Settings | None = None,
) -> None:
    _ensure_schema(settings)
    fields: list[str] = ["updated_at = ?"]
    values: list[Any] = [_now_iso()]

    for column, value in (
        ("status", status),
        ("total_chunks", total_chunks),
        ("completed_chunks", completed_chunks),
        ("failed_chunks", failed_chunks),
        ("issue_count", issue_count),
    ):
        if value is not None:
            fields.append(f"{column} = ?")
            values.append(value)

    if error_message is not None:
        fields.append("error_message = ?")
        values.append(_sanitize_text(error_message))
    elif status in {"running", "succeeded", "partial_succeeded", "cancelled"}:
        fields.append("error_message = NULL")

    values.append(run_id)
    with _connect(settings) as connection:
        connection.execute(f"UPDATE agent_runs SET {', '.join(fields)} WHERE run_id = ?", values)
        connection.commit()


@contextmanager
def record_node(run_id: str, node_name: str, settings: Settings | None = None) -> Iterator[None]:
    node_id = start_node(run_id, node_name, settings)
    try:
        yield
    except Exception as exc:
        finish_node(node_id, "failed", str(exc), settings)
        raise
    finish_node(node_id, "succeeded", None, settings)


def start_node(run_id: str, node_name: str, settings: Settings | None = None) -> int:
    _ensure_schema(settings)
    started_at = _now_iso()
    with _connect(settings) as connection:
        cursor = connection.execute(
            """
            INSERT INTO agent_node_traces (run_id, node_name, status, started_at)
            VALUES (?, ?, 'running', ?)
            """,
            (run_id, node_name, started_at),
        )
        connection.commit()
        return int(cursor.lastrowid)


def finish_node(
    node_id: int,
    status: str,
    error_message: str | None = None,
    settings: Settings | None = None,
) -> None:
    _ensure_schema(settings)
    ended_at = _now_iso()
    with _connect(settings) as connection:
        row = connection.execute("SELECT started_at FROM agent_node_traces WHERE id = ?", (node_id,)).fetchone()
        elapsed_seconds = _elapsed(row[0], ended_at) if row else None
        connection.execute(
            """
            UPDATE agent_node_traces
            SET status = ?, ended_at = ?, elapsed_seconds = ?, error_message = ?
            WHERE id = ?
            """,
            (status, ended_at, elapsed_seconds, _sanitize_text(error_message), node_id),
        )
        connection.commit()


def start_chunk(
    run_id: str,
    *,
    chunk_index: int,
    chunk_start: int,
    chunk_end: int,
    chunk_len: int,
    retry_count: int = 0,
    settings: Settings | None = None,
) -> None:
    _ensure_schema(settings)
    now = _now_iso()
    with _connect(settings) as connection:
        connection.execute(
            """
            INSERT INTO agent_chunk_traces (
                run_id, chunk_index, chunk_start, chunk_end, chunk_len,
                status, issue_count, retry_count, error_message, started_at,
                ended_at, elapsed_seconds
            )
            VALUES (?, ?, ?, ?, ?, 'running', 0, ?, NULL, ?, NULL, NULL)
            ON CONFLICT(run_id, chunk_index) DO UPDATE SET
                chunk_start = excluded.chunk_start,
                chunk_end = excluded.chunk_end,
                chunk_len = excluded.chunk_len,
                status = 'running',
                retry_count = excluded.retry_count,
                error_message = NULL,
                started_at = excluded.started_at,
                ended_at = NULL,
                elapsed_seconds = NULL
            """,
            (run_id, chunk_index, chunk_start, chunk_end, chunk_len, retry_count, now),
        )
        connection.commit()


def finish_chunk(
    run_id: str,
    chunk_index: int,
    *,
    status: str,
    issue_count: int = 0,
    retry_count: int | None = None,
    error_message: str | None = None,
    settings: Settings | None = None,
) -> None:
    _ensure_schema(settings)
    ended_at = _now_iso()
    with _connect(settings) as connection:
        row = connection.execute(
            "SELECT started_at, retry_count FROM agent_chunk_traces WHERE run_id = ? AND chunk_index = ?",
            (run_id, chunk_index),
        ).fetchone()
        elapsed_seconds = _elapsed(row[0], ended_at) if row else None
        resolved_retry_count = retry_count if retry_count is not None else (row[1] if row else 0)
        connection.execute(
            """
            UPDATE agent_chunk_traces
            SET status = ?, issue_count = ?, retry_count = ?, error_message = ?,
                ended_at = ?, elapsed_seconds = ?
            WHERE run_id = ? AND chunk_index = ?
            """,
            (
                status,
                issue_count,
                resolved_retry_count,
                _sanitize_text(error_message),
                ended_at,
                elapsed_seconds,
                run_id,
                chunk_index,
            ),
        )
        connection.commit()


def get_trace(run_id: str, settings: Settings | None = None) -> AgentRunTrace:
    _ensure_schema(settings)
    with _connect(settings) as connection:
        run = connection.execute(
            """
            SELECT run_id, flow, task_id, status, created_at, updated_at,
                   total_chunks, completed_chunks, failed_chunks, issue_count,
                   error_message, metadata_json
            FROM agent_runs
            WHERE run_id = ?
            """,
            (run_id,),
        ).fetchone()
        if not run:
            raise AgentTraceNotFound(run_id)

        nodes = connection.execute(
            """
            SELECT node_name, status, started_at, ended_at, elapsed_seconds, error_message
            FROM agent_node_traces
            WHERE run_id = ?
            ORDER BY id
            """,
            (run_id,),
        ).fetchall()
        chunks = connection.execute(
            """
            SELECT chunk_index, chunk_start, chunk_end, chunk_len, status,
                   issue_count, retry_count, error_message, started_at, ended_at,
                   elapsed_seconds
            FROM agent_chunk_traces
            WHERE run_id = ?
            ORDER BY chunk_index
            """,
            (run_id,),
        ).fetchall()

    return AgentRunTrace(
        run_id=run[0],
        flow=run[1],
        task_id=run[2],
        status=run[3],
        created_at=run[4],
        updated_at=run[5],
        total_chunks=run[6],
        completed_chunks=run[7],
        failed_chunks=run[8],
        issue_count=run[9],
        error_message=run[10],
        metadata=json.loads(run[11] or "{}"),
        nodes=[
            AgentNodeTrace(
                node_name=row[0],
                status=row[1],
                started_at=row[2],
                ended_at=row[3],
                elapsed_seconds=row[4],
                error_message=row[5],
            )
            for row in nodes
        ],
        chunks=[
            AgentChunkTrace(
                chunk_index=row[0],
                chunk_start=row[1],
                chunk_end=row[2],
                chunk_len=row[3],
                status=row[4],
                issue_count=row[5],
                retry_count=row[6],
                error_message=row[7],
                started_at=row[8],
                ended_at=row[9],
                elapsed_seconds=row[10],
            )
            for row in chunks
        ],
    )


def clear_traces_for_tests(settings: Settings | None = None) -> None:
    path = _db_path(settings)
    if path.exists():
        path.unlink()


def trace_dir(settings: Settings | None = None) -> Path:
    resolved_settings = settings or get_settings()
    return resolved_settings.agent_trace_dir


@contextmanager
def _connect(settings: Settings | None = None) -> Iterator[sqlite3.Connection]:
    root = trace_dir(settings)
    root.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(_db_path(settings))
    try:
        yield connection
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def _db_path(settings: Settings | None = None) -> Path:
    return trace_dir(settings) / "traces.sqlite3"


def _ensure_schema(settings: Settings | None = None) -> None:
    with _connect(settings) as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS agent_runs (
                run_id TEXT PRIMARY KEY,
                flow TEXT NOT NULL,
                task_id TEXT,
                status TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                total_chunks INTEGER NOT NULL,
                completed_chunks INTEGER NOT NULL,
                failed_chunks INTEGER NOT NULL,
                issue_count INTEGER NOT NULL,
                error_message TEXT,
                metadata_json TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS agent_node_traces (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id TEXT NOT NULL,
                node_name TEXT NOT NULL,
                status TEXT NOT NULL,
                started_at TEXT NOT NULL,
                ended_at TEXT,
                elapsed_seconds REAL,
                error_message TEXT
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS agent_chunk_traces (
                run_id TEXT NOT NULL,
                chunk_index INTEGER NOT NULL,
                chunk_start INTEGER NOT NULL,
                chunk_end INTEGER NOT NULL,
                chunk_len INTEGER NOT NULL,
                status TEXT NOT NULL,
                issue_count INTEGER NOT NULL,
                retry_count INTEGER NOT NULL,
                error_message TEXT,
                started_at TEXT NOT NULL,
                ended_at TEXT,
                elapsed_seconds REAL,
                PRIMARY KEY (run_id, chunk_index)
            )
            """
        )
        connection.commit()


def _sanitize_metadata(metadata: dict[str, Any]) -> dict[str, Any]:
    sanitized: dict[str, Any] = {}
    for key, value in metadata.items():
        lowered = key.lower()
        if lowered in {"text", "content", "authorization", "api_key", "api-key"}:
            continue
        if isinstance(value, str):
            sanitized[key] = _sanitize_text(value)
        elif isinstance(value, int | float | bool) or value is None:
            sanitized[key] = value
        else:
            sanitized[key] = _sanitize_text(json.dumps(value, ensure_ascii=False, default=str))
    return sanitized


def _sanitize_text(value: str | None) -> str | None:
    if value is None:
        return None
    sanitized = value
    for pattern in _SECRET_PATTERNS:
        sanitized = pattern.sub(r"\1[REDACTED]", sanitized)
    if len(sanitized) > _MAX_ERROR_LENGTH:
        return sanitized[:_MAX_ERROR_LENGTH] + "...[truncated]"
    return sanitized


def _now_iso() -> str:
    return datetime.now(UTC).isoformat()


def _elapsed(started_at: str, ended_at: str) -> float:
    try:
        start = datetime.fromisoformat(started_at)
        end = datetime.fromisoformat(ended_at)
        return round((end - start).total_seconds(), 4)
    except ValueError:
        return round(time.monotonic(), 4)
