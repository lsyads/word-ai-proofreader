from __future__ import annotations

import asyncio
import logging
import tempfile
import time
import uuid
from collections import OrderedDict
from collections.abc import AsyncIterator
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from app.schemas import BookInfo, ChunkedProofreadIssue, ChunkedTaskStatus, DocxProofreadResult
from app.services import docx as docx_service
from app.services.ai_client import AIClientError, AIStreamEvent
from app.services.proofread import ProofreadMode, ProviderAPI, proofread_text

MAX_TASKS = 30
HEARTBEAT_INTERVAL_SECONDS = 15.0
OUTPUT_DIR = Path(tempfile.gettempdir()) / "word-ai-proofreader-docx"
logger = logging.getLogger(__name__)


class DocxProofreadTaskNotFound(KeyError):
    """Raised when an in-memory DOCX task is unavailable."""


class DocxProofreadTaskConflict(RuntimeError):
    """Raised when a DOCX task cannot accept the requested operation now."""


@dataclass
class DocxProofreadRequestData:
    filename: str
    content: bytes
    book: BookInfo
    session_id: str | None = None
    provider_api: ProviderAPI | None = None
    proofread_mode: ProofreadMode = "fast"
    reasoning_enabled: bool = False
    application_mode: docx_service.ApplicationMode = "comment"


@dataclass
class DocxProofreadTask:
    task_id: str
    request: DocxProofreadRequestData
    status: ChunkedTaskStatus = "queued"
    total_chunks: int = 0
    completed_chunks: int = 0
    failed_chunks: int = 0
    issue_count: int = 0
    error_message: str | None = None
    output_filename: str | None = None
    output_path: Path | None = None
    cancel_requested: bool = False
    current_chunk_index: int | None = None
    current_chunk_started_at: float | None = None
    retry_current_requested: bool = False
    runner_task: asyncio.Task[Any] | None = None
    completed_chunk_indices: set[int] = field(default_factory=set)
    failed_chunk_indices: set[int] = field(default_factory=set)
    issues: list[ChunkedProofreadIssue] = field(default_factory=list)
    events: list[AIStreamEvent] = field(default_factory=list)
    subscribers: list[asyncio.Queue[AIStreamEvent]] = field(default_factory=list)
    created_at: str = field(default_factory=lambda: _now_iso())
    updated_at: str = field(default_factory=lambda: _now_iso())


_tasks: OrderedDict[str, DocxProofreadTask] = OrderedDict()


def create_task(request: DocxProofreadRequestData) -> DocxProofreadResult:
    document = docx_service.parse_docx(request.content)
    chunks = docx_service.split_docx_into_chunks(document)
    task_id = f"docx_task_{uuid.uuid4().hex}"
    task = DocxProofreadTask(task_id=task_id, request=request, total_chunks=len(chunks))
    _tasks[task_id] = task
    _trim_tasks()
    _emit(
        task,
        "queued",
        {
            "task_id": task_id,
            "scope": "document",
            "total_chunks": task.total_chunks,
            "message": "DOCX 全书审校任务已创建。",
        },
    )
    task.runner_task = asyncio.create_task(_run_task(task, chunks))
    return snapshot_task(task)


def get_task(task_id: str) -> DocxProofreadResult:
    return snapshot_task(_require_task(task_id))


def cancel_task(task_id: str) -> DocxProofreadResult:
    task = _require_task(task_id)
    if task.status in {"queued", "running"}:
        task.cancel_requested = True
        _touch(task)
    return snapshot_task(task)


def retry_current_chunk(task_id: str) -> DocxProofreadResult:
    task = _require_task(task_id)
    if task.status != "running" or task.current_chunk_index is None or not task.runner_task:
        raise DocxProofreadTaskConflict("No running DOCX chunk is available to retry.")

    task.retry_current_requested = True
    task.runner_task.cancel()
    _touch(task)
    _emit(
        task,
        "chunk_retry_requested",
        _progress_payload(
            task,
            f"已请求重试第 {task.current_chunk_index + 1}/{task.total_chunks} 块。",
            extra={"chunk_index": task.current_chunk_index},
        ),
    )
    return snapshot_task(task)


def retry_failed_chunks(task_id: str) -> DocxProofreadResult:
    task = _require_task(task_id)
    if task.status in {"queued", "running"}:
        raise DocxProofreadTaskConflict("DOCX proofread task is still running.")

    failed_indices = sorted(task.failed_chunk_indices)
    if not failed_indices:
        raise DocxProofreadTaskConflict("DOCX proofread task has no failed chunks to retry.")

    document = docx_service.parse_docx(task.request.content)
    chunks = docx_service.split_docx_into_chunks(document)
    retry_chunks = [chunks[index] for index in failed_indices if index < len(chunks)]
    if not retry_chunks:
        raise DocxProofreadTaskConflict("Failed DOCX chunks are no longer available to retry.")

    task.cancel_requested = False
    task.error_message = None
    task.output_filename = None
    task.output_path = None
    task.status = "queued"
    _touch(task)
    _emit(
        task,
        "retry_queued",
        _progress_payload(
            task,
            f"已创建 DOCX 失败分块重试任务，共 {len(retry_chunks)} 块。",
            extra={"retry_chunk_indices": failed_indices},
        ),
    )
    task.runner_task = asyncio.create_task(_run_task(task, retry_chunks, retry_only=True))
    return snapshot_task(task)


def get_download_path(task_id: str) -> tuple[Path, str]:
    task = _require_task(task_id)
    if not task.output_path or not task.output_filename or not task.output_path.exists():
        raise DocxProofreadTaskConflict("DOCX proofread result is not available for download.")
    return task.output_path, task.output_filename


async def stream_task_events(task_id: str) -> AsyncIterator[AIStreamEvent]:
    task = _require_task(task_id)
    queue: asyncio.Queue[AIStreamEvent] = asyncio.Queue()

    for event in task.events:
        yield event

    if _is_terminal(task.status):
        return

    task.subscribers.append(queue)
    try:
        while True:
            if _is_terminal(task.status) and queue.empty():
                return

            event = await queue.get()
            yield event
            if event.event in {"completed", "cancelled", "error"}:
                return
    finally:
        if queue in task.subscribers:
            task.subscribers.remove(queue)


def clear_tasks_for_tests() -> None:
    _tasks.clear()


async def _run_task(
    task: DocxProofreadTask,
    chunks: list[Any],
    retry_only: bool = False,
) -> None:
    task.status = "running"
    task.runner_task = asyncio.current_task()
    _touch(task)
    _emit(
        task,
        "running",
        _progress_payload(task, "正在重试 DOCX 失败分块。" if retry_only else "DOCX 审校任务开始运行。"),
    )

    try:
        for chunk in chunks:
            if task.cancel_requested:
                _mark_cancelled(task)
                return

            try:
                chunk_issues = await _proofread_chunk_with_manual_retry(task, chunk)
            except AIClientError as exc:
                task.failed_chunk_indices.add(chunk.index)
                _sync_chunk_counts(task)
                _touch(task)
                elapsed_seconds = _elapsed_seconds_for_task(task)
                logger.warning(
                    "docx chunk proofread failed task_id=%s chunk_index=%s error=%s",
                    task.task_id,
                    chunk.index,
                    exc,
                    exc_info=True,
                )
                _emit(
                    task,
                    "chunk_failed",
                    _progress_payload(
                        task,
                        f"第 {chunk.index + 1} 块审校失败，已继续后续分块。",
                        extra={
                            "chunk_index": chunk.index,
                            "chunk_start": chunk.start,
                            "chunk_end": chunk.end,
                            "chunk_len": len(chunk.text),
                            "elapsed_seconds": elapsed_seconds,
                            "error_message": str(exc),
                        },
                    ),
                )
                continue

            task.failed_chunk_indices.discard(chunk.index)
            task.completed_chunk_indices.add(chunk.index)
            _sync_chunk_counts(task)
            task.issues = [issue for issue in task.issues if issue.chunk_index != chunk.index]
            task.issues.extend(_globalize_issues(chunk, chunk_issues))
            task.issue_count = len(task.issues)
            _touch(task)
            elapsed_seconds = _elapsed_seconds_for_task(task)
            _emit(
                task,
                "chunk_completed",
                _progress_payload(
                    task,
                    f"第 {chunk.index + 1} 块审校完成。",
                    extra={
                        "chunk_index": chunk.index,
                        "chunk_start": chunk.start,
                        "chunk_end": chunk.end,
                        "chunk_len": len(chunk.text),
                        "elapsed_seconds": elapsed_seconds,
                        "issue_count": task.issue_count,
                    },
                ),
            )

        if task.cancel_requested:
            _mark_cancelled(task)
            return

        if task.completed_chunks == 0 and task.failed_chunks > 0:
            task.status = "failed"
            task.error_message = "All DOCX chunks failed to proofread."
            _touch(task)
            _emit(task, "error", _progress_payload(task, "DOCX 审校任务失败。"))
            return

        _save_output_file(task)
        task.status = "partial_succeeded" if task.failed_chunks > 0 else "succeeded"
        _touch(task)
        message = "DOCX 审校任务部分完成，已生成可下载文件。" if task.failed_chunks else "DOCX 审校任务完成，已生成可下载文件。"
        _emit(task, "completed", _progress_payload(task, message))
    except Exception as exc:
        task.status = "failed"
        task.error_message = str(exc)
        _touch(task)
        logger.exception("docx proofread task crashed task_id=%s", task.task_id)
        _emit(task, "error", _progress_payload(task, "DOCX 审校任务失败。", extra={"error_message": str(exc)}))
    finally:
        task.current_chunk_index = None
        task.current_chunk_started_at = None
        task.retry_current_requested = False
        if task.runner_task is asyncio.current_task():
            task.runner_task = None


async def _proofread_chunk_with_manual_retry(task: DocxProofreadTask, chunk: Any) -> list[Any]:
    while True:
        chunk_started_at = time.monotonic()
        task.current_chunk_index = chunk.index
        task.current_chunk_started_at = chunk_started_at
        _emit(
            task,
            "chunk_started",
            _progress_payload(
                task,
                f"正在审校第 {chunk.index + 1}/{task.total_chunks} 块。",
                extra={
                    "chunk_index": chunk.index,
                    "chunk_start": chunk.start,
                    "chunk_end": chunk.end,
                    "chunk_len": len(chunk.text),
                    "elapsed_seconds": 0,
                },
            ),
        )

        heartbeat_task = asyncio.create_task(_emit_heartbeats(task, chunk, chunk_started_at))
        try:
            return await _proofread_chunk(task, chunk)
        except asyncio.CancelledError:
            if task.retry_current_requested:
                task.retry_current_requested = False
                _emit(
                    task,
                    "chunk_retrying",
                    _progress_payload(
                        task,
                        f"正在重新审校第 {chunk.index + 1}/{task.total_chunks} 块。",
                        extra={
                            "chunk_index": chunk.index,
                            "chunk_start": chunk.start,
                            "chunk_end": chunk.end,
                            "chunk_len": len(chunk.text),
                            "elapsed_seconds": _elapsed_seconds(chunk_started_at),
                        },
                    ),
                )
                continue
            raise
        finally:
            heartbeat_task.cancel()
            try:
                await heartbeat_task
            except asyncio.CancelledError:
                pass


async def _proofread_chunk(task: DocxProofreadTask, chunk: Any) -> list[Any]:
    kwargs = {
        "session_id": task.request.session_id,
        "provider_api": task.request.provider_api,
        "proofread_mode": task.request.proofread_mode,
    }
    if task.request.reasoning_enabled:
        kwargs["reasoning_enabled"] = True
    return await proofread_text(chunk.text, task.request.book, **kwargs)


def _globalize_issues(chunk: Any, issues: list[Any]) -> list[ChunkedProofreadIssue]:
    global_issues: list[ChunkedProofreadIssue] = []
    for issue in issues:
        payload = issue.model_dump()
        payload["locator"] = None
        global_issues.append(
            ChunkedProofreadIssue(
                **payload,
                chunk_index=chunk.index,
                global_start=chunk.start + issue.start if issue.start is not None else None,
                global_end=chunk.start + issue.end if issue.end is not None else None,
            )
        )
    return global_issues


def _save_output_file(task: DocxProofreadTask) -> None:
    output_filename = docx_service.build_output_filename(task.request.filename, task.request.application_mode)
    output_path = OUTPUT_DIR / task.task_id / output_filename
    docx_service.write_docx_result(
        task.request.content,
        task.issues,
        task.request.application_mode,
        output_path,
    )
    task.output_filename = output_filename
    task.output_path = output_path


def snapshot_task(task: DocxProofreadTask) -> DocxProofreadResult:
    return DocxProofreadResult(
        task_id=task.task_id,
        status=task.status,
        total_chunks=task.total_chunks,
        completed_chunks=task.completed_chunks,
        failed_chunks=task.failed_chunks,
        issue_count=task.issue_count,
        source_filename=task.request.filename,
        application_mode=task.request.application_mode,
        output_filename=task.output_filename,
        download_url=f"/api/proofread/docx/tasks/{task.task_id}/download" if task.output_filename else None,
        error_message=task.error_message,
    )


def _require_task(task_id: str) -> DocxProofreadTask:
    try:
        return _tasks[task_id]
    except KeyError as exc:
        raise DocxProofreadTaskNotFound(task_id) from exc


def _emit(task: DocxProofreadTask, event: str, data: dict[str, Any]) -> None:
    stream_event = AIStreamEvent(event, data)
    task.events.append(stream_event)
    for queue in list(task.subscribers):
        queue.put_nowait(stream_event)


def _progress_payload(
    task: DocxProofreadTask,
    message: str,
    extra: dict[str, Any] | None = None,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "task_id": task.task_id,
        "scope": "document",
        "status": task.status,
        "total_chunks": task.total_chunks,
        "completed_chunks": task.completed_chunks,
        "failed_chunks": task.failed_chunks,
        "issue_count": task.issue_count,
        "source_filename": task.request.filename,
        "output_filename": task.output_filename,
        "download_url": f"/api/proofread/docx/tasks/{task.task_id}/download" if task.output_filename else None,
        "message": message,
    }
    if extra:
        payload.update(extra)
    return payload


def _mark_cancelled(task: DocxProofreadTask) -> None:
    task.status = "cancelled"
    _touch(task)
    _emit(task, "cancelled", _progress_payload(task, "DOCX 审校任务已停止。"))


async def _emit_heartbeats(task: DocxProofreadTask, chunk: Any, chunk_started_at: float) -> None:
    while True:
        await asyncio.sleep(HEARTBEAT_INTERVAL_SECONDS)
        if task.cancel_requested or _is_terminal(task.status):
            return
        _emit(
            task,
            "heartbeat",
            _progress_payload(
                task,
                f"第 {chunk.index + 1}/{task.total_chunks} 块仍在审校。",
                extra={
                    "chunk_index": chunk.index,
                    "chunk_start": chunk.start,
                    "chunk_end": chunk.end,
                    "chunk_len": len(chunk.text),
                    "elapsed_seconds": _elapsed_seconds(chunk_started_at),
                },
            ),
        )


def _touch(task: DocxProofreadTask) -> None:
    task.updated_at = _now_iso()


def _sync_chunk_counts(task: DocxProofreadTask) -> None:
    task.completed_chunks = len(task.completed_chunk_indices)
    task.failed_chunks = len(task.failed_chunk_indices)


def _now_iso() -> str:
    return datetime.now(UTC).isoformat()


def _elapsed_seconds(started_at: float) -> float:
    return round(time.monotonic() - started_at, 1)


def _elapsed_seconds_for_task(task: DocxProofreadTask) -> float:
    if task.current_chunk_started_at is not None:
        return _elapsed_seconds(task.current_chunk_started_at)
    return 0.0


def _is_terminal(status: ChunkedTaskStatus) -> bool:
    return status in {"succeeded", "partial_succeeded", "failed", "cancelled"}


def _trim_tasks() -> None:
    while len(_tasks) > MAX_TASKS:
        _tasks.popitem(last=False)
