from __future__ import annotations

import asyncio
import uuid
from collections import OrderedDict
from collections.abc import AsyncIterator
from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import Any

from app.schemas import (
    ChunkedProofreadIssue,
    ChunkedProofreadRequest,
    ChunkedProofreadResult,
    ChunkedTaskStatus,
)
from app.services import chunking
from app.services.ai_client import AIClientError, AIStreamEvent
from app.services.proofread import proofread_text

MAX_TASKS = 50
HEARTBEAT_INTERVAL_SECONDS = 15.0


class ProofreadTaskNotFound(KeyError):
    """Raised when an in-memory proofread task is unavailable."""


@dataclass
class ProofreadTask:
    task_id: str
    request: ChunkedProofreadRequest
    status: ChunkedTaskStatus = "queued"
    total_chunks: int = 0
    completed_chunks: int = 0
    failed_chunks: int = 0
    issues: list[ChunkedProofreadIssue] = field(default_factory=list)
    error_message: str | None = None
    cancel_requested: bool = False
    created_at: str = field(default_factory=lambda: _now_iso())
    updated_at: str = field(default_factory=lambda: _now_iso())
    events: list[AIStreamEvent] = field(default_factory=list)
    subscribers: list[asyncio.Queue[AIStreamEvent]] = field(default_factory=list)


_tasks: OrderedDict[str, ProofreadTask] = OrderedDict()


def create_task(request: ChunkedProofreadRequest) -> ChunkedProofreadResult:
    chunks = chunking.split_text_into_chunks(request.text, request.scope, request.chunk_size)
    task_id = f"task_{uuid.uuid4().hex}"
    task = ProofreadTask(
        task_id=task_id,
        request=request,
        total_chunks=len(chunks),
    )
    _tasks[task_id] = task
    _trim_tasks()
    _emit(
        task,
        "queued",
        {
            "task_id": task_id,
            "scope": request.scope,
            "total_chunks": task.total_chunks,
            "message": "审校任务已创建。",
        },
    )
    asyncio.create_task(_run_task(task, chunks))
    return snapshot_task(task)


def get_task(task_id: str) -> ChunkedProofreadResult:
    return snapshot_task(_require_task(task_id))


def cancel_task(task_id: str) -> ChunkedProofreadResult:
    task = _require_task(task_id)

    if task.status in {"queued", "running"}:
        task.cancel_requested = True
        task.updated_at = _now_iso()

    return snapshot_task(task)


async def stream_task_events(task_id: str) -> AsyncIterator[AIStreamEvent]:
    task = _require_task(task_id)
    queue: asyncio.Queue[AIStreamEvent] = asyncio.Queue()
    replay_from = len(task.events)

    for event in task.events:
        yield event

    if _is_terminal(task.status):
        return

    task.subscribers.append(queue)

    try:
        while True:
            if _is_terminal(task.status) and queue.empty() and replay_from <= len(task.events):
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


async def _run_task(task: ProofreadTask, chunks: list[chunking.ProofreadChunk]) -> None:
    task.status = "running"
    _touch(task)
    _emit(task, "running", _progress_payload(task, "审校任务开始运行。"))

    try:
        for chunk in chunks:
            if task.cancel_requested:
                _mark_cancelled(task)
                return

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
                    },
                ),
            )

            heartbeat_task = asyncio.create_task(_emit_heartbeats(task, chunk))
            try:
                chunk_issues = await proofread_text(
                    chunk.text,
                    task.request.book,
                    session_id=task.request.session_id,
                    provider_api=task.request.provider_api,
                    proofread_mode=task.request.proofread_mode,
                )
            except AIClientError as exc:
                task.failed_chunks += 1
                _touch(task)
                _emit(
                    task,
                    "chunk_failed",
                    _progress_payload(
                        task,
                        f"第 {chunk.index + 1} 块审校失败，已继续后续分块。",
                        extra={"chunk_index": chunk.index, "error_message": str(exc)},
                    ),
                )
                continue
            finally:
                heartbeat_task.cancel()
                try:
                    await heartbeat_task
                except asyncio.CancelledError:
                    pass

            task.completed_chunks += 1
            task.issues.extend(chunking.globalize_issues(chunk, chunk_issues))
            _touch(task)
            _emit(
                task,
                "chunk_completed",
                _progress_payload(
                    task,
                    f"第 {chunk.index + 1} 块审校完成。",
                    extra={"chunk_index": chunk.index, "issue_count": len(task.issues)},
                ),
            )

        if task.cancel_requested:
            _mark_cancelled(task)
            return

        if task.completed_chunks == 0 and task.failed_chunks > 0:
            task.status = "failed"
            task.error_message = "All chunks failed to proofread."
            _touch(task)
            _emit(task, "error", _progress_payload(task, "审校任务失败。"))
            return

        task.status = "partial_succeeded" if task.failed_chunks > 0 else "succeeded"
        _touch(task)
        message = "审校任务部分完成。" if task.status == "partial_succeeded" else "审校任务完成。"
        _emit(task, "completed", _progress_payload(task, message))
    except Exception as exc:
        task.status = "failed"
        task.error_message = str(exc)
        _touch(task)
        _emit(task, "error", _progress_payload(task, "审校任务失败。", extra={"error_message": str(exc)}))


def snapshot_task(task: ProofreadTask) -> ChunkedProofreadResult:
    return ChunkedProofreadResult(
        task_id=task.task_id,
        scope=task.request.scope,
        status=task.status,
        total_chunks=task.total_chunks,
        completed_chunks=task.completed_chunks,
        failed_chunks=task.failed_chunks,
        issues=task.issues,
        error_message=task.error_message,
    )


def _require_task(task_id: str) -> ProofreadTask:
    try:
        return _tasks[task_id]
    except KeyError as exc:
        raise ProofreadTaskNotFound(task_id) from exc


def _emit(task: ProofreadTask, event: str, data: dict[str, Any]) -> None:
    stream_event = AIStreamEvent(event, data)
    task.events.append(stream_event)

    for queue in list(task.subscribers):
        queue.put_nowait(stream_event)


def _progress_payload(
    task: ProofreadTask,
    message: str,
    extra: dict[str, Any] | None = None,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "task_id": task.task_id,
        "scope": task.request.scope,
        "status": task.status,
        "total_chunks": task.total_chunks,
        "completed_chunks": task.completed_chunks,
        "failed_chunks": task.failed_chunks,
        "issue_count": len(task.issues),
        "message": message,
    }

    if extra:
        payload.update(extra)

    return payload


def _mark_cancelled(task: ProofreadTask) -> None:
    task.status = "cancelled"
    _touch(task)
    _emit(task, "cancelled", _progress_payload(task, "审校任务已停止。"))


async def _emit_heartbeats(task: ProofreadTask, chunk: chunking.ProofreadChunk) -> None:
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
                },
            ),
        )


def _touch(task: ProofreadTask) -> None:
    task.updated_at = _now_iso()


def _now_iso() -> str:
    return datetime.now(UTC).isoformat()


def _is_terminal(status: ChunkedTaskStatus) -> bool:
    return status in {"succeeded", "partial_succeeded", "failed", "cancelled"}


def _trim_tasks() -> None:
    while len(_tasks) > MAX_TASKS:
        _tasks.popitem(last=False)
