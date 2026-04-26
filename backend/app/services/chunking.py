from __future__ import annotations

import logging
from collections.abc import Awaitable, Callable

from app.schemas import (
    BookInfo,
    ChunkedProofreadIssue,
    ChunkedProofreadRequest,
    ChunkedProofreadResult,
    ChunkedTaskStatus,
    ProofreadChunk,
    ProofreadIssue,
    ProofreadScope,
)
from app.services import proofread as proofread_service

DEFAULT_CHUNK_SIZE = 3000
SELECTION_CHUNK_THRESHOLD = 5000
MIN_CHUNK_SIZE = 500
BOUNDARY_LOOKBACK = 1500
PARAGRAPH_BOUNDARIES = ("\n\n", "\r\n\r\n", "\n", "\r")
SENTENCE_BOUNDARIES = ("。", "！", "？", "；", ".", "!", "?", ";")

ProofreadChunkCallable = Callable[
    [
        str,
        BookInfo,
        str | None,
        proofread_service.ProviderAPI | None,
        proofread_service.ProofreadMode,
        bool,
    ],
    Awaitable[list[ProofreadIssue]],
]

logger = logging.getLogger(__name__)


def should_use_chunked_flow(text: str, scope: ProofreadScope) -> bool:
    return scope == "document" or len(text) > SELECTION_CHUNK_THRESHOLD


def split_text_into_chunks(
    text: str,
    scope: ProofreadScope = "selection",
    chunk_size: int = DEFAULT_CHUNK_SIZE,
) -> list[ProofreadChunk]:
    normalized_size = max(MIN_CHUNK_SIZE, chunk_size)

    if not should_use_chunked_flow(text, scope):
        return [ProofreadChunk(index=0, start=0, end=len(text), text=text)]

    chunks: list[ProofreadChunk] = []
    start = 0

    while start < len(text):
        hard_end = min(start + normalized_size, len(text))
        end = len(text) if hard_end == len(text) else _find_chunk_boundary(text, start, hard_end)
        chunk_text = text[start:end]

        if chunk_text.strip():
            chunks.append(
                ProofreadChunk(
                    index=len(chunks),
                    start=start,
                    end=end,
                    text=chunk_text,
                )
            )

        start = end

    return chunks or [ProofreadChunk(index=0, start=0, end=len(text), text=text)]


async def proofread_chunked(
    request: ChunkedProofreadRequest,
    task_id: str | None = None,
    status: ChunkedTaskStatus = "succeeded",
) -> ChunkedProofreadResult:
    issues, completed_chunks, failed_chunks = await proofread_chunks(
        request,
        proofread_chunk=proofread_service.proofread_text,
    )
    result_status = status
    error_message = None
    if completed_chunks == 0 and failed_chunks > 0:
        result_status = "failed"
        error_message = "All chunks failed to proofread."
    elif completed_chunks > 0 and failed_chunks > 0:
        result_status = "partial_succeeded"

    return ChunkedProofreadResult(
        task_id=task_id,
        scope=request.scope,
        status=result_status,
        total_chunks=len(split_text_into_chunks(request.text, request.scope, request.chunk_size)),
        completed_chunks=completed_chunks,
        failed_chunks=failed_chunks,
        issues=issues,
        error_message=error_message,
    )


async def proofread_chunks(
    request: ChunkedProofreadRequest,
    proofread_chunk: ProofreadChunkCallable,
) -> tuple[list[ChunkedProofreadIssue], int, int]:
    chunks = split_text_into_chunks(request.text, request.scope, request.chunk_size)
    aggregated: list[ChunkedProofreadIssue] = []
    completed_chunks = 0
    failed_chunks = 0

    for chunk in chunks:
        try:
            if request.reasoning_enabled:
                chunk_issues = await proofread_chunk(
                    chunk.text,
                    request.book,
                    request.session_id,
                    request.provider_api,
                    request.proofread_mode,
                    True,
                )
            else:
                chunk_issues = await proofread_chunk(
                    chunk.text,
                    request.book,
                    request.session_id,
                    request.provider_api,
                    request.proofread_mode,
                )
        except Exception:
            failed_chunks += 1
            logger.exception("chunk proofread failed chunk_index=%s", chunk.index)
            continue

        completed_chunks += 1
        aggregated.extend(globalize_issues(chunk, chunk_issues))

    return aggregated, completed_chunks, failed_chunks


def globalize_issues(chunk: ProofreadChunk, issues: list[ProofreadIssue]) -> list[ChunkedProofreadIssue]:
    return [
        ChunkedProofreadIssue(
            **issue.model_dump(),
            chunk_index=chunk.index,
            global_start=chunk.start + issue.start if issue.start is not None else None,
            global_end=chunk.start + issue.end if issue.end is not None else None,
        )
        for issue in issues
    ]


def _find_chunk_boundary(text: str, start: int, hard_end: int) -> int:
    window_start = max(start + MIN_CHUNK_SIZE, hard_end - BOUNDARY_LOOKBACK)
    search_window = text[window_start:hard_end]

    for boundary in PARAGRAPH_BOUNDARIES:
        relative = search_window.rfind(boundary)
        if relative != -1:
            return window_start + relative + len(boundary)

    best_sentence = -1
    for boundary in SENTENCE_BOUNDARIES:
        relative = search_window.rfind(boundary)
        if relative > best_sentence:
            best_sentence = relative + len(boundary)

    if best_sentence != -1:
        return window_start + best_sentence

    return hard_end
