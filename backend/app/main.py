from __future__ import annotations

import json
import logging
from collections.abc import AsyncIterator
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from app.schemas import ChunkedProofreadRequest, ChunkedProofreadResult, ProofreadRequest, ProofreadResponse, SessionResponse
import app.services.proofread as proofread_service
from app.services import chunking
from app.services import tasks as task_service
from app.services.ai_client import AIClientError
from app.services.sessions import create_session
from app.settings import get_settings


def configure_logging(level_name: str) -> None:
    level = getattr(logging, level_name.upper(), logging.INFO)
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
    )
    logging.getLogger().setLevel(level)
    for handler in logging.getLogger().handlers:
        handler.setLevel(level)
    logging.getLogger("app").setLevel(level)


app = FastAPI(title="Word AI Proofreader", version="0.1.0")
settings = get_settings()
configure_logging(settings.backend_log_level)
logger = logging.getLogger(__name__)

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.backend_cors_origin_list,
    allow_credentials=False,
    allow_methods=["GET", "POST", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/sessions", response_model=SessionResponse)
async def create_ai_session() -> SessionResponse:
    session = create_session()
    logger.info("created AI session session_id=%s", _mask_session_id(session.session_id))
    return session


@app.post("/api/proofread", response_model=ProofreadResponse)
async def proofread(request: ProofreadRequest) -> ProofreadResponse:
    logger.info(
        "proofread request received text_len=%s session_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s",
        len(request.text),
        _mask_session_id(request.session_id),
        request.provider_api or "default",
        request.proofread_mode,
        request.reasoning_enabled,
    )
    _debug_log_json("proofread request body", request.model_dump())
    try:
        issues = await proofread_service.proofread_text(
            request.text,
            request.book,
            session_id=request.session_id,
            provider_api=request.provider_api,
            proofread_mode=request.proofread_mode,
            reasoning_enabled=request.reasoning_enabled,
        )
    except AIClientError as exc:
        logger.warning(
            "proofread request failed text_len=%s session_id=%s error=%s",
            len(request.text),
            _mask_session_id(request.session_id),
            exc,
            exc_info=True,
        )
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    logger.info(
        "proofread request completed issue_count=%s located_issue_count=%s",
        len(issues),
        _count_located_issues(issues),
    )
    response = ProofreadResponse(issues=issues)
    _debug_log_json("proofread response body", response.model_dump())
    return response


@app.post("/api/proofread/stream")
async def proofread_stream(request: ProofreadRequest) -> StreamingResponse:
    provider_api = proofread_service.resolve_provider_api(request.provider_api)
    if provider_api == "chat":
        raise HTTPException(
            status_code=400,
            detail="Chat mode uses /api/proofread with standard Chat Completions, not SSE.",
        )

    logger.info(
        "proofread stream accepted text_len=%s session_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s",
        len(request.text),
        _mask_session_id(request.session_id),
        provider_api,
        request.proofread_mode,
        request.reasoning_enabled,
    )
    _debug_log_json("proofread stream request body", request.model_dump())
    return StreamingResponse(
        _proofread_event_stream(request),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache, no-transform",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@app.post("/api/proofread/chunked", response_model=ChunkedProofreadResult)
async def proofread_chunked(request: ChunkedProofreadRequest) -> ChunkedProofreadResult:
    logger.info(
        "chunked proofread request received text_len=%s scope=%s chunk_size=%s session_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s",
        len(request.text),
        request.scope,
        request.chunk_size,
        _mask_session_id(request.session_id),
        request.provider_api or "default",
        request.proofread_mode,
        request.reasoning_enabled,
    )
    _debug_log_json("chunked proofread request body", request.model_dump())
    response = await chunking.proofread_chunked(request)
    logger.info(
        "chunked proofread request completed total_chunks=%s completed_chunks=%s failed_chunks=%s issue_count=%s",
        response.total_chunks,
        response.completed_chunks,
        response.failed_chunks,
        len(response.issues),
    )
    _debug_log_json("chunked proofread response body", response.model_dump())
    return response


@app.post("/api/proofread/tasks", response_model=ChunkedProofreadResult)
async def create_proofread_task(request: ChunkedProofreadRequest) -> ChunkedProofreadResult:
    logger.info(
        "proofread task create requested text_len=%s scope=%s chunk_size=%s session_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s",
        len(request.text),
        request.scope,
        request.chunk_size,
        _mask_session_id(request.session_id),
        request.provider_api or "default",
        request.proofread_mode,
        request.reasoning_enabled,
    )
    _debug_log_json("proofread task request body", request.model_dump())
    return task_service.create_task(request)


@app.get("/api/proofread/tasks/{task_id}", response_model=ChunkedProofreadResult)
async def get_proofread_task(task_id: str) -> ChunkedProofreadResult:
    try:
        return task_service.get_task(task_id)
    except task_service.ProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="Proofread task not found") from exc


@app.delete("/api/proofread/tasks/{task_id}", response_model=ChunkedProofreadResult)
async def cancel_proofread_task(task_id: str) -> ChunkedProofreadResult:
    try:
        return task_service.cancel_task(task_id)
    except task_service.ProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="Proofread task not found") from exc


@app.get("/api/proofread/tasks/{task_id}/events")
async def proofread_task_events(task_id: str) -> StreamingResponse:
    try:
        task_service.get_task(task_id)
    except task_service.ProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="Proofread task not found") from exc

    return StreamingResponse(
        _task_event_stream(task_service.stream_task_events(task_id)),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache, no-transform",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


async def _proofread_event_stream(request: ProofreadRequest) -> AsyncIterator[str]:
    try:
        stream_kwargs = {
            "session_id": request.session_id,
            "provider_api": request.provider_api,
            "proofread_mode": request.proofread_mode,
        }
        if request.reasoning_enabled:
            stream_kwargs["reasoning_enabled"] = True

        async for event in proofread_service.stream_proofread_text(
            request.text,
            request.book,
            **stream_kwargs,
        ):
            if event.event == "result":
                issues = event.data.get("issues", [])
                logger.info(
                    "proofread stream result issue_count=%s located_issue_count=%s",
                    len(issues) if isinstance(issues, list) else 0,
                    _count_located_issue_dicts(issues) if isinstance(issues, list) else 0,
                )
                _debug_log_json("proofread stream result body", event.data)
            yield _format_sse(event.event, event.data)
    except AIClientError as exc:
        logger.warning(
            "proofread stream failed text_len=%s session_id=%s error=%s",
            len(request.text),
            _mask_session_id(request.session_id),
            exc,
            exc_info=True,
        )
        error_data = {"message": str(exc)}
        _debug_log_json("proofread stream error body", error_data)
        yield _format_sse("error", error_data)


async def _task_event_stream(events: AsyncIterator[Any]) -> AsyncIterator[str]:
    async for event in events:
        yield _format_sse(event.event, event.data)


def _format_sse(event: str, data: dict[str, Any]) -> str:
    padding = ":" + (" " * 2048)
    return f"{padding}\nevent: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"


def _mask_session_id(session_id: str | None) -> str:
    if not session_id:
        return "none"

    return f"...{session_id[-8:]}"


def _count_located_issues(issues: list[Any]) -> int:
    return sum(
        1
        for issue in issues
        if getattr(issue, "start", None) is not None and getattr(issue, "end", None) is not None
    )


def _count_located_issue_dicts(issues: list[Any]) -> int:
    return sum(
        1
        for issue in issues
        if isinstance(issue, dict) and issue.get("start") is not None and issue.get("end") is not None
    )


def _debug_log_json(message: str, payload: Any) -> None:
    if logger.isEnabledFor(logging.DEBUG):
        logger.debug("%s: %s", message, json.dumps(payload, ensure_ascii=False, default=str))
