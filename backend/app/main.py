from __future__ import annotations

import json
import logging
from collections.abc import AsyncIterator
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from app.schemas import ProofreadRequest, ProofreadResponse, SessionResponse
import app.services.proofread as proofread_service
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
    allow_methods=["GET", "POST", "OPTIONS"],
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
        "proofread request received text_len=%s session_id=%s provider_api=%s proofread_mode=%s",
        len(request.text),
        _mask_session_id(request.session_id),
        request.provider_api or "default",
        request.proofread_mode,
    )
    try:
        issues = await proofread_service.proofread_text(
            request.text,
            session_id=request.session_id,
            provider_api=request.provider_api,
            proofread_mode=request.proofread_mode,
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
    return ProofreadResponse(issues=issues)


@app.post("/api/proofread/stream")
async def proofread_stream(request: ProofreadRequest) -> StreamingResponse:
    logger.info(
        "proofread stream accepted text_len=%s session_id=%s provider_api=%s proofread_mode=%s",
        len(request.text),
        _mask_session_id(request.session_id),
        request.provider_api or "default",
        request.proofread_mode,
    )
    return StreamingResponse(
        _proofread_event_stream(request),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache, no-transform",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


async def _proofread_event_stream(request: ProofreadRequest) -> AsyncIterator[str]:
    try:
        async for event in proofread_service.stream_proofread_text(
            request.text,
            session_id=request.session_id,
            provider_api=request.provider_api,
            proofread_mode=request.proofread_mode,
        ):
            if event.event == "result":
                issues = event.data.get("issues", [])
                logger.info(
                    "proofread stream result issue_count=%s located_issue_count=%s",
                    len(issues) if isinstance(issues, list) else 0,
                    _count_located_issue_dicts(issues) if isinstance(issues, list) else 0,
                )
            yield _format_sse(event.event, event.data)
    except AIClientError as exc:
        logger.warning(
            "proofread stream failed text_len=%s session_id=%s error=%s",
            len(request.text),
            _mask_session_id(request.session_id),
            exc,
            exc_info=True,
        )
        yield _format_sse("error", {"message": str(exc)})


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
