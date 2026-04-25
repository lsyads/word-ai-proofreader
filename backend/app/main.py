from __future__ import annotations

import json
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


app = FastAPI(title="Word AI Proofreader", version="0.1.0")
settings = get_settings()

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
    return create_session()


@app.post("/api/proofread", response_model=ProofreadResponse)
async def proofread(request: ProofreadRequest) -> ProofreadResponse:
    try:
        issues = await proofread_service.proofread_text(request.text, session_id=request.session_id)
    except AIClientError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    return ProofreadResponse(issues=issues)


@app.post("/api/proofread/stream")
async def proofread_stream(request: ProofreadRequest) -> StreamingResponse:
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
        async for event in proofread_service.stream_proofread_text(request.text, session_id=request.session_id):
            yield _format_sse(event.event, event.data)
    except AIClientError as exc:
        yield _format_sse("error", {"message": str(exc)})


def _format_sse(event: str, data: dict[str, Any]) -> str:
    padding = ":" + (" " * 2048)
    return f"{padding}\nevent: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"
