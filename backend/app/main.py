from __future__ import annotations

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware

from app.schemas import ProofreadRequest, ProofreadResponse
from app.services.ai_client import AIClientError
from app.services.proofread import proofread_text


app = FastAPI(title="Word AI Proofreader", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000", "http://localhost:3000"],
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/proofread", response_model=ProofreadResponse)
async def proofread(request: ProofreadRequest) -> ProofreadResponse:
    try:
        issues = await proofread_text(request.text)
    except AIClientError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    return ProofreadResponse(issues=issues)
