from __future__ import annotations

from langchain_core.runnables import RunnableLambda

from app.schemas import BookInfo, ProofreadIssue
from app.services import proofread
from app.services.ai_client import DEFAULT_TEMPERATURE
from app.services.proofread import ProofreadMode, ProviderAPI


async def proofread_text_with_provider(
    text: str,
    book: BookInfo,
    *,
    session_id: str | None = None,
    ai_profile_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
    temperature: float = DEFAULT_TEMPERATURE,
) -> list[ProofreadIssue]:
    """Call the configured AI provider and return located proofreading issues."""
    return await proofread.proofread_text(
        text,
        book,
        session_id=session_id,
        ai_profile_id=ai_profile_id,
        provider_api=provider_api,
        proofread_mode=proofread_mode,
        reasoning_enabled=reasoning_enabled,
        temperature=temperature,
    )


async def _proofread_model(payload: dict) -> list[ProofreadIssue]:
    return await proofread_text_with_provider(
        payload["text"],
        payload["book"],
        session_id=payload.get("session_id"),
        ai_profile_id=payload.get("ai_profile_id"),
        provider_api=payload.get("provider_api"),
        proofread_mode=payload.get("proofread_mode", "fast"),
        reasoning_enabled=payload.get("reasoning_enabled", False),
        temperature=payload.get("temperature", DEFAULT_TEMPERATURE),
    )


proofread_model_runnable = RunnableLambda(_proofread_model)
