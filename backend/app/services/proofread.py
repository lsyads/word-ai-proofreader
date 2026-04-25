from __future__ import annotations

import logging
from collections.abc import AsyncIterator
from typing import Literal

from app.schemas import ProofreadIssue
from app.services.ai_client import AIClientError, AIStreamEvent, proofread_with_ai, stream_proofread_with_ai
from app.services.sessions import SessionNotFoundError, get_last_response_id, update_last_response_id
from app.settings import get_settings

ProviderAPI = Literal["responses", "chat"]
ProofreadMode = Literal["fast", "thinking"]
logger = logging.getLogger(__name__)


async def proofread_text(
    text: str,
    session_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
) -> list[ProofreadIssue]:
    settings = get_settings()
    provider_api = resolve_provider_api(provider_api)
    logger.info(
        "proofread service started text_len=%s provider_api=%s proofread_mode=%s has_api_key=%s session_id=%s",
        len(text),
        provider_api,
        proofread_mode,
        bool(settings.ai_api_key),
        _mask_session_id(session_id),
    )

    if settings.ai_api_key:
        previous_response_id = None
        if provider_api == "responses":
            try:
                previous_response_id = get_last_response_id(session_id)
            except SessionNotFoundError as exc:
                logger.warning("proofread service session not found session_id=%s", _mask_session_id(session_id))
                raise AIClientError("AI session was not found. Please create a new session.") from exc

        logger.info(
            "proofread service calling AI provider provider_api=%s proofread_mode=%s has_previous_response=%s",
            provider_api,
            proofread_mode,
            bool(previous_response_id),
        )
        result = await proofread_with_ai(
            text,
            previous_response_id=previous_response_id,
            provider_api=provider_api,
            proofread_mode=proofread_mode,
        )
        if provider_api == "responses":
            update_last_response_id(session_id, result.response_id)
            logger.info(
                "proofread service updated session response_id_present=%s session_id=%s",
                bool(result.response_id),
                _mask_session_id(session_id),
            )
        return locate_issues(text, result.issues)

    logger.info("proofread service using mock issues")
    return locate_issues(text, build_mock_issues(text))


async def stream_proofread_text(
    text: str,
    session_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
) -> AsyncIterator[AIStreamEvent]:
    settings = get_settings()
    provider_api = resolve_provider_api(provider_api)
    logger.info(
        "proofread stream service started text_len=%s provider_api=%s proofread_mode=%s has_api_key=%s session_id=%s",
        len(text),
        provider_api,
        proofread_mode,
        bool(settings.ai_api_key),
        _mask_session_id(session_id),
    )

    yield AIStreamEvent("status", {"stage": "received", "message": "已接收选区文本。"})

    if not settings.ai_api_key:
        logger.info("proofread stream service using mock issues")
        yield AIStreamEvent("status", {"stage": "calling_ai", "message": "当前未配置 AI_API_KEY，正在返回本地 mock 审校结果。"})
        yield AIStreamEvent("status", {"stage": "normalizing", "message": "正在整理结构化审校结果。"})
        yield AIStreamEvent("result", {"issues": [issue.model_dump() for issue in locate_issues(text, build_mock_issues(text))]})
        yield AIStreamEvent("status", {"stage": "completed", "message": "审校完成。"})
        return

    previous_response_id = None
    if provider_api == "responses":
        try:
            previous_response_id = get_last_response_id(session_id)
        except SessionNotFoundError as exc:
            logger.warning("proofread stream service session not found session_id=%s", _mask_session_id(session_id))
            raise AIClientError("AI session was not found. Please create a new session.") from exc

    if provider_api == "responses":
        yield AIStreamEvent("status", {"stage": "calling_ai", "message": "正在调用 AI 原生 Responses session。"})
    else:
        yield AIStreamEvent("status", {"stage": "calling_ai", "message": "正在调用 AI Chat Completions。"})

    async for event in stream_proofread_with_ai(
        text,
        previous_response_id=previous_response_id,
        provider_api=provider_api,
        proofread_mode=proofread_mode,
    ):
        if event.event == "result":
            response_id = event.data.pop("response_id", None)
            raw_issues = [ProofreadIssue.model_validate(issue) for issue in event.data.get("issues", [])]
            located_issues = locate_issues(text, raw_issues)
            event.data["issues"] = [issue.model_dump() for issue in located_issues]
            if provider_api == "responses":
                update_last_response_id(session_id, response_id if isinstance(response_id, str) else None)
                logger.info(
                    "proofread stream service updated session response_id_present=%s session_id=%s",
                    isinstance(response_id, str),
                    _mask_session_id(session_id),
                )

        yield event


def build_mock_issues(text: str) -> list[ProofreadIssue]:
    sample = text[: min(len(text), 12)]

    return [
        ProofreadIssue(
            id="mock-issue-1",
            category="style",
            severity="medium",
            original=sample,
            replacement=f"{sample}（建议核对）" if sample else None,
            suggestion="请结合上下文检查该表述是否准确、简洁，并确认是否符合出版物体例。",
        )
    ]


def resolve_provider_api(provider_api: ProviderAPI | None) -> ProviderAPI:
    if provider_api:
        return provider_api

    settings = get_settings()
    if settings.ai_provider_api in {"responses", "chat"}:
        return settings.ai_provider_api

    raise AIClientError("AI_PROVIDER_API must be responses or chat")


def locate_issues(text: str, issues: list[ProofreadIssue]) -> list[ProofreadIssue]:
    search_from_by_original: dict[str, int] = {}
    located: list[ProofreadIssue] = []

    for issue in issues:
        original = issue.original.strip()
        start: int | None = None
        end: int | None = None

        if original:
            search_from = search_from_by_original.get(original, 0)
            found_at = text.find(original, search_from)

            if found_at == -1 and search_from > 0:
                found_at = text.find(original)

            if found_at != -1:
                start = found_at
                end = found_at + len(original)
                search_from_by_original[original] = end

        located.append(issue.model_copy(update={"start": start, "end": end}))
        logger.debug(
            "issue located issue_id=%s original_len=%s start=%s end=%s",
            issue.id,
            len(original),
            start,
            end,
        )

    logger.info(
        "issue location completed issue_count=%s located_issue_count=%s unlocated_issue_count=%s",
        len(located),
        sum(1 for issue in located if issue.start is not None and issue.end is not None),
        sum(1 for issue in located if issue.start is None or issue.end is None),
    )
    return located


def _mask_session_id(session_id: str | None) -> str:
    if not session_id:
        return "none"

    return f"...{session_id[-8:]}"
