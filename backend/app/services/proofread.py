from __future__ import annotations

import logging
import re
from collections.abc import AsyncIterator
from typing import Literal

from app.schemas import BookInfo, ProofreadIssue
from app.services.ai_client import AIClientError, AIStreamEvent, proofread_with_ai, stream_proofread_with_ai
from app.settings import get_settings

ProviderAPI = Literal["responses", "chat"]
ProofreadMode = Literal["fast", "thinking"]
logger = logging.getLogger(__name__)


async def proofread_text(
    text: str,
    book: BookInfo,
    session_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
) -> list[ProofreadIssue]:
    settings = get_settings()
    provider_api = resolve_provider_api(provider_api)
    logger.info(
        "proofread service started text_len=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s has_api_key=%s session_id=%s",
        len(text),
        provider_api,
        proofread_mode,
        reasoning_enabled,
        bool(settings.ai_api_key),
        _mask_session_id(session_id),
    )

    if settings.ai_api_key:
        logger.info(
            "proofread service calling AI provider provider_api=%s proofread_mode=%s reasoning_enabled=%s",
            provider_api,
            proofread_mode,
            reasoning_enabled,
        )
        ai_kwargs = {"provider_api": provider_api, "proofread_mode": proofread_mode}
        if reasoning_enabled:
            ai_kwargs["reasoning_enabled"] = True

        result = await proofread_with_ai(text, book, **ai_kwargs)
        return locate_issues(text, result.issues)

    logger.info("proofread service using mock issues")
    return locate_issues(text, build_mock_issues(text))


async def stream_proofread_text(
    text: str,
    book: BookInfo,
    session_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
) -> AsyncIterator[AIStreamEvent]:
    settings = get_settings()
    provider_api = resolve_provider_api(provider_api)
    logger.info(
        "proofread stream service started text_len=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s has_api_key=%s session_id=%s",
        len(text),
        provider_api,
        proofread_mode,
        reasoning_enabled,
        bool(settings.ai_api_key),
        _mask_session_id(session_id),
    )

    yield AIStreamEvent("status", {"stage": "received", "message": "已接收选区文本。"})

    if provider_api == "chat":
        raise AIClientError("Chat mode uses /api/proofread with standard Chat Completions, not SSE.")

    if not settings.ai_api_key:
        logger.info("proofread stream service using mock issues")
        yield AIStreamEvent("status", {"stage": "calling_ai", "message": "当前未配置 AI_API_KEY，正在返回本地 mock 审校结果。"})
        yield AIStreamEvent("status", {"stage": "normalizing", "message": "正在整理结构化审校结果。"})
        yield AIStreamEvent("result", {"issues": [issue.model_dump() for issue in locate_issues(text, build_mock_issues(text))]})
        yield AIStreamEvent("status", {"stage": "completed", "message": "审校完成。"})
        return

    yield AIStreamEvent("status", {"stage": "calling_ai", "message": "正在调用 AI Responses API。"})

    ai_kwargs = {"provider_api": provider_api, "proofread_mode": proofread_mode}
    if reasoning_enabled:
        ai_kwargs["reasoning_enabled"] = True

    async for event in stream_proofread_with_ai(text, book, **ai_kwargs):
        if event.event == "result":
            event.data.pop("response_id", None)
            raw_issues = [ProofreadIssue.model_validate(issue) for issue in event.data.get("issues", [])]
            located_issues = locate_issues(text, raw_issues)
            event.data["issues"] = [issue.model_dump() for issue in located_issues]

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
    filtered_count = 0

    for issue in issues:
        if _is_whitespace_only_change(issue):
            filtered_count += 1
            logger.debug("issue filtered as whitespace-only change issue_id=%s", issue.id)
            continue

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
        "issue location completed issue_count=%s filtered_whitespace_issue_count=%s located_issue_count=%s unlocated_issue_count=%s",
        len(issues),
        filtered_count,
        sum(1 for issue in located if issue.start is not None and issue.end is not None),
        sum(1 for issue in located if issue.start is None or issue.end is None),
    )
    return located


def _is_whitespace_only_change(issue: ProofreadIssue) -> bool:
    if issue.replacement is None:
        return False

    return _remove_all_whitespace(issue.original) == _remove_all_whitespace(issue.replacement)


def _remove_all_whitespace(value: str) -> str:
    return re.sub(r"\s+", "", value)


def _mask_session_id(session_id: str | None) -> str:
    if not session_id:
        return "none"

    return f"...{session_id[-8:]}"
