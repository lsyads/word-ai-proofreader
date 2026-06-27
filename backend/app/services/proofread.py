from __future__ import annotations

import logging
import re
import unicodedata
from collections.abc import AsyncIterator
from typing import Literal

from app.schemas import BookInfo, ProofreadIssue, ProofreadLocator
from app.services.ai_client import (
    AIClientError,
    AIStreamEvent,
    DEFAULT_TEMPERATURE,
    V2PromptContext,
    proofread_with_ai,
    stream_proofread_with_ai,
)
from app.services.ai_profiles import resolve_ai_profile
from app.services.ai_profiles import resolve_provider_api as resolve_profile_provider_api
from app.settings import get_settings

ProviderAPI = Literal["responses", "chat"]
ProofreadMode = Literal["fast", "thinking"]
logger = logging.getLogger(__name__)
MIN_ORIGINAL_LOCATOR_LENGTH = 6
MAX_LOCATOR_OCCURRENCES = 3
CONTEXT_LOCATOR_WINDOWS = (16, 32, 64)
MECHANICAL_ISSUE_KEYWORDS = (
    "标点",
    "空格",
    "空白字符",
    "制表",
    "换行",
    "全角",
    "半角",
    "中英文符号",
    "中文符号",
    "英文符号",
    "重复符号",
)


async def proofread_text(
    text: str,
    book: BookInfo,
    session_id: str | None = None,
    ai_profile_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
    temperature: float = DEFAULT_TEMPERATURE,
) -> list[ProofreadIssue]:
    return await proofread_text_with_context(
        text,
        book,
        session_id=session_id,
        ai_profile_id=ai_profile_id,
        provider_api=provider_api,
        proofread_mode=proofread_mode,
        reasoning_enabled=reasoning_enabled,
        temperature=temperature,
        v2_context=None,
    )


async def proofread_text_with_context(
    text: str,
    book: BookInfo,
    session_id: str | None = None,
    ai_profile_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
    temperature: float = DEFAULT_TEMPERATURE,
    v2_context: V2PromptContext | None = None,
) -> list[ProofreadIssue]:
    settings = get_settings()
    profile = resolve_ai_profile(settings, ai_profile_id)
    provider_api = resolve_profile_provider_api(profile, provider_api)
    logger.info(
        "proofread service started text_len=%s ai_profile_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s temperature=%s has_api_key=%s session_id=%s",
        len(text),
        profile.id,
        provider_api,
        proofread_mode,
        reasoning_enabled,
        temperature,
        bool(profile.api_key),
        _mask_session_id(session_id),
    )

    if profile.api_key:
        logger.info(
            "proofread service calling AI provider ai_profile_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s temperature=%s",
            profile.id,
            provider_api,
            proofread_mode,
            reasoning_enabled,
            temperature,
        )
        ai_kwargs = {
            "provider_api": provider_api,
            "proofread_mode": proofread_mode,
        }
        if ai_profile_id is not None:
            ai_kwargs["ai_profile_id"] = profile.id
        if reasoning_enabled:
            ai_kwargs["reasoning_enabled"] = True
        if temperature != DEFAULT_TEMPERATURE:
            ai_kwargs["temperature"] = temperature

        if v2_context is not None:
            ai_kwargs["v2_context"] = v2_context
        result = await proofread_with_ai(text, book, **ai_kwargs)
        return locate_issues(text, result.issues)

    logger.info("proofread service using mock issues")
    return locate_issues(text, build_mock_issues(text, v2_context))


async def stream_proofread_text(
    text: str,
    book: BookInfo,
    session_id: str | None = None,
    ai_profile_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
    temperature: float = DEFAULT_TEMPERATURE,
) -> AsyncIterator[AIStreamEvent]:
    settings = get_settings()
    profile = resolve_ai_profile(settings, ai_profile_id)
    provider_api = resolve_profile_provider_api(profile, provider_api)
    logger.info(
        "proofread stream service started text_len=%s ai_profile_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s temperature=%s has_api_key=%s session_id=%s",
        len(text),
        profile.id,
        provider_api,
        proofread_mode,
        reasoning_enabled,
        temperature,
        bool(profile.api_key),
        _mask_session_id(session_id),
    )

    yield AIStreamEvent("status", {"stage": "received", "message": "已接收选区文本。"})

    if provider_api == "chat":
        raise AIClientError("Chat mode uses /api/proofread with standard Chat Completions, not SSE.")

    if not profile.api_key:
        logger.info("proofread stream service using mock issues")
        yield AIStreamEvent("status", {"stage": "calling_ai", "message": f"当前未配置 {profile.api_key_env}，正在返回本地 mock 审校结果。"})
        yield AIStreamEvent("status", {"stage": "normalizing", "message": "正在整理结构化审校结果。"})
        yield AIStreamEvent("result", {"issues": [issue.model_dump() for issue in locate_issues(text, build_mock_issues(text))]})
        yield AIStreamEvent("status", {"stage": "completed", "message": "审校完成。"})
        return

    yield AIStreamEvent("status", {"stage": "calling_ai", "message": "正在调用 AI Responses API。"})

    ai_kwargs = {
        "provider_api": provider_api,
        "proofread_mode": proofread_mode,
    }
    if ai_profile_id is not None:
        ai_kwargs["ai_profile_id"] = profile.id
    if reasoning_enabled:
        ai_kwargs["reasoning_enabled"] = True
    if temperature != DEFAULT_TEMPERATURE:
        ai_kwargs["temperature"] = temperature

    async for event in stream_proofread_with_ai(text, book, **ai_kwargs):
        if event.event == "result":
            event.data.pop("response_id", None)
            raw_issues = [ProofreadIssue.model_validate(issue) for issue in event.data.get("issues", [])]
            located_issues = locate_issues(text, raw_issues)
            event.data["issues"] = [issue.model_dump() for issue in located_issues]

        yield event


def build_mock_issues(text: str, v2_context: V2PromptContext | None = None) -> list[ProofreadIssue]:
    sample = text[: min(len(text), 12)]
    prefix = f"[{v2_context.pass_name}] " if v2_context else ""

    return [
        ProofreadIssue(
            id="mock-issue-1",
            category="style",
            severity="medium",
            original=sample,
            replacement=f"{sample}（建议核对）" if sample else None,
            suggestion=f"{prefix}请结合审校目标检查该表述是否准确、简洁，并确认是否符合出版物体例。",
        )
    ]


def resolve_provider_api(
    provider_api: ProviderAPI | None,
    ai_profile_id: str | None = None,
) -> ProviderAPI:
    settings = get_settings()
    profile = resolve_ai_profile(settings, ai_profile_id)
    return resolve_profile_provider_api(profile, provider_api)


def locate_issues(text: str, issues: list[ProofreadIssue]) -> list[ProofreadIssue]:
    search_from_by_original: dict[str, int] = {}
    located: list[ProofreadIssue] = []
    filtered_count = 0

    for issue in issues:
        if _is_mechanical_copyediting_issue(issue):
            filtered_count += 1
            logger.debug("issue filtered as mechanical copyediting issue issue_id=%s", issue.id)
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

        locator = build_locator(text, original, start, end)
        located.append(issue.model_copy(update={"start": start, "end": end, "locator": locator}))
        logger.debug(
            "issue located issue_id=%s original_len=%s start=%s end=%s locator_strategy=%s",
            issue.id,
            len(original),
            start,
            end,
            locator.strategy if locator else None,
        )

    logger.info(
        "issue location completed issue_count=%s filtered_mechanical_issue_count=%s located_issue_count=%s unlocated_issue_count=%s",
        len(issues),
        filtered_count,
        sum(1 for issue in located if issue.start is not None and issue.end is not None),
        sum(1 for issue in located if issue.start is None or issue.end is None),
    )
    return located


def build_locator(
    text: str,
    original: str,
    start: int | None,
    end: int | None,
) -> ProofreadLocator | None:
    if start is None or end is None or not original:
        return None

    if text[start:end] != original:
        return None

    if (
        len(original) >= MIN_ORIGINAL_LOCATOR_LENGTH
        and _count_occurrences(text, original) <= MAX_LOCATOR_OCCURRENCES
    ):
        return ProofreadLocator(
            key=original,
            key_start=start,
            key_end=end,
            original_start_in_key=0,
            original_end_in_key=len(original),
            strategy="original",
            key_occurrence_index=_count_occurrences_before_offset(text, original, start),
        )

    for window in CONTEXT_LOCATOR_WINDOWS:
        key_start = max(0, start - window)
        key_end = min(len(text), end + window)
        key = text[key_start:key_end]

        if not key or _count_occurrences(text, key) > MAX_LOCATOR_OCCURRENCES:
            continue

        return ProofreadLocator(
            key=key,
            key_start=key_start,
            key_end=key_end,
            original_start_in_key=start - key_start,
            original_end_in_key=end - key_start,
            strategy="context",
            key_occurrence_index=_count_occurrences_before_offset(text, key, key_start),
        )

    return None


def _count_occurrences(text: str, needle: str) -> int:
    if not needle:
        return 0

    count = 0
    search_from = 0

    while search_from < len(text):
        found_at = text.find(needle, search_from)
        if found_at == -1:
            break

        count += 1
        search_from = found_at + 1

    return count


def _count_occurrences_before_offset(text: str, needle: str, target_start: int) -> int:
    if not needle:
        return 0

    count = 0
    search_from = 0

    while search_from < target_start:
        found_at = text.find(needle, search_from)
        if found_at == -1 or found_at >= target_start:
            break

        count += 1
        search_from = found_at + 1

    return count


def _is_whitespace_only_change(issue: ProofreadIssue) -> bool:
    if issue.replacement is None:
        return False

    return _remove_all_whitespace(issue.original) == _remove_all_whitespace(issue.replacement)


def _is_mechanical_copyediting_issue(issue: ProofreadIssue) -> bool:
    if _is_punctuation_category(issue.category):
        return True
    if _is_whitespace_only_change(issue):
        return True
    if _is_mechanical_suggestion(issue):
        return True
    return _is_punctuation_or_spacing_only_change(issue)


def _is_punctuation_category(category: str) -> bool:
    normalized = category.strip().lower()
    return normalized == "punctuation" or "标点" in category


def _is_mechanical_suggestion(issue: ProofreadIssue) -> bool:
    text = f"{issue.category} {issue.suggestion}".lower()
    return any(keyword in text for keyword in MECHANICAL_ISSUE_KEYWORDS)


def _is_punctuation_or_spacing_only_change(issue: ProofreadIssue) -> bool:
    if issue.replacement is None:
        return False

    return _editorial_semantic_text(issue.original) == _editorial_semantic_text(issue.replacement)


def _remove_all_whitespace(value: str) -> str:
    return re.sub(r"\s+", "", value)


def _editorial_semantic_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value)
    return "".join(
        char
        for char in normalized
        if not char.isspace() and not unicodedata.category(char).startswith("P")
    )


def _mask_session_id(session_id: str | None) -> str:
    if not session_id:
        return "none"

    return f"...{session_id[-8:]}"
