from __future__ import annotations

import json
import logging
from collections.abc import AsyncIterator
from dataclasses import dataclass
from typing import Any, Literal

import httpx
from pydantic import ValidationError

from app.schemas import ProofreadIssue
from app.settings import Settings, get_settings

ProviderAPI = Literal["responses", "chat"]
ProofreadMode = Literal["fast", "thinking"]
logger = logging.getLogger(__name__)


class AIClientError(RuntimeError):
    """Raised when the AI provider cannot return normalized proofread issues."""


@dataclass
class AIProofreadResult:
    issues: list[ProofreadIssue]
    response_id: str | None


@dataclass
class AIStreamEvent:
    event: str
    data: dict[str, Any]


BASE_SYSTEM_PROMPT = """
你是出版社责任编辑的中文审校助手。只返回紧凑 JSON，不要返回 Markdown。
返回格式：
{
  "issues": [
    {
      "id": "issue-1",
      "category": "typo",
      "severity": "low|medium|high",
      "original": "原文片段",
      "replacement": "可直接替换原文的新文本，不能直接替换时用 null",
      "suggestion": "给责任编辑看的修改建议"
    }
  ]
}
如果整体没有发现问题，返回 {"issues": []}。
suggestion是空格类问题，不写入issues。
后端会按 original 定位，不能修改原文。
replacement 只写可直接进入正文的替换文本；事实待核、需人工判断、体例疑问、无问题等情况必须返回 null。
""".strip()

MODE_PROMPTS: dict[ProofreadMode, str] = {
    "fast": "快速审校：只指出明显错别字、病句、事实矛盾或出版物体例硬伤，忽略空格问题，优先响应速度。",
    "thinking": "深度审校：细致检查文字、语法、风格、事实一致性和出版物体例，忽略空格问题，优先审校质量。",
}


async def proofread_with_ai(
    text: str,
    previous_response_id: str | None = None,
    settings: Settings | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
) -> AIProofreadResult:
    settings = settings or get_settings()
    provider_api = _resolve_provider_api(settings, provider_api)

    if provider_api == "chat":
        return await _proofread_with_chat(text, settings, proofread_mode=proofread_mode)

    _ensure_responses_api(settings)

    payload = _build_responses_payload(
        text,
        settings,
        previous_response_id=previous_response_id,
        proofread_mode=proofread_mode,
    )
    logger.info(
        "AI responses request started model=%s proofread_mode=%s text_len=%s max_output_tokens=%s has_previous_response=%s",
        settings.openai_model,
        proofread_mode,
        len(text),
        payload["max_output_tokens"],
        bool(previous_response_id),
    )

    async with httpx.AsyncClient(timeout=settings.ai_request_timeout_seconds) as client:
        response = await client.post(
            _responses_url(settings),
            headers=_auth_headers(settings),
            json=payload,
        )

    logger.info("AI responses request completed status_code=%s", response.status_code)
    _raise_for_provider_error(response.status_code)

    try:
        data = response.json()
    except ValueError as exc:
        raise AIClientError("AI provider response body was not valid JSON") from exc

    result = _parse_response_payload(data)
    logger.info(
        "AI responses payload parsed issue_count=%s response_id_present=%s",
        len(result.issues),
        bool(result.response_id),
    )
    return result


async def stream_proofread_with_ai(
    text: str,
    previous_response_id: str | None = None,
    settings: Settings | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
) -> AsyncIterator[AIStreamEvent]:
    settings = settings or get_settings()
    provider_api = _resolve_provider_api(settings, provider_api)

    if provider_api == "chat":
        raise AIClientError("Chat mode uses the standard Chat Completions response, not SSE.")

    _ensure_responses_api(settings)

    payload = _build_responses_payload(
        text,
        settings,
        previous_response_id=previous_response_id,
        proofread_mode=proofread_mode,
        stream=True,
    )
    output_text = ""
    logger.info(
        "AI responses stream started model=%s proofread_mode=%s text_len=%s max_output_tokens=%s has_previous_response=%s",
        settings.openai_model,
        proofread_mode,
        len(text),
        payload["max_output_tokens"],
        bool(previous_response_id),
    )

    async with httpx.AsyncClient(timeout=settings.ai_request_timeout_seconds) as client:
        async with client.stream(
            "POST",
            _responses_url(settings),
            headers=_auth_headers(settings),
            json=payload,
        ) as response:
            logger.info("AI responses stream connected status_code=%s", response.status_code)
            _raise_for_provider_error(response.status_code)

            async for provider_event in _iter_sse_events(response.aiter_lines()):
                event_name = provider_event["event"]
                data = provider_event["data"]
                logger.debug("AI responses stream event event=%s", event_name)

                if event_name == "response.created":
                    yield AIStreamEvent("status", {"stage": "calling_ai", "message": "AI 原生 session 已创建响应。"})
                    continue

                if event_name == "response.in_progress":
                    yield AIStreamEvent("status", {"stage": "calling_ai", "message": "AI 正在生成审校结果。"})
                    continue

                if event_name == "response.output_text.delta":
                    yield AIStreamEvent("status", {"stage": "calling_ai", "message": "AI 正在流式生成审校结果。"})
                    continue

                if event_name == "response.output_text.done":
                    output_text = _coerce_string(data.get("text"))
                    yield AIStreamEvent("status", {"stage": "normalizing", "message": "已收到 AI 输出，正在解析结构化结果。"})
                    continue

                if event_name == "response.completed":
                    provider_response = _coerce_dict(data.get("response"))
                    result = _parse_response_payload(provider_response, fallback_content=output_text)
                    logger.info(
                        "AI responses stream completed issue_count=%s response_id_present=%s",
                        len(result.issues),
                        bool(result.response_id),
                    )
                    yield AIStreamEvent(
                        "result",
                        {
                            "issues": [issue.model_dump() for issue in result.issues],
                            "response_id": result.response_id,
                        },
                    )
                    yield AIStreamEvent("status", {"stage": "completed", "message": "审校完成。"})
                    continue

                if event_name in {"response.failed", "response.incomplete"}:
                    raise AIClientError(f"AI provider Responses API returned {event_name}")


def _ensure_responses_api(settings: Settings) -> None:
    _ensure_api_key(settings)


async def _proofread_with_chat(
    text: str,
    settings: Settings,
    proofread_mode: ProofreadMode,
) -> AIProofreadResult:
    _ensure_api_key(settings)
    payload = _build_chat_payload(text, settings, proofread_mode=proofread_mode)
    logger.info(
        "AI chat request started model=%s proofread_mode=%s text_len=%s max_tokens=%s",
        settings.openai_model,
        proofread_mode,
        len(text),
        payload["max_tokens"],
    )

    async with httpx.AsyncClient(timeout=settings.ai_request_timeout_seconds) as client:
        response = await client.post(
            _chat_url(settings),
            headers=_auth_headers(settings),
            json=payload,
        )

    logger.info("AI chat request completed status_code=%s", response.status_code)
    _raise_for_provider_error(response.status_code, provider_api="chat")

    try:
        data = response.json()
    except ValueError as exc:
        raise AIClientError("AI provider response body was not valid JSON") from exc

    result = _parse_chat_payload(data)
    logger.info(
        "AI chat payload parsed issue_count=%s response_id_present=%s",
        len(result.issues),
        bool(result.response_id),
    )
    return result


def _build_responses_payload(
    text: str,
    settings: Settings,
    previous_response_id: str | None = None,
    proofread_mode: ProofreadMode = "fast",
    stream: bool = False,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "model": settings.openai_model,
        "input": f"{_build_system_prompt(proofread_mode)}\n\n{_build_user_prompt(text)}",
        "temperature": 0.2,
        "max_output_tokens": _max_tokens_for_mode(settings, proofread_mode),
        "text": {"format": {"type": "json_object"}},
    }

    if previous_response_id:
        payload["previous_response_id"] = previous_response_id

    if stream:
        payload["stream"] = True

    return payload


def _build_chat_payload(
    text: str,
    settings: Settings,
    proofread_mode: ProofreadMode = "fast",
) -> dict[str, Any]:
    return {
        "model": settings.openai_model,
        "messages": [
            {"role": "system", "content": _build_system_prompt(proofread_mode)},
            {"role": "user", "content": _build_user_prompt(text)},
        ],
        "temperature": 0.2,
        "max_tokens": _max_tokens_for_mode(settings, proofread_mode),
        "response_format": {"type": "json_object"},
    }


def _build_system_prompt(proofread_mode: ProofreadMode) -> str:
    return f"{BASE_SYSTEM_PROMPT}\n{MODE_PROMPTS[proofread_mode]}"


def _build_user_prompt(text: str) -> str:
    return f"请审校 <text> 内的 Word 选区文本：\n<text>\n{text}\n</text>"


def _max_tokens_for_mode(settings: Settings, proofread_mode: ProofreadMode) -> int:
    if proofread_mode == "thinking":
        return settings.ai_thinking_max_tokens

    return settings.ai_fast_max_tokens


def _responses_url(settings: Settings) -> str:
    return f"{settings.openai_api_base_url.rstrip('/')}/responses"


def _chat_url(settings: Settings) -> str:
    return f"{settings.openai_api_base_url.rstrip('/')}/chat/completions"


def _auth_headers(settings: Settings) -> dict[str, str]:
    return {"Authorization": f"Bearer {settings.ai_api_key}"}


def _raise_for_provider_error(status_code: int, provider_api: ProviderAPI = "responses") -> None:
    if status_code < 400:
        return

    if provider_api == "responses" and status_code in {400, 404, 405}:
        logger.warning("AI provider native responses API unsupported status_code=%s", status_code)
        raise AIClientError(f"AI provider does not support native Responses session API (HTTP {status_code})")

    logger.warning("AI provider returned error status_code=%s provider_api=%s", status_code, provider_api)
    raise AIClientError(f"AI provider returned HTTP {status_code}")


def _resolve_provider_api(settings: Settings, provider_api: ProviderAPI | None) -> ProviderAPI:
    resolved = provider_api or settings.ai_provider_api

    if resolved not in {"responses", "chat"}:
        raise AIClientError("AI_PROVIDER_API must be responses or chat")

    return resolved


def _ensure_api_key(settings: Settings) -> None:
    if not settings.ai_api_key:
        raise AIClientError("AI_API_KEY is not configured")


def _parse_response_payload(data: dict[str, Any], fallback_content: str = "") -> AIProofreadResult:
    response_id = _coerce_optional_string(data.get("id"))
    content = fallback_content or _extract_response_output_text(data)
    return AIProofreadResult(issues=_parse_issues(content), response_id=response_id)


def _parse_chat_payload(data: dict[str, Any]) -> AIProofreadResult:
    response_id = _coerce_optional_string(data.get("id"))

    try:
        content = data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise AIClientError("AI provider chat response did not include message content") from exc

    if not isinstance(content, str) or not content.strip():
        raise AIClientError("AI provider chat response did not include message content")

    return AIProofreadResult(issues=_parse_issues(content), response_id=response_id)


def _extract_response_output_text(data: dict[str, Any]) -> str:
    try:
        for output_item in data["output"]:
            for content_item in output_item.get("content", []):
                text = content_item.get("text")
                if isinstance(text, str) and text.strip():
                    return text
    except (KeyError, TypeError) as exc:
        raise AIClientError("AI provider response did not include output text") from exc

    raise AIClientError("AI provider response did not include output text")


async def _iter_sse_events(lines: AsyncIterator[str]) -> AsyncIterator[dict[str, Any]]:
    event_name = "message"
    data_lines: list[str] = []

    async for line in lines:
        if line == "":
            if data_lines:
                yield _parse_sse_event(event_name, data_lines)
            event_name = "message"
            data_lines = []
            continue

        if line.startswith(":"):
            continue

        if line.startswith("event: "):
            event_name = line.removeprefix("event: ")
            continue

        if line.startswith("data: "):
            data_lines.append(line.removeprefix("data: "))

    if data_lines:
        yield _parse_sse_event(event_name, data_lines)


def _parse_sse_event(event_name: str, data_lines: list[str]) -> dict[str, Any]:
    try:
        data = json.loads("\n".join(data_lines))
    except json.JSONDecodeError as exc:
        raise AIClientError("AI provider stream event was not valid JSON") from exc

    return {"event": event_name, "data": data}


def _parse_issues(content: str) -> list[ProofreadIssue]:
    try:
        parsed = json.loads(content)
    except json.JSONDecodeError as exc:
        raise AIClientError("AI provider response was not valid JSON") from exc

    if not isinstance(parsed, dict) or not isinstance(parsed.get("issues"), list):
        raise AIClientError("AI provider response did not match the expected schema")

    try:
        issues = [ProofreadIssue.model_validate(issue) for issue in parsed["issues"]]
    except ValidationError as exc:
        raise AIClientError("AI provider response issues did not match the expected schema") from exc

    logger.debug("AI issues parsed issue_count=%s content_len=%s", len(issues), len(content))
    return issues


def _coerce_dict(value: Any) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise AIClientError("AI provider response stream did not include response payload")

    return value


def _coerce_string(value: Any) -> str:
    if not isinstance(value, str):
        raise AIClientError("AI provider response stream did not include output text")

    return value


def _coerce_optional_string(value: Any) -> str | None:
    return value if isinstance(value, str) else None
