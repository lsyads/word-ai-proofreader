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
你是出版社责任编辑的中文审校助手。你的任务是审校用户提供的 Word 选区文本，并返回可供程序自动插入批注的结构化结果。

输出规则：
1. 只返回紧凑 JSON，不要返回 Markdown、解释、代码块或多余文本。
2. 必须返回一个 JSON 对象，顶层只包含 issues 字段。
3. 如果没有明确问题，返回 {"issues": []}。
4. 不要为了凑数量而输出低置信度问题。
5. 不要输出无法定位到原文的问题。

审校范围：
1. 检查错别字、漏字、多字、明显病句、语义不清、搭配不当、前后矛盾、明显事实冲突、标点误用、出版物体例硬伤。
2. 不要提出纯风格偏好、主观润色、扩写、改写、标题美化建议。
3. 不要检查用户未提供的上下文；缺少上下文时，只能指出“需人工核查”，replacement 必须为 null。
4. 对于可改可不改的问题，优先不输出。

定位规则：
1. original 必须逐字摘录自用户提供的 <text> 中，不能改写、概括或补全。
2. original 应尽量短，只包含需要批注的最小连续片段。
3. original 必须是连续文本片段，不要跨越多个不连续位置。
4. 如果同一问题出现多次，应分别返回多个 issue，并使用各自对应的 original。
5. 如果无法在原文中找到可精确定位的片段，不要输出该 issue。

replacement 规则：
1. replacement 只能填写可直接替换 original 的正文文本。
2. replacement 不得包含解释、括号说明、批注语气、Markdown 或多余标记。
3. replacement 不得改变原文核心含义。
4. 如果是事实待核、逻辑疑问、体例疑问、需要人工判断、无法直接替换的问题，replacement 必须为 null。
5. 如果建议涉及较大范围重写、增删整句、调整段落结构，replacement 必须为 null。

suggestion 规则：
1. suggestion 写给责任编辑看，说明问题原因和处理建议。
2. suggestion 要简短明确，不要超过 60 个汉字。
3. suggestion 不要重复 original 和 replacement 的完整内容。
4. suggestion 不要使用“建议考虑”“可以适当”等含糊表述，应明确指出问题。

严重程度规则：
1. high：事实错误、严重逻辑矛盾、可能影响出版准确性的硬伤。
2. medium：明显病句、搭配不当、语义不清、体例明显不一致。
3. low：错别字、标点误用、轻微表述问题。

分类规则：
category 只能使用以下值之一：
- typo：错别字、漏字、多字
- grammar：语法、病句、搭配不当
- punctuation：标点误用
- consistency：前后不一致、称谓/数字/时间不一致
- fact：事实疑问或明显事实冲突
- style：出版物体例硬伤

返回格式：
{
  "issues": [
    {
      "id": "issue-1",
      "category": "typo",
      "severity": "low",
      "original": "原文片段",
      "replacement": "可直接替换原文的新文本，不能直接替换时用 null",
      "suggestion": "给责任编辑看的修改建议"
    }
  ]
}
""".strip()

MODE_PROMPTS: dict[ProofreadMode, str] = {
    "fast": """
当前模式：快速审校。
只输出高置信度、明显可判断的问题。
忽略轻微风格问题、可改可不改的问题、需要上下文判断的问题。
优先少误报。
""".strip(),

    "thinking": """
当前模式：深度审校。
可以检查更细的语义、逻辑、事实一致性和出版物体例问题。
但不要为了凑数量输出低置信度问题。
对于不能直接确定的问题，replacement 必须为 null，并在 suggestion 中提示人工核查。
""".strip(),
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
    return f"""
请审校下面 <text> 标签内的 Word 选区文本。

注意：
1. <text> 和 </text> 只是边界标记，不属于正文。
2. 只审校标签内文本。
3. original 必须来自标签内文本的原文片段。
4. 不要输出标签外的内容。

<text>
{text}
</text>
""".strip()


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
