from __future__ import annotations

import json
import logging
from collections.abc import AsyncIterator
from dataclasses import dataclass
from typing import Any, Literal

import httpx
from pydantic import ValidationError

from app.schemas import BookInfo, ProofreadIssue
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

总原则：
1. 只审校用户提供的 <text> 标签内文本。
2. 只输出能够定位到原文的明确问题。
3. 不要为了凑数量而输出低置信度问题。
4. 不要提出纯风格偏好、主观润色、扩写、标题美化建议。
5. 对于可改可不改的问题，优先不输出。
6. 缺少上下文或需要外部资料确认时，只能提示“需人工核查”，replacement 必须为 null。

输出规则：
1. 只返回紧凑 JSON，不要返回 Markdown、解释、代码块或多余文本。
2. 必须返回一个 JSON 对象，顶层只包含 issues 字段。
3. 如果没有明确问题，返回 {"issues": []}。

定位规则：
1. original 必须逐字摘录自用户提供的 <text> 中，不能改写、概括或补全。
2. original 必须是连续文本片段，不要跨越多个不连续位置。
3. 如果同一问题出现多次，应分别返回多个 issue，并使用各自对应的 original。

replacement 规则：
1. replacement 只能填写可直接替换 original 的正文文本。
2. replacement 不得包含解释、括号说明、批注语气、Markdown 或多余标记。
3. replacement 不得改变原文核心含义。
4. 如果是事实待核、知识点待核、数据待核、逻辑疑问、体例疑问、需要人工判断、无法直接替换的问题，replacement 必须为 null。
5. 如果建议涉及较大范围重写、增删整句、调整段落结构，replacement 必须为 null。

suggestion 规则：
1. suggestion 写给责任编辑看，说明问题原因和处理建议。
2. suggestion 要简短明确，不要超过 100 个汉字。
3. suggestion 不要重复 original 和 replacement 的完整内容。
4. suggestion 不要使用“建议考虑”“可以适当”等含糊表述，应明确指出问题。
5. 对需要外部资料、全书上下文或人工判断的问题，必须明确写“需人工核查”。

严重程度规则：
1. high：事实错误、知识点错误、数据错误、严重逻辑矛盾、可能影响出版准确性的硬伤。
2. medium：明显病句、搭配不当、语义不清、整段表达不通顺、体例明显不一致。
3. low：错别字、标点误用、轻微表述问题。

分类规则：
category 只能使用以下值之一：
- typo：错别字、漏字、多字
- grammar：语法、病句、搭配不当、语义不清、指代不明、整段不通顺
- punctuation：标点误用
- consistency：前后不一致、称谓/数字/时间/单位/数据不一致
- fact：事实疑问、知识点错误、概念混淆、数据错误、公式错误、明显事实冲突
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

审校范围：
1. 错别字、漏字、多字。
2. 明显病句、搭配不当、语义不清、指代不明。
3. 明显前后矛盾、称谓不一致、数字或时间前后不一致。
4. 标点误用。
5. 明显出版物体例硬伤。

取舍标准：
1. 只检查高置信度、明显可判断、通常不需要外部资料的问题。
2. 忽略轻微风格问题、可改可不改的问题、需要上下文判断的问题、需要专业资料核查的问题。
3. 优先少误报。
""".strip(),

    "thinking": """
当前模式：深度审校。

审校范围：
1. 基础语言问题：错别字、漏字、多字、明显病句、语义不清、搭配不当、指代不明、标点误用。
2. 整段表达问题：整段语句是否通顺，句间逻辑是否连贯，主谓宾关系是否清楚，表达是否符合正式出版物规范。
3. 知识点问题：概念、术语、定义、分类、原理、因果关系、适用条件、实验方法、专业表述是否准确、严谨。
4. 数据问题：数字、单位、比例、公式、范围、阈值、时间、数量级、统计口径、前后数据是否存在错误、矛盾或疑似缺少依据。
5. 事实与逻辑问题：明显事实冲突、前后矛盾、因果倒置、结论与依据不匹配、表述过度绝对。
6. 出版物体例问题：明显不符合出版规范的表达、格式、称谓、数字用法、单位写法等硬伤。

取舍标准：
1. 能根据文本本身或通用知识明确判断的问题，可以给出 replacement。
2. 需要查证外部资料、依赖全书上下文、存在多种可能解释的问题，replacement 必须为 null，并在 suggestion 中明确写“需人工核查”。
3. 不要为了凑数量输出低置信度问题。
4. 不要把正常表达改成个人偏好的表达。
""".strip(),
}


async def proofread_with_ai(
    text: str,
    book: BookInfo,
    settings: Settings | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
) -> AIProofreadResult:
    settings = settings or get_settings()
    provider_api = _resolve_provider_api(settings, provider_api)

    if provider_api == "chat":
        return await _proofread_with_chat(
            text,
            book,
            settings,
            proofread_mode=proofread_mode,
            reasoning_enabled=reasoning_enabled,
        )

    _ensure_responses_api(settings)

    payload = _build_responses_payload(
        text,
        book,
        settings,
        proofread_mode=proofread_mode,
    )
    logger.info(
        "AI responses request started model=%s proofread_mode=%s text_len=%s max_output_tokens=%s",
        settings.openai_model,
        proofread_mode,
        len(text),
        payload["max_output_tokens"],
    )

    _debug_log_json("AI responses request payload", payload)
    headers = _auth_headers(settings)

    async with httpx.AsyncClient(timeout=settings.ai_request_timeout_seconds) as client:
        response = await client.post(
            _responses_url(settings),
            headers=headers,
            json=payload,
        )

    logger.info("AI responses request completed status_code=%s", response.status_code)
    _raise_for_provider_error(response.status_code)

    try:
        data = response.json()
    except ValueError as exc:
        raise AIClientError("AI provider response body was not valid JSON") from exc

    _info_log_json("AI responses response payload", data)
    _debug_log_json("AI responses response payload", data)
    result = _parse_response_payload(data)
    logger.info(
        "AI responses payload parsed issue_count=%s response_id_present=%s",
        len(result.issues),
        bool(result.response_id),
    )
    return result


async def stream_proofread_with_ai(
    text: str,
    book: BookInfo,
    settings: Settings | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
) -> AsyncIterator[AIStreamEvent]:
    settings = settings or get_settings()
    provider_api = _resolve_provider_api(settings, provider_api)

    if provider_api == "chat":
        raise AIClientError("Chat mode uses the standard Chat Completions response, not SSE.")

    _ensure_responses_api(settings)

    payload = _build_responses_payload(
        text,
        book,
        settings,
        proofread_mode=proofread_mode,
        stream=True,
    )
    output_text = ""
    logger.info(
        "AI responses stream started model=%s proofread_mode=%s text_len=%s max_output_tokens=%s",
        settings.openai_model,
        proofread_mode,
        len(text),
        payload["max_output_tokens"],
    )

    _debug_log_json("AI responses stream request payload", payload)
    headers = _auth_headers(settings)

    async with httpx.AsyncClient(timeout=settings.ai_request_timeout_seconds) as client:
        async with client.stream(
            "POST",
            _responses_url(settings),
            headers=headers,
            json=payload,
        ) as response:
            logger.info("AI responses stream connected status_code=%s", response.status_code)
            _raise_for_provider_error(response.status_code)

            async for provider_event in _iter_sse_events(response.aiter_lines()):
                event_name = provider_event["event"]
                data = provider_event["data"]
                logger.debug("AI responses stream event event=%s", event_name)

                if event_name == "response.created":
                    yield AIStreamEvent("status", {"stage": "calling_ai", "message": "AI 已创建响应。"})
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
                    _info_log_json(
                        "AI responses stream completed payload",
                        {"response": provider_response, "output_text": output_text},
                    )
                    _debug_log_json(
                        "AI responses stream completed payload",
                        {"response": provider_response, "output_text": output_text},
                    )
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
                    _info_log_json("AI responses stream terminal event payload", data)
                    _debug_log_json("AI responses stream terminal event payload", data)
                    raise AIClientError(f"AI provider Responses API returned {event_name}")


def _ensure_responses_api(settings: Settings) -> None:
    _ensure_api_key(settings)


async def _proofread_with_chat(
    text: str,
    book: BookInfo,
    settings: Settings,
    proofread_mode: ProofreadMode,
    reasoning_enabled: bool,
) -> AIProofreadResult:
    _ensure_api_key(settings)
    payload = _build_chat_payload(
        text,
        book,
        settings,
        proofread_mode=proofread_mode,
        reasoning_enabled=reasoning_enabled,
    )
    logger.info(
        "AI chat request started model=%s proofread_mode=%s reasoning_enabled=%s text_len=%s max_tokens=%s",
        settings.openai_model,
        proofread_mode,
        reasoning_enabled,
        len(text),
        payload["max_tokens"],
    )

    _debug_log_json("AI chat request payload", payload)
    headers = _auth_headers(settings)

    async with httpx.AsyncClient(timeout=settings.ai_request_timeout_seconds) as client:
        response = await client.post(
            _chat_url(settings),
            headers=headers,
            json=payload,
        )

    logger.info("AI chat request completed status_code=%s", response.status_code)
    _raise_for_provider_error(response.status_code, provider_api="chat")

    try:
        data = response.json()
    except ValueError as exc:
        raise AIClientError("AI provider response body was not valid JSON") from exc

    _info_log_json("AI chat response payload", data)
    _debug_log_json("AI chat response payload", data)
    result = _parse_chat_payload(data)
    logger.info(
        "AI chat payload parsed issue_count=%s response_id_present=%s",
        len(result.issues),
        bool(result.response_id),
    )
    return result


def _build_responses_payload(
    text: str,
    book: BookInfo,
    settings: Settings,
    proofread_mode: ProofreadMode = "fast",
    stream: bool = False,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "model": settings.openai_model,
        "input": f"{_build_system_prompt(proofread_mode)}\n\n{_build_user_prompt(text, book)}",
        "temperature": 0.2,
        "max_output_tokens": _max_tokens_for_mode(settings, proofread_mode),
        "text": {"format": {"type": "json_object"}},
    }

    if stream:
        payload["stream"] = True

    return payload


def _build_chat_payload(
    text: str,
    book: BookInfo,
    settings: Settings,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
) -> dict[str, Any]:
    return {
        "model": settings.openai_model,
        "messages": [
            {"role": "system", "content": _build_system_prompt(proofread_mode)},
            {"role": "user", "content": _build_user_prompt(text, book)},
        ],
        "temperature": 0.2,
        "max_tokens": _max_tokens_for_mode(settings, proofread_mode),
        "reasoning": {"enabled": reasoning_enabled},
        # "response_format": {"type": "json_object"},
    }


def _build_system_prompt(proofread_mode: ProofreadMode) -> str:
    return f"{BASE_SYSTEM_PROMPT}\n{MODE_PROMPTS[proofread_mode]}"


def _build_user_prompt(text: str, book: BookInfo) -> str:
    book_context = json.dumps(
        book.model_dump(),
        ensure_ascii=False,
        separators=(",", ":"),
    )
    return f"""
请审校下面 <text> 标签内的 Word 选区文本。

书籍背景信息如下，仅用于理解文本语境，不属于待审正文：
<book>
{book_context}
</book>

注意：
1. <text> 和 </text> 只是边界标记，不属于正文。
2. 只审校标签内文本。
3. original 必须来自标签内文本的原文片段。
4. 不要对 <book> 中的书名或介绍本身输出问题。
5. 不要输出标签外的内容。

<text>
{text}
</text>
""".strip()


def _debug_log_json(message: str, payload: Any) -> None:
    if logger.isEnabledFor(logging.DEBUG):
        logger.debug("%s: %s", message, json.dumps(payload, ensure_ascii=False, default=str))


def _info_log_json(message: str, payload: Any) -> None:
    if logger.isEnabledFor(logging.INFO):
        logger.info(
            "%s: %s",
            message,
            json.dumps(_redact_sensitive_payload(payload), ensure_ascii=False, default=str),
        )


def _redact_sensitive_payload(value: Any) -> Any:
    if isinstance(value, dict):
        return {
            key: "[REDACTED]" if _is_sensitive_key(key) else _redact_sensitive_payload(item)
            for key, item in value.items()
        }

    if isinstance(value, list):
        return [_redact_sensitive_payload(item) for item in value]

    return value


def _is_sensitive_key(key: str) -> bool:
    normalized = key.lower().replace("-", "_")
    return normalized in {
        "authorization",
        "api_key",
        "access_token",
        "refresh_token",
    } or normalized.endswith("_key")


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
        logger.warning("AI provider responses API unsupported status_code=%s", status_code)
        raise AIClientError(f"AI provider does not support Responses API (HTTP {status_code})")

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
    parsed = _parse_provider_json_content(content)

    if not isinstance(parsed, dict) or not isinstance(parsed.get("issues"), list):
        raise AIClientError("AI provider response did not match the expected schema")

    try:
        issues = [ProofreadIssue.model_validate(issue) for issue in parsed["issues"]]
    except ValidationError as exc:
        raise AIClientError("AI provider response issues did not match the expected schema") from exc

    logger.debug("AI issues parsed issue_count=%s content_len=%s", len(issues), len(content))
    return issues


def _parse_provider_json_content(content: str) -> Any:
    candidates = _json_content_candidates(content)
    first_error: json.JSONDecodeError | None = None

    for index, candidate in enumerate(candidates):
        try:
            parsed = json.loads(candidate)
            if index > 0:
                logger.warning("AI provider response JSON required cleanup before parsing")
            return parsed
        except json.JSONDecodeError as exc:
            first_error = first_error or exc

        try:
            parsed = json.loads(candidate, strict=False)
            logger.warning(
                "AI provider response JSON required lenient parsing error_message=%s",
                str(first_error or ""),
            )
            return parsed
        except json.JSONDecodeError:
            continue

    raise AIClientError("AI provider response was not valid JSON") from first_error


def _json_content_candidates(content: str) -> list[str]:
    stripped = content.strip()
    candidates = [stripped]
    without_fence = _strip_markdown_json_fence(stripped)

    if without_fence != stripped:
        candidates.append(without_fence)

    extracted = _extract_first_json_object(without_fence)
    if extracted and extracted not in candidates:
        candidates.append(extracted)

    cleaned_candidates = list(candidates)
    for candidate in candidates:
        cleaned = _remove_trailing_commas(candidate)
        if cleaned != candidate and cleaned not in cleaned_candidates:
            cleaned_candidates.append(cleaned)

    return cleaned_candidates


def _strip_markdown_json_fence(content: str) -> str:
    lines = content.splitlines()

    if len(lines) >= 2 and lines[0].strip().startswith("```") and lines[-1].strip() == "```":
        return "\n".join(lines[1:-1]).strip()

    return content


def _extract_first_json_object(content: str) -> str | None:
    start = content.find("{")

    if start == -1:
        return None

    in_string = False
    escaped = False
    depth = 0

    for index in range(start, len(content)):
        char = content[index]

        if escaped:
            escaped = False
            continue

        if char == "\\":
            escaped = True
            continue

        if char == '"':
            in_string = not in_string
            continue

        if in_string:
            continue

        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return content[start : index + 1].strip()

    return None


def _remove_trailing_commas(content: str) -> str:
    cleaned: list[str] = []
    in_string = False
    escaped = False
    index = 0

    while index < len(content):
        char = content[index]

        if escaped:
            cleaned.append(char)
            escaped = False
            index += 1
            continue

        if char == "\\":
            cleaned.append(char)
            escaped = True
            index += 1
            continue

        if char == '"':
            cleaned.append(char)
            in_string = not in_string
            index += 1
            continue

        if char == "," and not in_string:
            lookahead = index + 1
            while lookahead < len(content) and content[lookahead].isspace():
                lookahead += 1
            if lookahead < len(content) and content[lookahead] in "]}":
                index += 1
                continue

        cleaned.append(char)
        index += 1

    return "".join(cleaned)


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
