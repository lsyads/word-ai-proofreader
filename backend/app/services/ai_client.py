from __future__ import annotations

import json
import logging
import math
import re
from collections.abc import AsyncIterator
from dataclasses import dataclass
from typing import Any, Literal
from urllib.parse import urlparse

import httpx
import tiktoken
from pydantic import ValidationError

from app.schemas import BookInfo, ProofreadIssue
from app.services.ai_profiles import (
    AIProfile,
    AIProfileError,
    resolve_ai_profile,
    resolve_provider_api,
)
from app.settings import Settings, get_settings

ProviderAPI = Literal["responses", "chat"]
ProofreadMode = Literal["fast", "thinking"]
ChatDialect = Literal["default", "deepseek", "xiaomimimo"]
DEFAULT_TEMPERATURE = 0.2
logger = logging.getLogger(__name__)
_TOKENIZER_UNAVAILABLE_MODELS: set[str] = set()


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


@dataclass(frozen=True)
class AIRequestTimeoutEstimate:
    timeout_seconds: float
    estimated_input_tokens: int
    estimated_output_tokens: int
    estimated_total_tokens: int
    token_units: int
    proofread_mode: ProofreadMode
    reasoning_enabled: bool

    def model_dump(self) -> dict[str, Any]:
        return {
            "timeout_seconds": self.timeout_seconds,
            "estimated_input_tokens": self.estimated_input_tokens,
            "estimated_output_tokens": self.estimated_output_tokens,
            "estimated_total_tokens": self.estimated_total_tokens,
            "token_units": self.token_units,
            "proofread_mode": self.proofread_mode,
            "reasoning_enabled": self.reasoning_enabled,
        }


BASE_SYSTEM_PROMPT = """
你是出版社责任编辑的中文审校助手。审校 <text> 中的待审文本片段，并返回紧凑 JSON。

输出 JSON：
{"issues":[{"id":"issue-1","category":"typo","severity":"low","original":"原文片段","replacement":"可直接替换文本或 null","suggestion":"给责任编辑看的建议"}]}

硬性规则：
1. 只判断 <text> 内文本；<book> 和标签本身只作背景；original 必须是 <text> 内连续原文，不能改写、概括、补全或跨不连续位置。
2. 可结合模型已有知识、通用常识、专业知识和 <book> 背景判断问题；可输出语言、事实、知识、数据、公式、逻辑、前后一致性、术语和出版风险问题。
3. replacement 在有确定替换文本时填写；没有唯一替换、需实时查证、需编辑取舍或需大段改写时填 null，并在 suggestion 说明。
4. 不输出纯风格偏好、主观润色、扩写、标题美化、化学表达式小标，以及标点、空格、制表符、换行、全半角和中英文符号替换等机械校对项。
5. 没有问题返回 {"issues":[]}；只返回 JSON；顶层只包含 issues；不要 Markdown、解释、代码块或多余文本。
6. suggestion 写给责任编辑看，说明问题原因和处理建议，不超过 100 个汉字；不复述完整正文、密钥或认证头。

严重程度：
- high：事实、知识点、数据、公式错误，严重逻辑矛盾，影响出版准确性的硬伤。
- medium：明显病句、搭配不当、语义不清、指代不明、段落逻辑不顺、体例明显不一致。
- low：错别字、漏字、多字、轻微但明确的问题。

category 只能使用：
- typo：错别字、漏字、多字
- grammar：语法、病句、搭配不当、语义不清、指代不明、整段不通顺
- punctuation：兼容字段，模型不要主动使用
- consistency：前后不一致、称谓/数字/时间/单位/数据不一致
- fact：事实疑问、知识点错误、概念混淆、数据错误、公式错误、明显事实冲突
- style：出版物体例硬伤
- other：其他

""".strip()

MODE_PROMPTS: dict[ProofreadMode, str] = {
    "fast": """
当前模式：快速审校。

审校范围：
1. 局部文字问题：错别字、漏字、多字。
2. 局部表达问题：病句、搭配不当、语义不清、指代不明。
3. 局部一致性问题：同一片段内称谓、数字、时间、单位、术语前后不一致。
4. 常识性事实、知识、数据或逻辑问题。
5. 直接影响理解或出版准确性的风险。
""".strip(),

    "thinking": """
当前模式：深度审校。

审校范围：
1. 文字与语句：错别字、漏字、多字、病句、搭配不当、语义不清、指代不明。
2. 段落与表达：段落是否通顺，句间关系是否连贯，主谓宾关系是否清楚，表述是否符合正式出版物规范。
3. 事实与知识：事实冲突、概念混淆、定义错误、分类错误、原理错误、适用条件错误、专业表述不当。
4. 数据与公式：数字、单位、比例、公式、范围、阈值、时间、数量级、统计口径错误或前后矛盾。
5. 逻辑与结论：前后矛盾、因果倒置、结论与依据不匹配、表述过度绝对、推理链条断裂。
6. 出版风险：影响读者理解、知识准确性、体例一致性或出版判断的问题。
7. 其他审校问题：不限于上述类型；应充分结合模型已有知识、通用常识和专业知识，发现各类影响出版质量、知识准确性、读者理解、论述可信度或出版判断的问题。
""".strip(),
}


async def proofread_with_ai(
    text: str,
    book: BookInfo,
    settings: Settings | None = None,
    ai_profile_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
    temperature: float = DEFAULT_TEMPERATURE,
) -> AIProofreadResult:
    settings = settings or get_settings()
    try:
        profile = resolve_ai_profile(settings, ai_profile_id)
        provider_api = resolve_provider_api(profile, provider_api)
    except AIProfileError as exc:
        raise AIClientError(str(exc)) from exc

    if provider_api == "chat":
        return await _proofread_with_chat(
            text,
            book,
            settings,
            profile,
            proofread_mode=proofread_mode,
            reasoning_enabled=reasoning_enabled,
            temperature=temperature,
        )

    _ensure_responses_api(profile)

    payload = _build_responses_payload(
        text,
        book,
        settings,
        profile,
        proofread_mode=proofread_mode,
        temperature=temperature,
    )
    logger.info(
        "AI responses request started profile_id=%s model=%s proofread_mode=%s temperature=%s text_len=%s max_output_tokens=%s",
        profile.id,
        profile.model,
        proofread_mode,
        payload["temperature"],
        len(text),
        payload["max_output_tokens"],
    )

    _debug_log_json("AI responses request payload", payload)
    headers = _auth_headers(profile)
    timeout_seconds, timeout_meta = _responses_timeout(settings, profile, payload, proofread_mode, reasoning_enabled)
    logger.info(
        "AI responses request timeout estimated profile_id=%s model=%s estimated_input_tokens=%s estimated_output_tokens=%s estimated_total_tokens=%s token_units=%s timeout_seconds=%s proofread_mode=%s reasoning_enabled=%s",
        profile.id,
        profile.model,
        timeout_meta["estimated_input_tokens"],
        timeout_meta["estimated_output_tokens"],
        timeout_meta["estimated_total_tokens"],
        timeout_meta["token_units"],
        timeout_seconds,
        proofread_mode,
        reasoning_enabled,
    )

    async with httpx.AsyncClient(timeout=timeout_seconds) as client:
        response = await client.post(
            _responses_url(profile),
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
    ai_profile_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
    temperature: float = DEFAULT_TEMPERATURE,
) -> AsyncIterator[AIStreamEvent]:
    settings = settings or get_settings()
    try:
        profile = resolve_ai_profile(settings, ai_profile_id)
        provider_api = resolve_provider_api(profile, provider_api)
    except AIProfileError as exc:
        raise AIClientError(str(exc)) from exc

    if provider_api == "chat":
        raise AIClientError("Chat mode uses the standard Chat Completions response, not SSE.")

    _ensure_responses_api(profile)

    payload = _build_responses_payload(
        text,
        book,
        settings,
        profile,
        proofread_mode=proofread_mode,
        stream=True,
        temperature=temperature,
    )
    output_text = ""
    logger.info(
        "AI responses stream started profile_id=%s model=%s proofread_mode=%s temperature=%s text_len=%s max_output_tokens=%s",
        profile.id,
        profile.model,
        proofread_mode,
        payload["temperature"],
        len(text),
        payload["max_output_tokens"],
    )

    _debug_log_json("AI responses stream request payload", payload)
    headers = _auth_headers(profile)
    timeout_seconds, timeout_meta = _responses_timeout(settings, profile, payload, proofread_mode, reasoning_enabled)
    logger.info(
        "AI responses stream timeout estimated profile_id=%s model=%s estimated_input_tokens=%s estimated_output_tokens=%s estimated_total_tokens=%s token_units=%s timeout_seconds=%s proofread_mode=%s reasoning_enabled=%s",
        profile.id,
        profile.model,
        timeout_meta["estimated_input_tokens"],
        timeout_meta["estimated_output_tokens"],
        timeout_meta["estimated_total_tokens"],
        timeout_meta["token_units"],
        timeout_seconds,
        proofread_mode,
        reasoning_enabled,
    )

    async with httpx.AsyncClient(timeout=timeout_seconds) as client:
        async with client.stream(
            "POST",
            _responses_url(profile),
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


def estimate_proofread_request_timeout(
    text: str,
    book: BookInfo,
    settings: Settings | None = None,
    ai_profile_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
    temperature: float = DEFAULT_TEMPERATURE,
) -> AIRequestTimeoutEstimate:
    settings = settings or get_settings()
    try:
        profile = resolve_ai_profile(settings, ai_profile_id)
        provider_api = resolve_provider_api(profile, provider_api)
    except AIProfileError as exc:
        raise AIClientError(str(exc)) from exc

    if provider_api == "chat":
        payload = _build_chat_payload(
            text,
            book,
            settings,
            profile,
            proofread_mode=proofread_mode,
            reasoning_enabled=reasoning_enabled,
            temperature=temperature,
        )
        timeout_seconds, timeout_meta = _chat_timeout(settings, profile, payload, proofread_mode, reasoning_enabled)
    else:
        payload = _build_responses_payload(
            text,
            book,
            settings,
            profile,
            proofread_mode=proofread_mode,
            temperature=temperature,
        )
        timeout_seconds, timeout_meta = _responses_timeout(settings, profile, payload, proofread_mode, reasoning_enabled)

    return AIRequestTimeoutEstimate(
        timeout_seconds=timeout_seconds,
        estimated_input_tokens=timeout_meta["estimated_input_tokens"],
        estimated_output_tokens=timeout_meta["estimated_output_tokens"],
        estimated_total_tokens=timeout_meta["estimated_total_tokens"],
        token_units=timeout_meta["token_units"],
        proofread_mode=proofread_mode,
        reasoning_enabled=reasoning_enabled,
    )


def _ensure_responses_api(profile: AIProfile) -> None:
    _ensure_api_key(profile)


async def _proofread_with_chat(
    text: str,
    book: BookInfo,
    settings: Settings,
    profile: AIProfile,
    proofread_mode: ProofreadMode,
    reasoning_enabled: bool,
    temperature: float,
) -> AIProofreadResult:
    _ensure_api_key(profile)
    dialect = _chat_dialect(profile)
    payload = _build_chat_payload(
        text,
        book,
        settings,
        profile,
        proofread_mode=proofread_mode,
        reasoning_enabled=reasoning_enabled,
        temperature=temperature,
        dialect=dialect,
    )
    logger.info(
        "AI chat request started profile_id=%s model=%s dialect=%s proofread_mode=%s reasoning_enabled=%s temperature=%s text_len=%s output_token_limit=%s",
        profile.id,
        profile.model,
        dialect,
        proofread_mode,
        reasoning_enabled,
        payload.get("temperature"),
        len(text),
        _chat_output_token_limit(payload),
    )

    _debug_log_json("AI chat request payload", payload)
    headers = _auth_headers(profile)
    timeout_seconds, timeout_meta = _chat_timeout(settings, profile, payload, proofread_mode, reasoning_enabled)
    logger.info(
        "AI chat request timeout estimated profile_id=%s model=%s estimated_input_tokens=%s estimated_output_tokens=%s estimated_total_tokens=%s token_units=%s timeout_seconds=%s proofread_mode=%s reasoning_enabled=%s",
        profile.id,
        profile.model,
        timeout_meta["estimated_input_tokens"],
        timeout_meta["estimated_output_tokens"],
        timeout_meta["estimated_total_tokens"],
        timeout_meta["token_units"],
        timeout_seconds,
        proofread_mode,
        reasoning_enabled,
    )

    async with httpx.AsyncClient(timeout=timeout_seconds) as client:
        response = await client.post(
            _chat_url(profile),
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
    profile: AIProfile,
    proofread_mode: ProofreadMode = "fast",
    stream: bool = False,
    temperature: float = DEFAULT_TEMPERATURE,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "model": profile.model,
        "input": f"{_build_system_prompt(proofread_mode)}\n\n{_build_user_prompt(text, book)}",
        "temperature": temperature,
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
    profile: AIProfile,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
    temperature: float = DEFAULT_TEMPERATURE,
    dialect: ChatDialect | None = None,
) -> dict[str, Any]:
    resolved_dialect = dialect or _chat_dialect(profile)
    payload: dict[str, Any] = {
        "model": profile.model,
        "messages": [
            {"role": "system", "content": _build_system_prompt(proofread_mode)},
            {"role": "user", "content": _build_user_prompt(text, book)},
        ],
    }

    if resolved_dialect == "xiaomimimo":
        payload["max_completion_tokens"] = _max_tokens_for_mode(settings, proofread_mode)
        payload["thinking"] = {"type": "enabled" if reasoning_enabled else "disabled"}
        payload["response_format"] = {"type": "json_object"}
        if not reasoning_enabled:
            payload["temperature"] = temperature
        return payload

    if resolved_dialect == "deepseek":
        payload["max_tokens"] = _max_tokens_for_mode(settings, proofread_mode)
        payload["thinking"] = {"type": "enabled" if reasoning_enabled else "disabled"}
        payload["response_format"] = {"type": "json_object"}
        if reasoning_enabled:
            payload["reasoning_effort"] = "high"
        else:
            payload["temperature"] = temperature
        return payload

    payload["temperature"] = temperature
    payload["max_tokens"] = _max_tokens_for_mode(settings, proofread_mode)
    payload["reasoning"] = {"enabled": reasoning_enabled}
    # "response_format": {"type": "json_object"},
    return payload


def _chat_dialect(profile: AIProfile) -> ChatDialect:
    host = urlparse(profile.api_base_url).hostname or ""
    if host.lower() == "api.deepseek.com":
        return "deepseek"

    if host.lower() == "api.xiaomimimo.com":
        return "xiaomimimo"

    return "default"


def _chat_output_token_limit(payload: dict[str, Any]) -> Any:
    return payload.get("max_tokens", payload.get("max_completion_tokens"))


def _responses_timeout(
    settings: Settings,
    profile: AIProfile,
    payload: dict[str, Any],
    proofread_mode: ProofreadMode,
    reasoning_enabled: bool,
) -> tuple[float, dict[str, int]]:
    input_tokens = _count_text_tokens(profile.model, _coerce_string(payload.get("input")))
    output_token_limit = _coerce_int(payload.get("max_output_tokens"))
    return _dynamic_timeout(settings, input_tokens, proofread_mode, reasoning_enabled, output_token_limit)


def _chat_timeout(
    settings: Settings,
    profile: AIProfile,
    payload: dict[str, Any],
    proofread_mode: ProofreadMode,
    reasoning_enabled: bool,
) -> tuple[float, dict[str, int]]:
    messages = payload.get("messages")
    input_tokens = _count_chat_tokens(profile.model, messages if isinstance(messages, list) else [])
    output_token_limit = _coerce_int(_chat_output_token_limit(payload))
    return _dynamic_timeout(settings, input_tokens, proofread_mode, reasoning_enabled, output_token_limit)


def _dynamic_timeout(
    settings: Settings,
    input_tokens: int,
    proofread_mode: ProofreadMode,
    reasoning_enabled: bool,
    output_token_limit: int = 0,
) -> tuple[float, dict[str, int]]:
    estimated_input_tokens = max(input_tokens, 0)
    estimated_output_tokens = max(output_token_limit, 0)
    estimated_total_tokens = estimated_input_tokens + estimated_output_tokens
    token_units = max(1, math.ceil(max(estimated_total_tokens, 1) / 1000))
    seconds_per_1k = (
        settings.ai_thinking_timeout_seconds_per_1k_tokens
        if proofread_mode == "thinking" or reasoning_enabled
        else settings.ai_fast_timeout_seconds_per_1k_tokens
    )
    raw_timeout = settings.ai_request_timeout_base_seconds + token_units * seconds_per_1k
    timeout = min(
        settings.resolved_ai_request_timeout_max_seconds,
        max(settings.ai_request_timeout_min_seconds, raw_timeout),
    )
    return timeout, {
        "estimated_input_tokens": estimated_input_tokens,
        "estimated_output_tokens": estimated_output_tokens,
        "estimated_total_tokens": estimated_total_tokens,
        "token_units": token_units,
    }


def _coerce_int(value: Any) -> int:
    if isinstance(value, bool):
        return 0
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        try:
            return int(value)
        except ValueError:
            return 0
    return 0


def _count_text_tokens(model: str, text: str) -> int:
    if model in _TOKENIZER_UNAVAILABLE_MODELS:
        return _approximate_token_count(text)
    try:
        return len(_token_encoding(model).encode(text))
    except Exception as exc:
        _TOKENIZER_UNAVAILABLE_MODELS.add(model)
        logger.warning(
            "tiktoken encoding unavailable, using approximate token count model=%s error_type=%s",
            model,
            type(exc).__name__,
        )
        return _approximate_token_count(text)


def _count_chat_tokens(model: str, messages: list[Any]) -> int:
    if model in _TOKENIZER_UNAVAILABLE_MODELS:
        content = "\n".join(
            _coerce_string(message.get("content", ""))
            for message in messages
            if isinstance(message, dict)
        )
        return _approximate_token_count(content) + 4 * len(messages) + 2
    try:
        encoding = _token_encoding(model)
        content_tokens = 0
        message_count = 0
        for message in messages:
            if not isinstance(message, dict):
                continue
            message_count += 1
            content_tokens += len(encoding.encode(_coerce_string(message.get("content", ""))))
        return content_tokens + 4 * message_count + 2
    except Exception as exc:
        _TOKENIZER_UNAVAILABLE_MODELS.add(model)
        logger.warning(
            "tiktoken chat encoding unavailable, using approximate token count model=%s error_type=%s",
            model,
            type(exc).__name__,
        )
        content = "\n".join(
            _coerce_string(message.get("content", ""))
            for message in messages
            if isinstance(message, dict)
        )
        return _approximate_token_count(content) + 4 * len(messages) + 2


def _token_encoding(model: str) -> tiktoken.Encoding:
    try:
        return tiktoken.encoding_for_model(model)
    except KeyError:
        return tiktoken.get_encoding("cl100k_base")


def _approximate_token_count(text: str) -> int:
    return max(1, math.ceil(len(text) / 2))


def _build_system_prompt(proofread_mode: ProofreadMode) -> str:
    return "\n\n".join([BASE_SYSTEM_PROMPT, MODE_PROMPTS[proofread_mode]])


def _build_user_prompt(text: str, book: BookInfo) -> str:
    book_context = json.dumps(
        book.model_dump(),
        ensure_ascii=False,
        separators=(",", ":"),
    )
    prompt_parts = [
        f"""
<book>
{book_context}
</book>
""".strip()
    ]

    prompt_parts.append(
        f"""
<text>
{text}
</text>
""".strip()
    )
    return "\n\n".join(prompt_parts)


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


def _responses_url(profile: AIProfile) -> str:
    return f"{profile.api_base_url.rstrip('/')}/responses"


def _chat_url(profile: AIProfile) -> str:
    return f"{profile.api_base_url.rstrip('/')}/chat/completions"


def _auth_headers(profile: AIProfile) -> dict[str, str]:
    return {"Authorization": f"Bearer {profile.api_key}"}


def _raise_for_provider_error(status_code: int, provider_api: ProviderAPI = "responses") -> None:
    if status_code < 400:
        return

    if provider_api == "responses" and status_code in {400, 404, 405}:
        logger.warning("AI provider responses API unsupported status_code=%s", status_code)
        raise AIClientError(f"AI provider does not support Responses API (HTTP {status_code})")

    logger.warning("AI provider returned error status_code=%s provider_api=%s", status_code, provider_api)
    raise AIClientError(f"AI provider returned HTTP {status_code}")


def _ensure_api_key(profile: AIProfile) -> None:
    if not profile.api_key:
        raise AIClientError(f"{profile.api_key_env} is not configured")


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

    for candidate in candidates:
        salvaged_issues = _salvage_issue_objects(candidate)
        if salvaged_issues:
            logger.warning(
                "AI provider response JSON required issue-level salvage salvaged_issue_count=%s",
                len(salvaged_issues),
            )
            return {"issues": salvaged_issues}

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


def _salvage_issue_objects(content: str) -> list[dict[str, Any]]:
    if '"issues"' not in content:
        return []

    starts = [match.start() for match in re.finditer(r'\{\s*"id"\s*:', content)]
    if not starts:
        return []

    issues: list[dict[str, Any]] = []
    for index, start in enumerate(starts):
        limit = starts[index + 1] if index + 1 < len(starts) else len(content)
        fragment = _extract_json_object_from_start(content, start, limit)
        if not fragment:
            fragment = content[start:limit].strip().rstrip(",")

        parsed = _parse_issue_fragment(fragment)
        if parsed is None:
            continue

        try:
            ProofreadIssue.model_validate(parsed)
        except ValidationError:
            continue

        issues.append(parsed)

    return issues


def _extract_json_object_from_start(content: str, start: int, limit: int) -> str | None:
    in_string = False
    escaped = False
    depth = 0

    for index in range(start, limit):
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


def _parse_issue_fragment(fragment: str) -> dict[str, Any] | None:
    candidates = [fragment]
    cleaned = _remove_trailing_commas(fragment)
    if cleaned != fragment:
        candidates.append(cleaned)

    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
        except json.JSONDecodeError:
            try:
                parsed = json.loads(candidate, strict=False)
            except json.JSONDecodeError:
                continue

        if isinstance(parsed, dict):
            return parsed

    return None


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
