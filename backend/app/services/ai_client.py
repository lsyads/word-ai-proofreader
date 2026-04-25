from __future__ import annotations

import json
from collections.abc import AsyncIterator
from dataclasses import dataclass
from typing import Any

import httpx
from pydantic import ValidationError

from app.schemas import ProofreadIssue
from app.settings import Settings, get_settings


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


SYSTEM_PROMPT = """
你是出版社责任编辑的中文审校助手。请只返回 JSON，不要返回 Markdown。
返回格式必须是：
{
  "issues": [
    {
      "id": "issue-1",
      "category": "typo",
      "severity": "low|medium|high",
      "original": "原文片段",
      "suggestion": "修改建议",
      "comment": "给责任编辑看的批注内容",
      "start": 0,
      "end": 4
    }
  ]
}
如果没有发现问题，返回 {"issues": []}。
""".strip()


async def proofread_with_ai(
    text: str,
    previous_response_id: str | None = None,
    settings: Settings | None = None,
) -> AIProofreadResult:
    settings = settings or get_settings()
    _ensure_responses_api(settings)

    payload = _build_responses_payload(text, settings, previous_response_id=previous_response_id)

    async with httpx.AsyncClient(timeout=settings.ai_request_timeout_seconds) as client:
        response = await client.post(
            _responses_url(settings),
            headers=_auth_headers(settings),
            json=payload,
        )

    _raise_for_provider_error(response.status_code)

    try:
        data = response.json()
    except ValueError as exc:
        raise AIClientError("AI provider response body was not valid JSON") from exc

    return _parse_response_payload(data)


async def stream_proofread_with_ai(
    text: str,
    previous_response_id: str | None = None,
    settings: Settings | None = None,
) -> AsyncIterator[AIStreamEvent]:
    settings = settings or get_settings()
    _ensure_responses_api(settings)

    payload = _build_responses_payload(
        text,
        settings,
        previous_response_id=previous_response_id,
        stream=True,
    )
    output_text = ""

    async with httpx.AsyncClient(timeout=settings.ai_request_timeout_seconds) as client:
        async with client.stream(
            "POST",
            _responses_url(settings),
            headers=_auth_headers(settings),
            json=payload,
        ) as response:
            _raise_for_provider_error(response.status_code)

            async for provider_event in _iter_sse_events(response.aiter_lines()):
                event_name = provider_event["event"]
                data = provider_event["data"]

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
    if not settings.ai_api_key:
        raise AIClientError("AI_API_KEY is not configured")

    if settings.ai_provider_api != "responses":
        raise AIClientError("AI_PROVIDER_API must be responses for native AI sessions")


def _build_responses_payload(
    text: str,
    settings: Settings,
    previous_response_id: str | None = None,
    stream: bool = False,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "model": settings.openai_model,
        "input": f"{SYSTEM_PROMPT}\n\n请审校以下 Word 选区文本：\n{text}",
        "temperature": 0.2,
        "max_output_tokens": settings.ai_max_tokens,
        "text": {"format": {"type": "json_object"}},
    }

    if previous_response_id:
        payload["previous_response_id"] = previous_response_id

    if stream:
        payload["stream"] = True

    return payload


def _responses_url(settings: Settings) -> str:
    return f"{settings.openai_api_base_url.rstrip('/')}/responses"


def _auth_headers(settings: Settings) -> dict[str, str]:
    return {"Authorization": f"Bearer {settings.ai_api_key}"}


def _raise_for_provider_error(status_code: int) -> None:
    if status_code < 400:
        return

    if status_code in {400, 404, 405}:
        raise AIClientError(f"AI provider does not support native Responses session API (HTTP {status_code})")

    raise AIClientError(f"AI provider returned HTTP {status_code}")


def _parse_response_payload(data: dict[str, Any], fallback_content: str = "") -> AIProofreadResult:
    response_id = _coerce_optional_string(data.get("id"))
    content = fallback_content or _extract_response_output_text(data)
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
        return [ProofreadIssue.model_validate(issue) for issue in parsed["issues"]]
    except ValidationError as exc:
        raise AIClientError("AI provider response issues did not match the expected schema") from exc


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
