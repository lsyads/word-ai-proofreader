from __future__ import annotations

import json
from typing import Any

import httpx
from pydantic import ValidationError

from app.schemas import ProofreadIssue
from app.settings import Settings, get_settings


class AIClientError(RuntimeError):
    """Raised when the AI provider cannot return normalized proofread issues."""


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


async def proofread_with_ai(text: str, settings: Settings | None = None) -> list[ProofreadIssue]:
    settings = settings or get_settings()

    if not settings.ai_api_key:
        raise AIClientError("AI_API_KEY is not configured")

    base_url = settings.openai_api_base_url.rstrip("/")

    payload = {
        "model": settings.openai_model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": text},
        ],
        "temperature": 0.2,
        "max_tokens": settings.ai_max_tokens,
        "response_format": {"type": "json_object"},
    }

    async with httpx.AsyncClient(timeout=settings.ai_request_timeout_seconds) as client:
        response = await client.post(
            f"{base_url}/chat/completions",
            headers={"Authorization": f"Bearer {settings.ai_api_key}"},
            json=payload,
        )

    if response.status_code >= 400:
        raise AIClientError(f"AI provider returned HTTP {response.status_code}")

    try:
        data = response.json()
    except ValueError as exc:
        raise AIClientError("AI provider response body was not valid JSON") from exc

    content = _extract_message_content(data)
    return _parse_issues(content)


def _extract_message_content(data: dict[str, Any]) -> str:
    try:
        return data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise AIClientError("AI provider response did not include message content") from exc


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
