from __future__ import annotations

from collections.abc import AsyncIterator

from app.schemas import ProofreadIssue
from app.services.ai_client import AIClientError, AIStreamEvent, proofread_with_ai, stream_proofread_with_ai
from app.services.sessions import SessionNotFoundError, get_last_response_id, update_last_response_id
from app.settings import get_settings


async def proofread_text(text: str, session_id: str | None = None) -> list[ProofreadIssue]:
    settings = get_settings()

    if settings.ai_api_key:
        try:
            previous_response_id = get_last_response_id(session_id)
        except SessionNotFoundError as exc:
            raise AIClientError("AI session was not found. Please create a new session.") from exc

        result = await proofread_with_ai(text, previous_response_id=previous_response_id)
        update_last_response_id(session_id, result.response_id)
        return result.issues

    return build_mock_issues(text)


async def stream_proofread_text(text: str, session_id: str | None = None) -> AsyncIterator[AIStreamEvent]:
    settings = get_settings()

    yield AIStreamEvent("status", {"stage": "received", "message": "已接收选区文本。"})

    if not settings.ai_api_key:
        yield AIStreamEvent("status", {"stage": "calling_ai", "message": "当前未配置 AI_API_KEY，正在返回本地 mock 审校结果。"})
        yield AIStreamEvent("status", {"stage": "normalizing", "message": "正在整理结构化审校结果。"})
        yield AIStreamEvent("result", {"issues": [issue.model_dump() for issue in build_mock_issues(text)]})
        yield AIStreamEvent("status", {"stage": "completed", "message": "审校完成。"})
        return

    try:
        previous_response_id = get_last_response_id(session_id)
    except SessionNotFoundError as exc:
        raise AIClientError("AI session was not found. Please create a new session.") from exc

    yield AIStreamEvent("status", {"stage": "calling_ai", "message": "正在调用 AI 原生 Responses session。"})

    async for event in stream_proofread_with_ai(text, previous_response_id=previous_response_id):
        if event.event == "result":
            response_id = event.data.pop("response_id", None)
            update_last_response_id(session_id, response_id if isinstance(response_id, str) else None)

        yield event


def build_mock_issues(text: str) -> list[ProofreadIssue]:
    sample = text[: min(len(text), 12)]
    end = len(sample)

    return [
        ProofreadIssue(
            id="mock-issue-1",
            category="style",
            severity="medium",
            original=sample,
            suggestion="请结合上下文检查该表述是否准确、简洁，并确认是否符合出版物体例。",
            comment="本地 mock 审校结果：当前未配置 AI_API_KEY，已返回一条示例审校建议，用于验证 Word 批注闭环。",
            start=0,
            end=end,
        )
    ]
