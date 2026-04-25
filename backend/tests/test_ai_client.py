import asyncio
import json

import httpx
import pytest

from app.schemas import ProofreadIssue
from app.services.ai_client import AIClientError, proofread_with_ai, stream_proofread_with_ai
from app.settings import Settings


class FakeResponse:
    def __init__(self, status_code=200, payload=None, json_error: Exception | None = None):
        self.status_code = status_code
        self.payload = payload
        self.json_error = json_error

    def json(self):
        if self.json_error:
            raise self.json_error

        return self.payload


class FakeStreamResponse:
    def __init__(self, status_code=200, lines=None):
        self.status_code = status_code
        self.lines = lines or []

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc, tb):
        return False

    async def aiter_lines(self):
        for line in self.lines:
            yield line


class FakeAsyncClient:
    calls = []
    response = FakeResponse()
    stream_response = FakeStreamResponse()

    def __init__(self, timeout):
        self.timeout = timeout

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc, tb):
        return False

    async def post(self, url, headers, json):
        self.__class__.calls.append(
            {
                "method": "POST",
                "url": url,
                "headers": headers,
                "json": json,
                "timeout": self.timeout,
            }
        )
        return self.__class__.response

    def stream(self, method, url, headers, json):
        self.__class__.calls.append(
            {
                "method": method,
                "url": url,
                "headers": headers,
                "json": json,
                "timeout": self.timeout,
            }
        )
        return self.__class__.stream_response


def response_payload(response_id="resp-1", content=None):
    return {
        "id": response_id,
        "output": [
            {
                "type": "message",
                "content": [
                    {
                        "type": "output_text",
                        "text": content
                        or (
                            '{"issues":[{"id":"issue-1","category":"typo","severity":"low",'
                            '"original":"错字","replacement":"改字","suggestion":"改字","comment":"请修正","start":0,"end":2}]}'
                        ),
                    }
                ],
            }
        ],
    }


def chat_payload(response_id="chatcmpl-1", content=None):
    return {
        "id": response_id,
        "choices": [
            {
                "message": {
                    "content": content
                    or (
                        '{"issues":[{"id":"issue-1","category":"typo","severity":"low",'
                        '"original":"错字","replacement":"改字","suggestion":"改字","comment":"请修正"}]}'
                    )
                }
            }
        ],
    }


def settings():
    return Settings(
        AI_API_KEY="test-key",
        OPENAI_API_BASE_URL="https://example.test/v1",
        OPENAI_MODEL="test-model",
        AI_REQUEST_TIMEOUT_SECONDS=12,
        AI_MAX_TOKENS=345,
    )


def test_proofread_with_ai_sends_responses_payload(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=response_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    result = asyncio.run(proofread_with_ai("这是一段文本。", previous_response_id="resp-prev", settings=settings()))

    assert result.issues == [
        ProofreadIssue(
            id="issue-1",
            category="typo",
            severity="low",
            original="错字",
            replacement="改字",
            suggestion="改字",
            comment="请修正",
            start=0,
            end=2,
        )
    ]
    assert result.response_id == "resp-1"

    call = FakeAsyncClient.calls[0]
    assert call["url"] == "https://example.test/v1/responses"
    assert call["headers"] == {"Authorization": "Bearer test-key"}
    assert call["timeout"] == 12
    assert call["json"]["model"] == "test-model"
    assert call["json"]["temperature"] == 0.2
    assert call["json"]["max_output_tokens"] == 345
    assert call["json"]["text"] == {"format": {"type": "json_object"}}
    assert call["json"]["previous_response_id"] == "resp-prev"
    assert "这是一段文本。" in call["json"]["input"]
    assert "不要返回 start 或 end" in call["json"]["input"]
    assert "replacement" in call["json"]["input"]


def test_proofread_with_ai_uses_thinking_token_limit_for_responses(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=response_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    result = asyncio.run(proofread_with_ai("文本", settings=settings(), proofread_mode="thinking"))

    assert result.response_id == "resp-1"
    assert FakeAsyncClient.calls[0]["json"]["max_output_tokens"] == 1200
    assert "思考模式" in FakeAsyncClient.calls[0]["json"]["input"]


def test_proofread_with_ai_sends_chat_payload(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=chat_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    result = asyncio.run(proofread_with_ai("这是一段文本。", provider_api="chat", settings=settings()))

    assert result.response_id == "chatcmpl-1"
    assert result.issues == [
        ProofreadIssue(
            id="issue-1",
            category="typo",
            severity="low",
            original="错字",
            replacement="改字",
            suggestion="改字",
            comment="请修正",
        )
    ]

    call = FakeAsyncClient.calls[0]
    assert call["url"] == "https://example.test/v1/chat/completions"
    assert call["headers"] == {"Authorization": "Bearer test-key"}
    assert call["json"]["response_format"] == {"type": "json_object"}
    assert call["json"]["max_tokens"] == 345
    assert call["json"]["messages"][0]["role"] == "system"
    assert "不要返回 start 或 end" in call["json"]["messages"][0]["content"]
    assert "replacement" in call["json"]["messages"][0]["content"]
    assert call["json"]["messages"][1]["content"].endswith("这是一段文本。")


def test_proofread_with_ai_raises_for_native_responses_unsupported(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(status_code=404, payload={})
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    with pytest.raises(AIClientError, match="native Responses session API"):
        asyncio.run(proofread_with_ai("文本", settings=settings()))


def test_proofread_with_ai_raises_for_non_json_response_body(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(json_error=ValueError("not json"))
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    with pytest.raises(AIClientError, match="response body was not valid JSON"):
        asyncio.run(proofread_with_ai("文本", settings=settings()))


def test_proofread_with_ai_raises_for_invalid_issue_schema(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=response_payload(content='{"issues":[{"id":"issue-1"}]}'))
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    with pytest.raises(AIClientError, match="issues did not match"):
        asyncio.run(proofread_with_ai("文本", settings=settings()))


def test_stream_proofread_with_ai_converts_provider_sse(monkeypatch):
    provider_events = [
        ("response.created", {"type": "response.created"}),
        ("response.in_progress", {"type": "response.in_progress"}),
        (
            "response.output_text.done",
            {
                "type": "response.output_text.done",
                "text": (
                    '{"issues":[{"id":"issue-1","category":"typo","severity":"low",'
                    '"original":"错字","replacement":"改字","suggestion":"改字","comment":"请修正","start":0,"end":2}]}'
                ),
            },
        ),
        ("response.completed", {"type": "response.completed", "response": response_payload(response_id="resp-stream")}),
    ]
    lines = []
    for event_name, data in provider_events:
        lines.extend([f"event: {event_name}", f"data: {json.dumps(data)}", ""])

    FakeAsyncClient.calls = []
    FakeAsyncClient.stream_response = FakeStreamResponse(lines=lines)
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    events = asyncio.run(_collect_stream_events(stream_proofread_with_ai("文本", settings=settings())))

    assert [event.event for event in events] == ["status", "status", "status", "result", "status"]
    assert events[3].data["response_id"] == "resp-stream"
    assert events[3].data["issues"][0]["id"] == "issue-1"
    assert FakeAsyncClient.calls[0]["json"]["stream"] is True


async def _collect_stream_events(stream):
    return [event async for event in stream]
