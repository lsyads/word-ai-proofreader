import asyncio

import httpx
import pytest

from app.schemas import ProofreadIssue
from app.services.ai_client import AIClientError, proofread_with_ai
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


class FakeAsyncClient:
    calls = []
    response = FakeResponse()

    def __init__(self, timeout):
        self.timeout = timeout

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc, tb):
        return False

    async def post(self, url, headers, json):
        self.__class__.calls.append(
            {
                "url": url,
                "headers": headers,
                "json": json,
                "timeout": self.timeout,
            }
        )
        return self.__class__.response


def test_proofread_with_ai_sends_openai_compatible_payload(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(
        payload={
            "choices": [
                {
                    "message": {
                        "content": (
                            '{"issues":[{"id":"issue-1","category":"typo","severity":"low",'
                            '"original":"错字","suggestion":"改字","comment":"请修正","start":0,"end":2}]}'
                        )
                    }
                }
            ]
        }
    )
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    settings = Settings(
        AI_API_KEY="test-key",
        OPENAI_API_BASE_URL="https://example.test/v1",
        OPENAI_MODEL="test-model",
        AI_REQUEST_TIMEOUT_SECONDS=12,
        AI_MAX_TOKENS=345,
    )

    issues = asyncio.run(proofread_with_ai("这是一段文本。", settings=settings))

    assert issues == [
        ProofreadIssue(
            id="issue-1",
            category="typo",
            severity="low",
            original="错字",
            suggestion="改字",
            comment="请修正",
            start=0,
            end=2,
        )
    ]

    call = FakeAsyncClient.calls[0]
    assert call["url"] == "https://example.test/v1/chat/completions"
    assert call["headers"] == {"Authorization": "Bearer test-key"}
    assert call["timeout"] == 12
    assert call["json"]["model"] == "test-model"
    assert call["json"]["temperature"] == 0.2
    assert call["json"]["max_tokens"] == 345
    assert call["json"]["response_format"] == {"type": "json_object"}
    assert call["json"]["messages"][0]["role"] == "system"
    assert call["json"]["messages"][1] == {"role": "user", "content": "这是一段文本。"}


def test_proofread_with_ai_raises_for_http_error(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(status_code=500, payload={})
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    settings = Settings(AI_API_KEY="test-key")

    with pytest.raises(AIClientError, match="HTTP 500"):
        asyncio.run(proofread_with_ai("文本", settings=settings))


def test_proofread_with_ai_raises_for_non_json_response_body(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(json_error=ValueError("not json"))
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    settings = Settings(AI_API_KEY="test-key")

    with pytest.raises(AIClientError, match="response body was not valid JSON"):
        asyncio.run(proofread_with_ai("文本", settings=settings))


def test_proofread_with_ai_raises_for_invalid_issue_schema(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(
        payload={"choices": [{"message": {"content": '{"issues":[{"id":"issue-1"}]}'}}]}
    )
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    settings = Settings(AI_API_KEY="test-key")

    with pytest.raises(AIClientError, match="issues did not match"):
        asyncio.run(proofread_with_ai("文本", settings=settings))
