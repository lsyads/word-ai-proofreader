import asyncio
import json
import logging

import httpx
import pytest

from app.schemas import BookInfo, ProofreadIssue
from app.services.ai_client import AIClientError, V2PromptContext, proofread_with_ai, stream_proofread_with_ai
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
                            '"original":"错字","replacement":"改字","suggestion":"改字","start":0,"end":2}]}'
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
                        '"original":"错字","replacement":"改字","suggestion":"改字"}]}'
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
    )


def book():
    return BookInfo(title="测试书名", introduction="这是一部测试图书。")


def test_proofread_with_ai_sends_responses_payload(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=response_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    result = asyncio.run(proofread_with_ai("这是一段文本。", book(), settings=settings()))

    assert result.issues == [
        ProofreadIssue(
            id="issue-1",
            category="typo",
            severity="low",
            original="错字",
            replacement="改字",
            suggestion="改字",
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
    assert call["json"]["max_output_tokens"] == 8192
    assert call["json"]["text"] == {"format": {"type": "json_object"}}
    assert "previous_response_id" not in call["json"]
    assert "这是一段文本。" in call["json"]["input"]
    assert "<text>\n这是一段文本。\n</text>" in call["json"]["input"]
    assert '"title":"测试书名"' in call["json"]["input"]
    assert '"introduction":"这是一部测试图书。"' in call["json"]["input"]
    assert "replacement" in call["json"]["input"]
    assert "comment" not in call["json"]["input"]


def test_proofread_with_ai_uses_thinking_token_limit_for_responses(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=response_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    result = asyncio.run(proofread_with_ai("文本", book(), settings=settings(), proofread_mode="thinking"))

    assert result.response_id == "resp-1"
    assert FakeAsyncClient.calls[0]["json"]["max_output_tokens"] == 16384
    assert "深度审校" in FakeAsyncClient.calls[0]["json"]["input"]


def test_proofread_with_ai_uses_custom_temperature_for_responses(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=response_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    asyncio.run(proofread_with_ai("文本", book(), settings=settings(), temperature=0.7))

    assert FakeAsyncClient.calls[0]["json"]["temperature"] == 0.7


def test_proofread_with_ai_includes_v2_prompt_context(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=response_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    context = V2PromptContext(
        review_goal="重点检查术语一致性。",
        source_type="docx",
        pass_name="terminology_pass",
        document_map_summary="text_len=100; blocks=2; chunks=1",
        memory_items=[{"kind": "preference", "key": "approved_issue_categories", "value": "style"}],
        style_rules=["统一术语。"],
    )

    asyncio.run(proofread_with_ai("文本", book(), settings=settings(), v2_context=context))

    payload_input = FakeAsyncClient.calls[0]["json"]["input"]
    assert "V2.1 Agent 工作台要求" in payload_input
    assert "重点检查术语一致性。" in payload_input
    assert "terminology_pass" in payload_input
    assert "text_len=100; blocks=2; chunks=1" in payload_input
    assert "approved_issue_categories" in payload_input


def test_proofread_with_ai_sends_chat_payload(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=chat_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    result = asyncio.run(proofread_with_ai("这是一段文本。", book(), provider_api="chat", settings=settings()))

    assert result.response_id == "chatcmpl-1"
    assert result.issues == [
        ProofreadIssue(
            id="issue-1",
            category="typo",
            severity="low",
            original="错字",
            replacement="改字",
            suggestion="改字",
        )
    ]

    call = FakeAsyncClient.calls[0]
    assert call["url"] == "https://example.test/v1/chat/completions"
    assert call["headers"] == {"Authorization": "Bearer test-key"}
    assert call["json"]["max_tokens"] == 8192
    assert call["json"]["messages"][0]["role"] == "system"
    assert "replacement" in call["json"]["messages"][0]["content"]
    assert "comment" not in call["json"]["messages"][0]["content"]
    assert call["json"]["reasoning"] == {"enabled": False}
    assert "max_completion_tokens" not in call["json"]
    assert "thinking" not in call["json"]
    assert "response_format" not in call["json"]
    assert call["json"]["messages"][1]["content"].endswith("<text>\n这是一段文本。\n</text>")
    assert '"title":"测试书名"' in call["json"]["messages"][1]["content"]


def test_proofread_with_ai_uses_selected_profile(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=chat_payload())
    monkeypatch.setenv("OPENROUTER_API_KEY", "profile-key")
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)
    profile_settings = Settings(
        AI_PROFILES_JSON=json.dumps(
            [
                {
                    "id": "openrouter-qwen",
                    "label": "OpenRouter / Qwen",
                    "api_base_url": "https://openrouter.ai/api/v1",
                    "api_key_env": "OPENROUTER_API_KEY",
                    "model": "qwen/test",
                    "default_api": "chat",
                    "supported_apis": ["chat"],
                }
            ]
        )
    )

    result = asyncio.run(
        proofread_with_ai(
            "这是一段文本。",
            book(),
            ai_profile_id="openrouter-qwen",
            settings=profile_settings,
        )
    )

    assert result.response_id == "chatcmpl-1"
    call = FakeAsyncClient.calls[0]
    assert call["url"] == "https://openrouter.ai/api/v1/chat/completions"
    assert call["headers"] == {"Authorization": "Bearer profile-key"}
    assert call["json"]["model"] == "qwen/test"


def test_proofread_with_ai_can_enable_chat_reasoning(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=chat_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    asyncio.run(
        proofread_with_ai(
            "这是一段文本。",
            book(),
            provider_api="chat",
            reasoning_enabled=True,
            settings=settings(),
        )
    )

    assert FakeAsyncClient.calls[0]["json"]["reasoning"] == {"enabled": True}


def test_proofread_with_ai_uses_custom_temperature_for_chat(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=chat_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    asyncio.run(
        proofread_with_ai(
            "这是一段文本。",
            book(),
            provider_api="chat",
            temperature=0.9,
            settings=settings(),
        )
    )

    assert FakeAsyncClient.calls[0]["json"]["temperature"] == 0.9


def test_proofread_with_ai_uses_xiaomimimo_chat_payload(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=chat_payload())
    monkeypatch.setenv("MIMO_API_KEY", "mimo-key")
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)
    profile_settings = Settings(
        AI_PROFILES_JSON=json.dumps(
            [
                {
                    "id": "xiaomi-mimo",
                    "label": "Xiaomi MiMo",
                    "api_base_url": "https://api.xiaomimimo.com/v1",
                    "api_key_env": "MIMO_API_KEY",
                    "model": "mimo-v2.5-pro",
                    "default_api": "chat",
                    "supported_apis": ["chat"],
                }
            ]
        )
    )

    result = asyncio.run(
        proofread_with_ai(
            "这是一段文本。",
            book(),
            ai_profile_id="xiaomi-mimo",
            provider_api="chat",
            settings=profile_settings,
        )
    )

    assert result.response_id == "chatcmpl-1"
    call = FakeAsyncClient.calls[0]
    assert call["url"] == "https://api.xiaomimimo.com/v1/chat/completions"
    assert call["headers"] == {"Authorization": "Bearer mimo-key"}
    assert call["json"]["model"] == "mimo-v2.5-pro"
    assert call["json"]["max_completion_tokens"] == 8192
    assert call["json"]["thinking"] == {"type": "disabled"}
    assert call["json"]["response_format"] == {"type": "json_object"}
    assert "max_tokens" not in call["json"]
    assert "reasoning" not in call["json"]


def test_proofread_with_ai_enables_xiaomimimo_thinking(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=chat_payload())
    monkeypatch.setenv("MIMO_API_KEY", "mimo-key")
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)
    profile_settings = Settings(
        AI_PROFILES_JSON=json.dumps(
            [
                {
                    "id": "xiaomi-mimo",
                    "label": "Xiaomi MiMo",
                    "api_base_url": "https://api.xiaomimimo.com/v1",
                    "api_key_env": "MIMO_API_KEY",
                    "model": "mimo-v2.5-pro",
                    "default_api": "chat",
                    "supported_apis": ["chat"],
                }
            ]
        )
    )

    asyncio.run(
        proofread_with_ai(
            "这是一段文本。",
            book(),
            ai_profile_id="xiaomi-mimo",
            provider_api="chat",
            proofread_mode="thinking",
            reasoning_enabled=True,
            settings=profile_settings,
        )
    )

    call = FakeAsyncClient.calls[0]
    assert call["json"]["max_completion_tokens"] == 16384
    assert call["json"]["thinking"] == {"type": "enabled"}
    assert "max_tokens" not in call["json"]
    assert "reasoning" not in call["json"]


def test_proofread_with_ai_raises_for_responses_unsupported(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(status_code=404, payload={})
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    with pytest.raises(AIClientError, match="Responses API"):
        asyncio.run(proofread_with_ai("文本", book(), settings=settings()))


def test_proofread_with_ai_raises_for_non_json_response_body(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(json_error=ValueError("not json"))
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    with pytest.raises(AIClientError, match="response body was not valid JSON"):
        asyncio.run(proofread_with_ai("文本", book(), settings=settings()))


def test_proofread_with_ai_raises_for_invalid_issue_schema(monkeypatch):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=response_payload(content='{"issues":[{"id":"issue-1"}]}'))
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)

    with pytest.raises(AIClientError, match="issues did not match"):
        asyncio.run(proofread_with_ai("文本", book(), settings=settings()))


def test_proofread_with_ai_accepts_unescaped_tabs_in_provider_json(monkeypatch, caplog):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(
        payload=chat_payload(
            content=(
                '{"issues":[{"id":"issue-1","category":"fact","severity":"high",'
                '"original":"空腹血糖受损\t\tADA\t5.6~7.0",'
                '"replacement":null,"suggestion":"需人工核查表格内容。"}]}'
            )
        )
    )
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)
    caplog.set_level(logging.INFO, logger="app")

    result = asyncio.run(proofread_with_ai("文本", book(), provider_api="chat", settings=settings()))

    assert result.issues[0].original == "空腹血糖受损\t\tADA\t5.6~7.0"
    assert "required lenient parsing" in caplog.text


def test_proofread_with_ai_salvages_valid_issues_when_one_issue_is_malformed(monkeypatch, caplog):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(
        payload=chat_payload(
            content="""{
  "issues": [
    {
      "id": "issue-1",
      "category": "fact",
      "severity": "high",
      "original": "用HCl调整溶液pH值至5.0",
      "replacement": null,
      "suggestion": "CTAB抽提液pH应为8.0，此处调至5.0与要求矛盾，需人工核查。"
    },
    {
      "id": "issue-2",
      "category": "fact",
      "severity": "high",
      "original": "1 μg/mL的双链DNA溶液在260 nm处的吸光度约为0.020",
      "replacement": null,
      "suggestion": "模型在这个字段里跑偏了。
      "original": "即可得到10 mM Tris - HCl（pH= 8.0）和1 mM EDTA的 TE 缓冲液",
      "replacement": null,
      "suggestion": "1.21g Tris定容至100mL得到的是100mM Tris，不是10mM，需人工核查。"
    },
    {
      "id": "issue-3",
      "category": "typo",
      "severity": "low",
      "original": "50 mmol/EDTA（pH=8.0）",
      "replacement": "50 mmol/L EDTA（pH=8.0）",
      "suggestion": "EDTA浓度单位漏写“L”，属于漏字错误。"
    }
  ]
}"""
        )
    )
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)
    caplog.set_level(logging.INFO, logger="app")

    result = asyncio.run(proofread_with_ai("文本", book(), provider_api="chat", settings=settings()))

    assert [issue.id for issue in result.issues] == ["issue-1", "issue-3"]
    assert result.issues[1].replacement == "50 mmol/L EDTA（pH=8.0）"
    assert "required issue-level salvage" in caplog.text


def test_proofread_with_ai_extracts_json_from_markdown_and_explanatory_text(monkeypatch, caplog):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(
        payload=response_payload(
            content=(
                "下面是审校结果：\n```json\n"
                '{"issues":[{"id":"issue-1","category":"typo","severity":"low",'
                '"original":"错字","replacement":"改字","suggestion":"应改为改字",}],}\n'
                "```\n请查收。"
            )
        )
    )
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)
    caplog.set_level(logging.INFO, logger="app")

    result = asyncio.run(proofread_with_ai("文本", book(), settings=settings()))

    assert result.issues[0].original == "错字"
    assert result.issues[0].replacement == "改字"
    assert "required cleanup" in caplog.text


def test_proofread_with_ai_debug_logs_provider_payloads_without_key(monkeypatch, caplog):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=response_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)
    caplog.set_level(logging.DEBUG, logger="app")

    asyncio.run(proofread_with_ai("调试文本", book(), settings=settings()))

    logs = caplog.text
    assert "AI responses request payload" in logs
    assert "AI responses response payload" in logs
    assert "调试文本" in logs
    assert "测试书名" in logs
    assert "test-key" not in logs
    assert "Authorization" not in logs
    assert "Bearer" not in logs


def test_proofread_with_ai_info_logs_provider_response_without_request_body_or_key(monkeypatch, caplog):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=response_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)
    caplog.set_level(logging.INFO, logger="app")

    asyncio.run(proofread_with_ai("请求正文不应出现在 INFO", book(), settings=settings()))

    logs = caplog.text
    assert "AI responses request started" in logs
    assert "AI responses request payload" not in logs
    assert "AI responses response payload" in logs
    assert "错字" in logs
    assert "请求正文不应出现在 INFO" not in logs
    assert "test-key" not in logs
    assert "Authorization" not in logs
    assert "Bearer" not in logs


def test_proofread_with_ai_info_logs_chat_provider_response(monkeypatch, caplog):
    FakeAsyncClient.calls = []
    FakeAsyncClient.response = FakeResponse(payload=chat_payload())
    monkeypatch.setattr(httpx, "AsyncClient", FakeAsyncClient)
    caplog.set_level(logging.INFO, logger="app")

    asyncio.run(proofread_with_ai("请求正文不应出现在 INFO", book(), provider_api="chat", settings=settings()))

    logs = caplog.text
    assert "AI chat response payload" in logs
    assert "错字" in logs
    assert "请求正文不应出现在 INFO" not in logs
    assert "test-key" not in logs


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
                    '"original":"错字","replacement":"改字","suggestion":"改字","start":0,"end":2}]}'
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

    events = asyncio.run(_collect_stream_events(stream_proofread_with_ai("文本", book(), settings=settings())))

    assert [event.event for event in events] == ["status", "status", "status", "result", "status"]
    assert events[3].data["response_id"] == "resp-stream"
    assert events[3].data["issues"][0]["id"] == "issue-1"
    assert FakeAsyncClient.calls[0]["json"]["stream"] is True


def test_stream_proofread_with_ai_rejects_chat_mode():
    with pytest.raises(AIClientError, match="not SSE"):
        asyncio.run(_collect_stream_events(stream_proofread_with_ai("文本", book(), settings=settings(), provider_api="chat")))


async def _collect_stream_events(stream):
    return [event async for event in stream]
