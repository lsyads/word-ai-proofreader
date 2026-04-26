import json
import logging
import asyncio
from fastapi.testclient import TestClient

import app.services.proofread as proofread_service
from app.services import tasks as task_service
from app.main import app
from app.schemas import BookInfo, ChunkedProofreadRequest, ProofreadIssue
from app.services.ai_client import AIClientError, AIProofreadResult, AIStreamEvent
from app.services.sessions import clear_sessions_for_tests


client = TestClient(app)
BOOK = {"title": "测试书名", "introduction": "这是一部测试图书。"}


def setup_function():
    clear_sessions_for_tests()
    task_service.clear_tasks_for_tests()


def parse_sse_events(body: str):
    events = []

    for chunk in body.strip().split("\n\n"):
        event_name = None
        event_data = None

        for line in chunk.splitlines():
            if line.startswith("event: "):
                event_name = line.removeprefix("event: ")
            elif line.startswith("data: "):
                event_data = json.loads(line.removeprefix("data: "))

        events.append({"event": event_name, "data": event_data})

    return events


def proofread_payload(text: str, **extra):
    return {"text": text, "book": BOOK, **extra}


def test_health_returns_ok():
    response = client.get("/health")

    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_create_session_returns_unique_session_ids():
    first = client.post("/api/sessions")
    second = client.post("/api/sessions")

    assert first.status_code == 200
    assert second.status_code == 200
    assert first.json()["session_id"].startswith("session_")
    assert second.json()["session_id"].startswith("session_")
    assert first.json()["session_id"] != second.json()["session_id"]
    assert first.json()["created_at"]


def test_proofread_rejects_blank_text():
    response = client.post("/api/proofread", json=proofread_payload("   "))

    assert response.status_code == 422


def test_proofread_rejects_missing_book():
    response = client.post("/api/proofread", json={"text": "这是一段文本。"})

    assert response.status_code == 422


def test_proofread_rejects_blank_book_title():
    response = client.post(
        "/api/proofread",
        json={"text": "这是一段文本。", "book": {"title": "   ", "introduction": "介绍"}},
    )

    assert response.status_code == 422


def test_proofread_returns_mock_issue_without_api_key(monkeypatch):
    monkeypatch.delenv("AI_API_KEY", raising=False)

    response = client.post(
        "/api/proofread",
        json=proofread_payload("这是一段需要审校的文本。", context={"source": "word-addin"}),
    )

    assert response.status_code == 200
    payload = response.json()
    assert len(payload["issues"]) == 1

    issue = payload["issues"][0]
    assert issue["id"] == "mock-issue-1"
    assert issue["category"] == "style"
    assert issue["severity"] == "medium"
    assert issue["original"]
    assert issue["replacement"]
    assert issue["suggestion"]
    assert issue["start"] == 0
    assert isinstance(issue["end"], int)


def test_proofread_calculates_offsets_when_ai_omits_them(monkeypatch):
    async def fake_proofread_with_ai(
        text,
        book,
        provider_api=None,
        proofread_mode="fast",
    ):
        assert provider_api == "responses"
        assert proofread_mode == "fast"
        assert book.title == "测试书名"
        return AIProofreadResult(
            response_id="resp-1",
            issues=[
                ProofreadIssue(
                    id="ai-issue-1",
                    category="typo",
                    severity="high",
                    original="需要定位",
                    replacement="需要定位后的文本",
                    suggestion="建议",
                )
            ],
        )

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    response = client.post("/api/proofread", json=proofread_payload("这是一段需要定位的文本。"))

    assert response.status_code == 200
    issue = response.json()["issues"][0]
    assert issue["start"] == 4
    assert issue["end"] == 8
    assert issue["replacement"] == "需要定位后的文本"


def test_proofread_normalizes_empty_replacement_to_null(monkeypatch):
    async def fake_proofread_with_ai(
        text,
        book,
        provider_api=None,
        proofread_mode="fast",
    ):
        return AIProofreadResult(
            response_id="resp-1",
            issues=[
                ProofreadIssue(
                    id="ai-issue-1",
                    category="style",
                    severity="medium",
                    original="文本",
                    replacement="   ",
                    suggestion="建议",
                )
            ],
        )

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    response = client.post("/api/proofread", json=proofread_payload("这是一段文本。"))

    assert response.status_code == 200
    issue = response.json()["issues"][0]
    assert issue["replacement"] is None
    assert issue["start"] == 4
    assert issue["end"] == 6


def test_proofread_filters_whitespace_only_changes(monkeypatch):
    async def fake_proofread_with_ai(
        text,
        book,
        provider_api=None,
        proofread_mode="fast",
    ):
        return AIProofreadResult(
            response_id="resp-1",
            issues=[
                ProofreadIssue(
                    id="ai-issue-space",
                    category="style",
                    severity="low",
                    original="A B",
                    replacement="AB",
                    suggestion="删除多余空格。",
                ),
                ProofreadIssue(
                    id="ai-issue-typo",
                    category="typo",
                    severity="low",
                    original="错字",
                    replacement="改字",
                    suggestion="修正错别字。",
                ),
            ],
        )

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    response = client.post("/api/proofread", json=proofread_payload("这里有 A B，还有错字。"))

    assert response.status_code == 200
    issues = response.json()["issues"]
    assert [issue["id"] for issue in issues] == ["ai-issue-typo"]
    assert issues[0]["start"] == 10
    assert issues[0]["end"] == 12


def test_proofread_calculates_offsets_for_repeated_originals(monkeypatch):
    async def fake_proofread_with_ai(
        text,
        book,
        provider_api=None,
        proofread_mode="fast",
    ):
        return AIProofreadResult(
            response_id="resp-1",
            issues=[
                ProofreadIssue(
                    id="ai-issue-1",
                    category="style",
                    severity="medium",
                    original="重复",
                    suggestion="建议一",
                ),
                ProofreadIssue(
                    id="ai-issue-2",
                    category="style",
                    severity="medium",
                    original="重复",
                    suggestion="建议二",
                ),
            ],
        )

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    response = client.post("/api/proofread", json=proofread_payload("重复内容，重复内容。"))

    assert response.status_code == 200
    assert [(issue["start"], issue["end"]) for issue in response.json()["issues"]] == [(0, 2), (5, 7)]


def test_proofread_returns_null_offsets_when_original_is_missing(monkeypatch):
    async def fake_proofread_with_ai(
        text,
        book,
        provider_api=None,
        proofread_mode="fast",
    ):
        return AIProofreadResult(
            response_id="resp-1",
            issues=[
                ProofreadIssue(
                    id="ai-issue-1",
                    category="fact",
                    severity="low",
                    original="不存在",
                    suggestion="建议",
                )
            ],
        )

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    response = client.post("/api/proofread", json=proofread_payload("这是一段文本。"))

    assert response.status_code == 200
    issue = response.json()["issues"][0]
    assert issue["start"] is None
    assert issue["end"] is None


def test_proofread_uses_ai_client_when_api_key_is_configured(monkeypatch):
    async def fake_proofread_with_ai(
        text,
        book,
        provider_api=None,
        proofread_mode="fast",
    ):
        assert text == "这是一段需要真实审校的文本。"
        assert book.title == "测试书名"
        assert provider_api == "responses"
        assert proofread_mode == "thinking"
        return AIProofreadResult(
            response_id="resp-1",
            issues=[
                ProofreadIssue(
                    id="ai-issue-1",
                    category="typo",
                    severity="high",
                    original="真实",
                    suggestion="真实建议",
                    start=0,
                    end=2,
                )
            ],
        )

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    response = client.post(
        "/api/proofread",
        json=proofread_payload("这是一段需要真实审校的文本。", proofread_mode="thinking"),
    )

    assert response.status_code == 200
    assert response.json()["issues"][0]["id"] == "ai-issue-1"


def test_proofread_responses_mode_does_not_require_session(monkeypatch):
    call_count = 0

    async def fake_proofread_with_ai(
        text,
        book,
        provider_api=None,
        proofread_mode="fast",
    ):
        nonlocal call_count
        call_count += 1
        return AIProofreadResult(response_id=f"resp-{call_count}", issues=[])

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    session_id = client.post("/api/sessions").json()["session_id"]

    first = client.post("/api/proofread", json=proofread_payload("第一段文本。", session_id=session_id))
    second = client.post("/api/proofread", json=proofread_payload("第二段文本。", session_id=session_id))
    third = client.post("/api/proofread", json=proofread_payload("第三段文本。", session_id="missing-session"))

    assert first.status_code == 200
    assert second.status_code == 200
    assert third.status_code == 200
    assert call_count == 3


def test_proofread_chat_mode_does_not_require_session(monkeypatch):
    calls = []

    async def fake_proofread_with_ai(
        text,
        book,
        provider_api=None,
        proofread_mode="fast",
    ):
        calls.append(
            {
                "provider_api": provider_api,
                "proofread_mode": proofread_mode,
            }
        )
        return AIProofreadResult(response_id="chatcmpl-1", issues=[])

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    first = client.post(
        "/api/proofread",
        json=proofread_payload("第一段文本。", provider_api="chat", proofread_mode="thinking"),
    )
    second = client.post(
        "/api/proofread",
        json=proofread_payload("第二段文本。", provider_api="chat", session_id="missing-session"),
    )

    assert first.status_code == 200
    assert second.status_code == 200
    assert calls == [
        {"provider_api": "chat", "proofread_mode": "thinking"},
        {"provider_api": "chat", "proofread_mode": "fast"},
    ]


def test_proofread_converts_ai_client_error_to_502(monkeypatch):
    async def fake_proofread_with_ai(
        text,
        book,
        provider_api=None,
        proofread_mode="fast",
    ):
        raise AIClientError("AI provider returned HTTP 500")

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    response = client.post("/api/proofread", json=proofread_payload("这是一段文本。"))

    assert response.status_code == 502
    assert response.json() == {"detail": "AI provider returned HTTP 500"}


def test_proofread_debug_logs_request_and_response_without_api_key(monkeypatch, caplog):
    monkeypatch.delenv("AI_API_KEY", raising=False)
    caplog.set_level(logging.DEBUG, logger="app")

    response = client.post("/api/proofread", json=proofread_payload("调试文本。"))

    assert response.status_code == 200
    logs = caplog.text
    assert "proofread request body" in logs
    assert "proofread response body" in logs
    assert "调试文本。" in logs
    assert "测试书名" in logs
    assert "Authorization" not in logs
    assert "Bearer" not in logs


def test_proofread_info_logs_do_not_include_full_body(monkeypatch, caplog):
    monkeypatch.delenv("AI_API_KEY", raising=False)
    caplog.set_level(logging.INFO, logger="app")

    response = client.post("/api/proofread", json=proofread_payload("不应出现在 INFO 的正文。"))

    assert response.status_code == 200
    logs = caplog.text
    assert "proofread request received" in logs
    assert "proofread request body" not in logs
    assert "不应出现在 INFO 的正文。" not in logs


def test_proofread_stream_returns_status_events_and_result(monkeypatch):
    async def fake_stream_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
    ):
        assert text == "这是一段需要真实审校的文本。"
        assert book.title == "测试书名"
        assert session_id == "session-test"
        assert provider_api == "responses"
        assert proofread_mode == "thinking"
        yield AIStreamEvent("status", {"stage": "received", "message": "已接收选区文本。"})
        yield AIStreamEvent("status", {"stage": "calling_ai", "message": "正在调用 AI Responses API。"})
        yield AIStreamEvent("status", {"stage": "normalizing", "message": "已收到 AI 输出，正在解析结构化结果。"})
        yield AIStreamEvent(
            "result",
            {
                "issues": [
                    {
                        "id": "ai-issue-1",
                        "category": "typo",
                        "severity": "high",
                        "original": "真实",
                        "suggestion": "真实建议",
                        "start": 0,
                        "end": 2,
                    }
                ]
            },
        )
        yield AIStreamEvent("status", {"stage": "completed", "message": "审校完成。"})

    monkeypatch.setattr(proofread_service, "stream_proofread_text", fake_stream_proofread_text)

    response = client.post(
        "/api/proofread/stream",
        json=proofread_payload(
            "这是一段需要真实审校的文本。",
            session_id="session-test",
            provider_api="responses",
            proofread_mode="thinking",
        ),
    )

    assert response.status_code == 200
    assert response.headers["content-type"].startswith("text/event-stream")

    events = parse_sse_events(response.text)
    assert [event["event"] for event in events] == ["status", "status", "status", "result", "status"]
    assert [event["data"]["stage"] for event in events if event["event"] == "status"] == [
        "received",
        "calling_ai",
        "normalizing",
        "completed",
    ]
    assert events[3]["data"]["issues"][0]["id"] == "ai-issue-1"


def test_proofread_stream_rejects_chat_mode():
    response = client.post(
        "/api/proofread/stream",
        json=proofread_payload("这是一段文本。", provider_api="chat"),
    )

    assert response.status_code == 400
    assert response.json() == {
        "detail": "Chat mode uses /api/proofread with standard Chat Completions, not SSE."
    }


def test_proofread_stream_does_not_require_session(monkeypatch):
    async def fake_stream_proofread_with_ai(
        text,
        book,
        provider_api=None,
        proofread_mode="fast",
    ):
        assert text == "这是一段文本。"
        assert book.title == "测试书名"
        assert provider_api == "responses"
        assert proofread_mode == "fast"
        yield AIStreamEvent("result", {"issues": [], "response_id": "resp-1"})

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "stream_proofread_with_ai", fake_stream_proofread_with_ai)

    response = client.post(
        "/api/proofread/stream",
        json=proofread_payload("这是一段文本。", session_id="missing-session"),
    )

    assert response.status_code == 200
    events = parse_sse_events(response.text)
    assert [event["event"] for event in events] == ["status", "status", "result"]
    assert events[2]["data"] == {"issues": []}


def test_proofread_stream_returns_error_event_for_ai_client_error(monkeypatch):
    async def fake_stream_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
    ):
        yield AIStreamEvent("status", {"stage": "received", "message": "已接收选区文本。"})
        yield AIStreamEvent("status", {"stage": "calling_ai", "message": "正在调用 AI Responses API。"})
        raise AIClientError("AI provider returned HTTP 500")

    monkeypatch.setattr(proofread_service, "stream_proofread_text", fake_stream_proofread_text)

    response = client.post("/api/proofread/stream", json=proofread_payload("这是一段文本。"))

    assert response.status_code == 200

    events = parse_sse_events(response.text)
    assert [event["event"] for event in events] == ["status", "status", "error"]
    assert events[-1]["data"] == {"message": "AI provider returned HTTP 500"}


def test_chunked_proofread_returns_aggregated_global_offsets(monkeypatch):
    async def fake_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
    ):
        return [
            ProofreadIssue(
                id=f"issue-{text[0]}",
                category="style",
                severity="medium",
                original=text[:2],
                suggestion="建议",
                start=0,
                end=2,
            )
        ]

    monkeypatch.setattr(proofread_service, "proofread_text", fake_proofread_text)

    response = client.post(
        "/api/proofread/chunked",
        json=proofread_payload(
            ("甲" * 3000) + ("乙" * 3000),
            scope="document",
            chunk_size=3000,
        ),
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["scope"] == "document"
    assert payload["status"] == "succeeded"
    assert payload["total_chunks"] == 2
    assert payload["completed_chunks"] == 2
    assert payload["failed_chunks"] == 0
    assert [issue["chunk_index"] for issue in payload["issues"]] == [0, 1]
    assert [issue["global_start"] for issue in payload["issues"]] == [0, 3000]


def test_proofread_task_lifecycle_and_events(monkeypatch):
    async def fake_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
    ):
        return [
            ProofreadIssue(
                id=f"issue-{text[0]}",
                category="typo",
                severity="low",
                original=text[:1],
                suggestion="建议",
                start=0,
                end=1,
            )
        ]

    monkeypatch.setattr(task_service, "proofread_text", fake_proofread_text)

    create_response = client.post(
        "/api/proofread/tasks",
        json=proofread_payload(
            ("甲" * 3000) + ("乙" * 3000),
            scope="document",
            chunk_size=3000,
        ),
    )

    assert create_response.status_code == 200
    task_id = create_response.json()["task_id"]

    events_response = client.get(f"/api/proofread/tasks/{task_id}/events")
    assert events_response.status_code == 200
    events = parse_sse_events(events_response.text)
    event_names = [event["event"] for event in events]
    assert event_names[0] == "queued"
    assert "chunk_started" in event_names
    assert "chunk_completed" in event_names
    assert event_names[-1] == "completed"

    status_response = client.get(f"/api/proofread/tasks/{task_id}")
    assert status_response.status_code == 200
    payload = status_response.json()
    assert payload["status"] == "succeeded"
    assert payload["completed_chunks"] == 2
    assert payload["failed_chunks"] == 0
    assert len(payload["issues"]) == 2


def test_proofread_task_records_failed_chunk_and_continues(monkeypatch, caplog):
    async def fake_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
    ):
        if text.startswith("乙"):
            raise AIClientError("chunk failed")

        return [
            ProofreadIssue(
                id="issue-1",
                category="typo",
                severity="low",
                original=text[:1],
                suggestion="建议",
                start=0,
                end=1,
            )
        ]

    monkeypatch.setattr(task_service, "proofread_text", fake_proofread_text)
    caplog.set_level(logging.INFO, logger="app.services.tasks")

    create_response = client.post(
        "/api/proofread/tasks",
        json=proofread_payload(
            ("甲" * 3000) + ("乙" * 3000) + ("丙" * 3000),
            scope="document",
            chunk_size=3000,
        ),
    )

    task_id = create_response.json()["task_id"]
    events = parse_sse_events(client.get(f"/api/proofread/tasks/{task_id}/events").text)
    status_response = client.get(f"/api/proofread/tasks/{task_id}")

    assert status_response.status_code == 200
    payload = status_response.json()
    assert payload["status"] == "partial_succeeded"
    assert payload["completed_chunks"] == 2
    assert payload["failed_chunks"] == 1
    assert "chunk_failed" in [event["event"] for event in events]
    failed_event = next(event for event in events if event["event"] == "chunk_failed")
    assert failed_event["data"]["error_message"] == "chunk failed"
    assert failed_event["data"]["chunk_start"] == 3000
    assert failed_event["data"]["chunk_end"] == 6000
    assert failed_event["data"]["chunk_len"] == 3000
    assert isinstance(failed_event["data"]["elapsed_seconds"], float | int)
    assert "chunked proofread chunk failed" in caplog.text
    assert "chunk_index=1" in caplog.text
    assert "chunk_start=3000" in caplog.text
    assert "chunk_end=6000" in caplog.text
    assert "elapsed_seconds=" in caplog.text
    assert "chunk failed" in caplog.text
    assert "乙乙乙乙乙" not in caplog.text


def test_proofread_task_logs_unhandled_task_crash(monkeypatch, caplog):
    async def fake_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
    ):
        return [
            ProofreadIssue(
                id="issue-1",
                category="typo",
                severity="low",
                original=text[:1],
                suggestion="建议",
                start=0,
                end=1,
            )
        ]

    def crash_globalize_issues(chunk, issues):
        raise RuntimeError("globalize exploded")

    monkeypatch.setattr(task_service, "proofread_text", fake_proofread_text)
    monkeypatch.setattr(task_service.chunking, "globalize_issues", crash_globalize_issues)
    caplog.set_level(logging.INFO, logger="app.services.tasks")

    create_response = client.post(
        "/api/proofread/tasks",
        json=proofread_payload("甲" * 3000, scope="document", chunk_size=3000),
    )

    task_id = create_response.json()["task_id"]
    events = parse_sse_events(client.get(f"/api/proofread/tasks/{task_id}/events").text)
    status_response = client.get(f"/api/proofread/tasks/{task_id}")

    assert status_response.json()["status"] == "failed"
    assert events[-1]["event"] == "error"
    assert "chunked proofread task crashed" in caplog.text
    assert task_id in caplog.text
    assert "globalize exploded" in caplog.text


def test_proofread_task_emits_heartbeat(monkeypatch):
    monkeypatch.setattr(task_service, "HEARTBEAT_INTERVAL_SECONDS", 0.001)

    async def run_heartbeat():
        task = task_service.ProofreadTask(
            task_id="task-test",
            request=ChunkedProofreadRequest(
                text="甲" * 3000,
                book=BookInfo.model_validate(BOOK),
                scope="document",
            ),
            status="running",
            total_chunks=1,
        )
        chunk = task_service.chunking.ProofreadChunk(index=0, start=0, end=3000, text="甲" * 3000)
        heartbeat_task = asyncio.create_task(task_service._emit_heartbeats(task, chunk))
        await asyncio.sleep(0.003)
        heartbeat_task.cancel()
        try:
            await heartbeat_task
        except asyncio.CancelledError:
            pass
        return task.events

    events = asyncio.run(run_heartbeat())

    heartbeat = next(event for event in events if event.event == "heartbeat")
    assert isinstance(heartbeat.data["elapsed_seconds"], float | int)
    assert heartbeat.data["chunk_start"] == 0
    assert heartbeat.data["chunk_end"] == 3000


def test_proofread_task_fails_when_all_chunks_fail(monkeypatch):
    async def fake_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
    ):
        raise AIClientError("chunk failed")

    monkeypatch.setattr(task_service, "proofread_text", fake_proofread_text)

    create_response = client.post(
        "/api/proofread/tasks",
        json=proofread_payload(
            ("甲" * 3000) + ("乙" * 3000),
            scope="document",
            chunk_size=3000,
        ),
    )

    task_id = create_response.json()["task_id"]
    events = parse_sse_events(client.get(f"/api/proofread/tasks/{task_id}/events").text)
    status_response = client.get(f"/api/proofread/tasks/{task_id}")

    assert status_response.json()["status"] == "failed"
    assert status_response.json()["failed_chunks"] == 2
    assert events[-1]["event"] == "error"


def test_proofread_task_can_be_cancelled(monkeypatch):
    async def fake_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
    ):
        return []

    monkeypatch.setattr(task_service, "proofread_text", fake_proofread_text)

    create_response = client.post(
        "/api/proofread/tasks",
        json=proofread_payload(
            ("甲" * 3000) + ("乙" * 3000),
            scope="document",
            chunk_size=3000,
        ),
    )

    task_id = create_response.json()["task_id"]
    cancel_response = client.delete(f"/api/proofread/tasks/{task_id}")

    assert cancel_response.status_code == 200
    assert cancel_response.json()["task_id"] == task_id


def test_proofread_task_not_found_returns_404():
    response = client.get("/api/proofread/tasks/missing-task")

    assert response.status_code == 404
    assert response.json() == {"detail": "Proofread task not found"}
