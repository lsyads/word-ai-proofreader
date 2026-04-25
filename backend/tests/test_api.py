import json

from fastapi.testclient import TestClient

import app.services.proofread as proofread_service
from app.main import app
from app.schemas import ProofreadIssue
from app.services.ai_client import AIClientError, AIProofreadResult, AIStreamEvent
from app.services.sessions import clear_sessions_for_tests


client = TestClient(app)


def setup_function():
    clear_sessions_for_tests()


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
    response = client.post("/api/proofread", json={"text": "   "})

    assert response.status_code == 422


def test_proofread_returns_mock_issue_without_api_key(monkeypatch):
    monkeypatch.delenv("AI_API_KEY", raising=False)

    response = client.post(
        "/api/proofread",
        json={"text": "这是一段需要审校的文本。", "context": {"source": "word-addin"}},
    )

    assert response.status_code == 200
    payload = response.json()
    assert len(payload["issues"]) == 1

    issue = payload["issues"][0]
    assert issue["id"] == "mock-issue-1"
    assert issue["category"] == "style"
    assert issue["severity"] == "medium"
    assert issue["original"]
    assert issue["suggestion"]
    assert issue["comment"]
    assert issue["start"] == 0
    assert isinstance(issue["end"], int)


def test_proofread_uses_ai_client_when_api_key_is_configured(monkeypatch):
    async def fake_proofread_with_ai(text, previous_response_id=None):
        assert text == "这是一段需要真实审校的文本。"
        assert previous_response_id is None
        return AIProofreadResult(
            response_id="resp-1",
            issues=[
                ProofreadIssue(
                    id="ai-issue-1",
                    category="typo",
                    severity="high",
                    original="真实",
                    suggestion="真实建议",
                    comment="真实 AI 分支",
                    start=0,
                    end=2,
                )
            ],
        )

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    response = client.post("/api/proofread", json={"text": "这是一段需要真实审校的文本。"})

    assert response.status_code == 200
    assert response.json()["issues"][0]["id"] == "ai-issue-1"


def test_proofread_sends_previous_response_id_for_same_session(monkeypatch):
    previous_response_ids = []

    async def fake_proofread_with_ai(text, previous_response_id=None):
        previous_response_ids.append(previous_response_id)
        return AIProofreadResult(response_id=f"resp-{len(previous_response_ids)}", issues=[])

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    session_id = client.post("/api/sessions").json()["session_id"]

    first = client.post("/api/proofread", json={"text": "第一段文本。", "session_id": session_id})
    second = client.post("/api/proofread", json={"text": "第二段文本。", "session_id": session_id})

    assert first.status_code == 200
    assert second.status_code == 200
    assert previous_response_ids == [None, "resp-1"]


def test_proofread_converts_ai_client_error_to_502(monkeypatch):
    async def fake_proofread_with_ai(text, previous_response_id=None):
        raise AIClientError("AI provider returned HTTP 500")

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    response = client.post("/api/proofread", json={"text": "这是一段文本。"})

    assert response.status_code == 502
    assert response.json() == {"detail": "AI provider returned HTTP 500"}


def test_proofread_stream_returns_status_events_and_result(monkeypatch):
    async def fake_stream_proofread_text(text, session_id=None):
        assert text == "这是一段需要真实审校的文本。"
        assert session_id == "session-test"
        yield AIStreamEvent("status", {"stage": "received", "message": "已接收选区文本。"})
        yield AIStreamEvent("status", {"stage": "calling_ai", "message": "正在调用 AI 原生 Responses session。"})
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
                        "comment": "真实 AI 分支",
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
        json={"text": "这是一段需要真实审校的文本。", "session_id": "session-test"},
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


def test_proofread_stream_returns_error_event_for_ai_client_error(monkeypatch):
    async def fake_stream_proofread_text(text, session_id=None):
        yield AIStreamEvent("status", {"stage": "received", "message": "已接收选区文本。"})
        yield AIStreamEvent("status", {"stage": "calling_ai", "message": "正在调用 AI 原生 Responses session。"})
        raise AIClientError("AI provider returned HTTP 500")

    monkeypatch.setattr(proofread_service, "stream_proofread_text", fake_stream_proofread_text)

    response = client.post("/api/proofread/stream", json={"text": "这是一段文本。"})

    assert response.status_code == 200

    events = parse_sse_events(response.text)
    assert [event["event"] for event in events] == ["status", "status", "error"]
    assert events[-1]["data"] == {"message": "AI provider returned HTTP 500"}
