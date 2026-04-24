from fastapi.testclient import TestClient

import app.services.proofread as proofread_service
from app.main import app
from app.schemas import ProofreadIssue
from app.services.ai_client import AIClientError


client = TestClient(app)


def test_health_returns_ok():
    response = client.get("/health")

    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


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
    async def fake_proofread_with_ai(text):
        assert text == "这是一段需要真实审校的文本。"
        return [
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
        ]

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    response = client.post("/api/proofread", json={"text": "这是一段需要真实审校的文本。"})

    assert response.status_code == 200
    assert response.json()["issues"][0]["id"] == "ai-issue-1"


def test_proofread_converts_ai_client_error_to_502(monkeypatch):
    async def fake_proofread_with_ai(text):
        raise AIClientError("AI provider returned HTTP 500")

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(proofread_service, "proofread_with_ai", fake_proofread_with_ai)

    response = client.post("/api/proofread", json={"text": "这是一段文本。"})

    assert response.status_code == 502
    assert response.json() == {"detail": "AI provider returned HTTP 500"}
