from fastapi.testclient import TestClient

from app.main import app


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
