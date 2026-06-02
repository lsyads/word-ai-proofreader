import io
import json
import zipfile

from fastapi.testclient import TestClient

from app.agents import workspace as workspace_agent
from app.main import app
from app.services import project_store


client = TestClient(app)
BOOK = {"title": "测试书名", "introduction": "这是一部测试图书。"}


def setup_function():
    project_store.clear_store_for_tests()


def make_docx(paragraphs: list[str]) -> bytes:
    body = "".join(
        f"""
        <w:p>
          <w:r><w:t>{paragraph}</w:t></w:r>
        </w:p>
        """
        for paragraph in paragraphs
    )
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:body>
        {body}
        <w:sectPr/>
      </w:body>
    </w:document>
    """
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>
    """
    package_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>
    """
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
    """
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", package_rels)
        archive.writestr("word/document.xml", document_xml)
        archive.writestr("word/_rels/document.xml.rels", document_rels)
    return buffer.getvalue()


def get_completed_run(project_id: str, run_id: str) -> dict:
    run_response = client.get(f"/api/v2/projects/{project_id}/runs/{run_id}")
    assert run_response.status_code == 200
    run = run_response.json()
    assert run["status"] == "waiting_for_approval"
    assert run["candidate_count"] >= 1
    return run


def test_v2_project_run_approval_writeback_report_and_download():
    source = make_docx(["第一章 开始", "这里有错字，需要审校。"])
    create_response = client.post(
        "/api/v2/projects",
        params={
            "filename": "书稿.docx",
            "book": json.dumps(BOOK, ensure_ascii=False),
            "review_goal": "检查明显出版审校问题。",
        },
        content=source,
    )

    assert create_response.status_code == 200
    project = create_response.json()
    project_id = project["project_id"]
    assert project["status"] == "created"
    assert project["source_type"] == "docx"

    map_response = client.get(f"/api/v2/projects/{project_id}/document-map")
    assert map_response.status_code == 200
    document_map = map_response.json()
    assert document_map["block_count"] == 2
    assert document_map["chunk_count"] >= 1

    run_response = client.post(f"/api/v2/projects/{project_id}/runs", json={})
    assert run_response.status_code == 200
    run = run_response.json()
    assert run["status"] == "queued"
    run = get_completed_run(project_id, run["run_id"])
    assert run["status"] == "waiting_for_approval"
    assert run["candidate_count"] >= 1

    projects_response = client.get("/api/v2/projects")
    assert projects_response.status_code == 200
    assert projects_response.json()["projects"][0]["project_id"] == project_id

    plan_response = client.get(f"/api/v2/projects/{project_id}/plan")
    assert plan_response.status_code == 200
    plan_steps = plan_response.json()["steps"]
    assert any(step["step_id"] == "proofread_pass" for step in plan_steps)
    assert any(step["step_id"] == "style_rule_pass" for step in plan_steps)

    trace_response = client.get(f"/api/v2/projects/{project_id}/runs/{run['run_id']}/trace")
    assert trace_response.status_code == 200
    assert "第一章 开始" not in trace_response.text
    assert any(event["event"] == "candidate_found" for event in trace_response.json()["events"])
    assert any(event["event"] == "pass_started" for event in trace_response.json()["events"])
    assert any(event["event"] == "candidate_evaluated" for event in trace_response.json()["events"])

    candidates_response = client.get(f"/api/v2/projects/{project_id}/candidates")
    assert candidates_response.status_code == 200
    candidate = candidates_response.json()["candidates"][0]
    assert candidate["pass_name"]
    assert candidate["confidence"] > 0
    assert candidate["evaluation_note"]

    decision_response = client.post(
        f"/api/v2/projects/{project_id}/candidates/decisions",
        json={"decisions": [{"candidate_id": candidate["candidate_id"], "status": "approved"}]},
    )
    assert decision_response.status_code == 200
    assert decision_response.json()["updated_count"] == 1

    memory_response = client.get(f"/api/v2/projects/{project_id}/memory")
    assert memory_response.status_code == 200
    assert any(item["key"] == "approved_issue_categories" for item in memory_response.json()["memory"])

    writeback_response = client.post(f"/api/v2/projects/{project_id}/writeback", json={"application_mode": "comment"})
    assert writeback_response.status_code == 200
    writeback = writeback_response.json()
    assert writeback["written_count"] == 1
    assert writeback["download_url"] == f"/api/v2/projects/{project_id}/download"

    report_response = client.get(f"/api/v2/projects/{project_id}/report")
    assert report_response.status_code == 200
    report = report_response.json()
    assert report["written_count"] == 1
    assert report["issue_count"] >= 1
    assert report["pass_counts"]

    download_response = client.get(f"/api/v2/projects/{project_id}/download")
    assert download_response.status_code == 200
    assert download_response.content.startswith(b"PK")


def test_v2_selection_project_run_writeback_conflict_and_mark_written():
    create_response = client.post(
        "/api/v2/projects/selection",
        json={
            "text": "这里有错字，需要审校。",
            "book": BOOK,
            "review_goal": "检查当前选区。",
            "session_id": "session-test",
        },
    )

    assert create_response.status_code == 200
    project = create_response.json()
    project_id = project["project_id"]
    assert project["source_type"] == "selection"
    assert project["text_preview"] == "这里有错字，需要审校。"

    map_response = client.get(f"/api/v2/projects/{project_id}/document-map")
    assert map_response.status_code == 200
    assert map_response.json()["text_len"] == len("这里有错字，需要审校。")

    run_response = client.post(f"/api/v2/projects/{project_id}/runs", json={})
    assert run_response.status_code == 200
    run = run_response.json()
    assert run["status"] == "queued"
    run = get_completed_run(project_id, run["run_id"])
    assert run["status"] == "waiting_for_approval"
    assert run["candidate_count"] >= 1

    trace_response = client.get(f"/api/v2/projects/{project_id}/runs/{run['run_id']}/trace")
    assert trace_response.status_code == 200
    assert "这里有错字" not in trace_response.text

    candidate = client.get(f"/api/v2/projects/{project_id}/candidates").json()["candidates"][0]
    decision_response = client.post(
        f"/api/v2/projects/{project_id}/candidates/decisions",
        json={"decisions": [{"candidate_id": candidate["candidate_id"], "status": "approved"}]},
    )
    assert decision_response.status_code == 200

    writeback_response = client.post(f"/api/v2/projects/{project_id}/writeback", json={"application_mode": "comment"})
    assert writeback_response.status_code == 409

    mark_response = client.post(
        f"/api/v2/projects/{project_id}/candidates/mark-written",
        json={"candidate_ids": [candidate["candidate_id"]]},
    )
    assert mark_response.status_code == 200
    assert mark_response.json()["updated_count"] == 1
    assert mark_response.json()["candidates"][0]["status"] == "written"

    report = client.get(f"/api/v2/projects/{project_id}/report").json()
    assert report["written_count"] == 1


def test_v2_selection_project_with_no_candidates_completes_successfully(monkeypatch):
    async def fake_proofread_text_with_context(*args, **kwargs):
        return []

    monkeypatch.setattr(
        workspace_agent.proofread_service,
        "proofread_text_with_context",
        fake_proofread_text_with_context,
    )
    create_response = client.post(
        "/api/v2/projects/selection",
        json={
            "text": "这是一段正常文字。",
            "book": BOOK,
            "review_goal": "检查当前选区。",
        },
    )
    assert create_response.status_code == 200
    project_id = create_response.json()["project_id"]

    run_response = client.post(f"/api/v2/projects/{project_id}/runs", json={})
    assert run_response.status_code == 200
    run = client.get(f"/api/v2/projects/{project_id}/runs/{run_response.json()['run_id']}").json()

    assert run["status"] == "succeeded"
    assert run["candidate_count"] == 0
    assert client.get(f"/api/v2/projects/{project_id}").json()["status"] == "succeeded"
    assert client.get(f"/api/v2/projects/{project_id}/candidates").json()["candidates"] == []
    report = client.get(f"/api/v2/projects/{project_id}/report").json()
    assert report["status"] == "succeeded"
    assert report["issue_count"] == 0
    trace = client.get(f"/api/v2/projects/{project_id}/runs/{run['run_id']}/trace").json()
    assert any(event["event"] == "review_completed" for event in trace["events"])
    assert not any(event["event"] == "waiting_for_approval" for event in trace["events"])


def test_v2_delete_selection_project_removes_related_workspace_data():
    create_response = client.post(
        "/api/v2/projects/selection",
        json={
            "text": "这里有错字，需要审校。",
            "book": BOOK,
            "review_goal": "检查当前选区。",
        },
    )
    assert create_response.status_code == 200
    project_id = create_response.json()["project_id"]

    run = client.post(f"/api/v2/projects/{project_id}/runs", json={}).json()
    run = get_completed_run(project_id, run["run_id"])
    candidate = client.get(f"/api/v2/projects/{project_id}/candidates").json()["candidates"][0]
    client.post(
        f"/api/v2/projects/{project_id}/candidates/decisions",
        json={"decisions": [{"candidate_id": candidate["candidate_id"], "status": "approved"}]},
    )
    memory_response = client.get(f"/api/v2/projects/{project_id}/memory")
    assert memory_response.status_code == 200
    assert memory_response.json()["memory"]

    delete_response = client.delete(f"/api/v2/projects/{project_id}")
    assert delete_response.status_code == 200
    assert delete_response.json() == {"project_id": project_id, "deleted": True}

    assert client.get(f"/api/v2/projects/{project_id}").status_code == 404
    assert client.get(f"/api/v2/projects/{project_id}/document-map").status_code == 404
    assert client.get(f"/api/v2/projects/{project_id}/plan").status_code == 404
    assert client.get(f"/api/v2/projects/{project_id}/runs/{run['run_id']}").status_code == 404
    assert client.get(f"/api/v2/projects/{project_id}/candidates").status_code == 404
    assert client.get(f"/api/v2/projects/{project_id}/report").status_code == 404
    assert client.get(f"/api/v2/projects/{project_id}/memory").status_code == 404
    assert project_id not in [project["project_id"] for project in client.get("/api/v2/projects").json()["projects"]]


def test_v2_delete_docx_project_removes_output_file_and_download():
    source = make_docx(["第一章 开始", "这里有错字，需要审校。"])
    create_response = client.post(
        "/api/v2/projects",
        params={
            "filename": "书稿.docx",
            "book": json.dumps(BOOK, ensure_ascii=False),
            "review_goal": "检查明显出版审校问题。",
        },
        content=source,
    )
    assert create_response.status_code == 200
    project_id = create_response.json()["project_id"]
    run = client.post(f"/api/v2/projects/{project_id}/runs", json={}).json()
    get_completed_run(project_id, run["run_id"])
    candidate = client.get(f"/api/v2/projects/{project_id}/candidates").json()["candidates"][0]
    client.post(
        f"/api/v2/projects/{project_id}/candidates/decisions",
        json={"decisions": [{"candidate_id": candidate["candidate_id"], "status": "approved"}]},
    )
    writeback_response = client.post(f"/api/v2/projects/{project_id}/writeback", json={"application_mode": "comment"})
    assert writeback_response.status_code == 200
    output_path, _ = project_store.output_download_path(project_id)
    assert output_path.exists()

    delete_response = client.delete(f"/api/v2/projects/{project_id}")
    assert delete_response.status_code == 200
    assert not output_path.exists()
    assert client.get(f"/api/v2/projects/{project_id}/download").status_code == 404


def test_v2_delete_missing_project_returns_404():
    response = client.delete("/api/v2/projects/project_missing")
    assert response.status_code == 404
