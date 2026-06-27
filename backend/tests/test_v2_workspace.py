import io
import json
import zipfile

from fastapi.testclient import TestClient

from app.agents import workspace as workspace_agent
from app.agents import local_rules
from app.main import app
from app.schemas import ProofreadIssue, V2CandidateIssue, V2RunCreateRequest
from app.services.ai_client import AIClientError, AIProofreadResult
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


def make_candidate(
    project_id: str,
    index: int,
    *,
    run_id: str = "v2_run_test",
    status: str = "pending",
    pass_name: str = "proofread_pass",
) -> V2CandidateIssue:
    now = f"2026-06-02T00:00:{index:02d}+00:00"
    return V2CandidateIssue(
        candidate_id=f"candidate_test_{index:02d}",
        project_id=project_id,
        run_id=run_id,
        status=status,
        category="typo",
        severity="medium",
        original=f"原文{index}",
        replacement=f"替换{index}",
        suggestion=f"建议{index}",
        evidence=f"证据{index}",
        chunk_index=0,
        global_start=index,
        global_end=index + 2,
        pass_name=pass_name,
        confidence=0.8,
        evidence_kind="locator",
        needs_human_review=True,
        created_at=now,
        updated_at=now,
    )


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
    assert [step["step_id"] for step in plan_steps] == [
        "plan_review",
        "proofread_pass",
        "merge_candidates",
        "evaluate_candidates",
        "human_approval",
    ]

    trace_response = client.get(f"/api/v2/projects/{project_id}/runs/{run['run_id']}/trace")
    assert trace_response.status_code == 200
    assert "第一章 开始" not in trace_response.text
    assert "review_goal_bound" not in trace_response.text
    assert "审校目标已进入 Agent 上下文" not in trace_response.text
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


def test_v2_candidates_endpoint_paginates_and_filters():
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
    run = project_store.create_run(project_id, total_chunks=1)
    candidates = [
        make_candidate(project_id, index, run_id=run.run_id, status="pending", pass_name="proofread_pass")
        for index in range(1, 26)
    ]
    candidates.extend(
        [
            make_candidate(project_id, 26, run_id=run.run_id, status="approved", pass_name="custom_pass"),
            make_candidate(project_id, 27, run_id=run.run_id, status="pending", pass_name="custom_pass"),
        ]
    )
    project_store.save_candidates(candidates)

    first_page = client.get(f"/api/v2/projects/{project_id}/candidates", params={"page_size": 10})
    assert first_page.status_code == 200
    payload = first_page.json()
    assert payload["run_id"] == run.run_id
    assert payload["page"] == 1
    assert payload["page_size"] == 10
    assert payload["total"] == 27
    assert payload["total_pages"] == 3
    assert payload["has_previous"] is False
    assert payload["has_next"] is True
    assert [item["candidate_id"] for item in payload["candidates"][:2]] == [
        "candidate_test_01",
        "candidate_test_02",
    ]

    second_page = client.get(
        f"/api/v2/projects/{project_id}/candidates",
        params={"page": 2, "page_size": 10},
    ).json()
    assert second_page["has_previous"] is True
    assert second_page["candidates"][0]["candidate_id"] == "candidate_test_11"

    filtered = client.get(
        f"/api/v2/projects/{project_id}/candidates",
        params={"status": "pending", "pass_name": "custom_pass"},
    )
    assert filtered.status_code == 200
    filtered_payload = filtered.json()
    assert filtered_payload["total"] == 1
    assert filtered_payload["candidates"][0]["candidate_id"] == "candidate_test_27"


def test_v2_bulk_decision_updates_all_pending_candidates():
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
    run = project_store.create_run(project_id, total_chunks=1)
    project_store.save_candidates(
        [
            make_candidate(project_id, 1, run_id=run.run_id, status="pending"),
            make_candidate(project_id, 2, run_id=run.run_id, status="pending"),
            make_candidate(project_id, 3, run_id=run.run_id, status="approved"),
        ]
    )

    response = client.post(
        f"/api/v2/projects/{project_id}/candidates/bulk-decisions",
        json={"status": "rejected"},
    )
    assert response.status_code == 200
    assert response.json()["updated_count"] == 2
    statuses = {
        item["candidate_id"]: item["status"]
        for item in client.get(f"/api/v2/projects/{project_id}/candidates", params={"page_size": 10}).json()[
            "candidates"
        ]
    }
    assert statuses == {
        "candidate_test_01": "rejected",
        "candidate_test_02": "rejected",
        "candidate_test_03": "approved",
    }


def test_v2_defaults_scope_candidates_summary_and_report_to_latest_run():
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

    old_run = project_store.create_run(project_id, total_chunks=1)
    project_store.save_candidates(
        [
            make_candidate(project_id, 31, run_id=old_run.run_id, status="pending"),
            make_candidate(project_id, 32, run_id=old_run.run_id, status="approved"),
        ]
    )
    new_run = project_store.create_run(project_id, total_chunks=1)
    project_store.save_candidates(
        [
            make_candidate(project_id, 41, run_id=new_run.run_id, status="pending"),
            make_candidate(project_id, 42, run_id=new_run.run_id, status="approved"),
            make_candidate(project_id, 43, run_id=new_run.run_id, status="rejected"),
        ]
    )

    project_payload = client.get(f"/api/v2/projects/{project_id}").json()
    assert project_payload["latest_run_id"] == new_run.run_id
    assert project_payload["candidate_count"] == 3
    assert project_payload["pending_count"] == 1
    assert project_payload["approved_count"] == 1

    default_candidates = client.get(f"/api/v2/projects/{project_id}/candidates", params={"page_size": 10})
    assert default_candidates.status_code == 200
    default_payload = default_candidates.json()
    assert default_payload["run_id"] == new_run.run_id
    assert default_payload["total"] == 3
    assert {item["candidate_id"] for item in default_payload["candidates"]} == {
        "candidate_test_41",
        "candidate_test_42",
        "candidate_test_43",
    }

    old_candidates = client.get(
        f"/api/v2/projects/{project_id}/candidates",
        params={"run_id": old_run.run_id, "page_size": 10},
    )
    assert old_candidates.status_code == 200
    old_payload = old_candidates.json()
    assert old_payload["run_id"] == old_run.run_id
    assert old_payload["total"] == 2
    assert {item["candidate_id"] for item in old_payload["candidates"]} == {
        "candidate_test_31",
        "candidate_test_32",
    }

    report = client.get(f"/api/v2/projects/{project_id}/report").json()
    assert report["issue_count"] == 3
    assert report["pending_count"] == 1
    assert report["approved_count"] == 1
    assert report["rejected_count"] == 1


def test_v2_rejects_new_run_when_project_has_active_run():
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
    project_store.create_run(project_id, total_chunks=1)

    response = client.post(f"/api/v2/projects/{project_id}/runs", json={})

    assert response.status_code == 409


def test_v2_bulk_decision_only_updates_latest_run_pending_candidates():
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

    old_run = project_store.create_run(project_id, total_chunks=1)
    new_run = project_store.create_run(project_id, total_chunks=1)
    project_store.save_candidates(
        [
            make_candidate(project_id, 51, run_id=old_run.run_id, status="pending"),
            make_candidate(project_id, 52, run_id=new_run.run_id, status="pending"),
            make_candidate(project_id, 53, run_id=new_run.run_id, status="approved"),
        ]
    )

    response = client.post(
        f"/api/v2/projects/{project_id}/candidates/bulk-decisions",
        json={"status": "rejected"},
    )

    assert response.status_code == 200
    assert response.json()["updated_count"] == 1
    old_candidate = client.get(
        f"/api/v2/projects/{project_id}/candidates",
        params={"run_id": old_run.run_id},
    ).json()["candidates"][0]
    latest_statuses = {
        item["candidate_id"]: item["status"]
        for item in client.get(f"/api/v2/projects/{project_id}/candidates", params={"page_size": 10}).json()[
            "candidates"
        ]
    }
    assert old_candidate["status"] == "pending"
    assert latest_statuses == {
        "candidate_test_52": "rejected",
        "candidate_test_53": "approved",
    }


def test_v2_docx_writeback_only_marks_latest_run_approved_candidates():
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

    old_run = project_store.create_run(project_id, total_chunks=1)
    new_run = project_store.create_run(project_id, total_chunks=1)
    project_store.save_candidates(
        [
            make_candidate(project_id, 61, run_id=old_run.run_id, status="approved"),
            make_candidate(project_id, 62, run_id=new_run.run_id, status="approved"),
        ]
    )

    response = client.post(f"/api/v2/projects/{project_id}/writeback", json={"application_mode": "comment"})

    assert response.status_code == 200
    assert response.json()["written_count"] == 1
    old_status = client.get(
        f"/api/v2/projects/{project_id}/candidates",
        params={"run_id": old_run.run_id},
    ).json()["candidates"][0]["status"]
    latest_status = client.get(f"/api/v2/projects/{project_id}/candidates").json()["candidates"][0]["status"]
    assert old_status == "approved"
    assert latest_status == "written"


def test_v2_start_project_run_uses_document_map_chunk_count_without_docx_parse(monkeypatch):
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
    expected_chunk_count = client.get(f"/api/v2/projects/{project_id}/document-map").json()["chunk_count"]

    def fail_parse_docx(*args, **kwargs):
        raise AssertionError("start_project_run should not parse DOCX")

    monkeypatch.setattr(workspace_agent.docx_service, "parse_docx", fail_parse_docx)

    run = workspace_agent.workspace_runner.start_project_run(project_id, V2RunCreateRequest())

    assert run.status == "queued"
    assert run.total_chunks == expected_chunk_count


def test_local_non_ai_rules_do_not_emit_default_candidates():
    text = "AI 与人工智能并用。人工智能再次出现。这里有中文,逗号。还有中文,逗号！！中文(括号)。不要处理？！"
    repeated_numbers = "第1章统计10页。第2章仍为10页。第3章记录20页。"

    assert local_rules.run_terminology_rules(text) == []
    assert local_rules.run_style_rules(text) == []
    assert local_rules.run_consistency_rules(repeated_numbers, source_type="selection") == []
    assert local_rules.run_consistency_rules(repeated_numbers, source_type="docx") == []


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


def test_v2_run_exposes_chunk_progress_and_safe_trace_events(monkeypatch):
    async def fake_proofread_text_with_context(text, *args, **kwargs):
        if "乙" in text:
            raise AIClientError("chunk failed with Authorization: Bearer secret-token")
        return [
            ProofreadIssue(
                id="issue-1",
                category="typo",
                severity="medium",
                original="甲甲",
                replacement="甲乙",
                suggestion="修正明显错字。",
                start=0,
                end=2,
            )
        ]

    monkeypatch.setattr(
        workspace_agent.proofread_service,
        "proofread_text_with_context",
        fake_proofread_text_with_context,
    )
    text = ("甲" * 4999 + "。") + ("乙" * 4999 + "。")
    create_response = client.post(
        "/api/v2/projects/selection",
        json={
            "text": text,
            "book": BOOK,
            "review_goal": "检查长选区。",
        },
    )
    assert create_response.status_code == 200
    project_id = create_response.json()["project_id"]

    run_response = client.post(f"/api/v2/projects/{project_id}/runs", json={})
    assert run_response.status_code == 200
    run = client.get(f"/api/v2/projects/{project_id}/runs/{run_response.json()['run_id']}").json()

    assert run["status"] == "waiting_for_approval"
    assert run["total_chunks"] == 2
    assert run["completed_chunks"] == 1
    assert run["failed_chunks"] == 1
    assert run["candidate_count"] == 1

    trace_response = client.get(f"/api/v2/projects/{project_id}/runs/{run['run_id']}/trace")
    assert trace_response.status_code == 200
    trace = trace_response.json()
    tool_started = [event for event in trace["events"] if event["event"] == "tool_started"]
    tool_completed = [event for event in trace["events"] if event["event"] == "tool_completed"]
    errors = [event for event in trace["events"] if event["event"] == "error"]
    assert len(tool_started) == 2
    assert len(tool_completed) == 1
    assert errors
    assert tool_started[0]["data"]["total_chunks"] == 2
    assert tool_completed[0]["data"]["completed_chunks"] == 1
    assert tool_completed[0]["data"]["failed_chunks"] == 0
    assert tool_completed[0]["data"]["candidate_count"] == 1
    chunk_error = next(event for event in errors if event["data"].get("tool_name") == "proofread_document_chunk")
    assert chunk_error["data"]["completed_chunks"] == 1
    assert chunk_error["data"]["failed_chunks"] == 1
    assert chunk_error["data"]["message"] == "chunk failed with Authorization: Bearer [REDACTED]"
    serialized_trace = json.dumps(trace, ensure_ascii=False)
    assert "secret-token" not in serialized_trace
    assert "甲甲甲甲" not in serialized_trace
    assert "乙乙乙乙" not in serialized_trace


def test_v2_docx_repeated_numbers_do_not_create_local_candidates(monkeypatch):
    async def fake_proofread_text_with_context(*args, **kwargs):
        return []

    monkeypatch.setattr(
        workspace_agent.proofread_service,
        "proofread_text_with_context",
        fake_proofread_text_with_context,
    )
    source = make_docx(["第1章统计10页。第2章仍为10页。第3章记录20页。"])
    create_response = client.post(
        "/api/v2/projects",
        params={
            "filename": "数字重复.docx",
            "book": json.dumps(BOOK, ensure_ascii=False),
            "review_goal": "检查全书一致性。",
        },
        content=source,
    )
    assert create_response.status_code == 200
    project_id = create_response.json()["project_id"]

    run_response = client.post(f"/api/v2/projects/{project_id}/runs", json={})
    assert run_response.status_code == 200
    run = client.get(f"/api/v2/projects/{project_id}/runs/{run_response.json()['run_id']}").json()

    assert run["status"] == "succeeded"
    assert run["candidate_count"] == 0
    assert client.get(f"/api/v2/projects/{project_id}/candidates").json()["candidates"] == []


def test_v2_local_non_ai_rules_do_not_emit_candidates_when_ai_returns_empty(monkeypatch):
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
            "text": "这里有中文,逗号！！中文(括号)。",
            "book": BOOK,
            "review_goal": "检查当前选区体例。",
        },
    )
    assert create_response.status_code == 200
    project_id = create_response.json()["project_id"]

    run_response = client.post(f"/api/v2/projects/{project_id}/runs", json={})
    assert run_response.status_code == 200
    run_status_response = client.get(f"/api/v2/projects/{project_id}/runs/{run_response.json()['run_id']}")
    assert run_status_response.status_code == 200
    run = run_status_response.json()
    assert run["status"] == "succeeded"
    assert run["candidate_count"] == 0

    candidates = client.get(f"/api/v2/projects/{project_id}/candidates").json()["candidates"]
    assert candidates == []


def test_v2_filters_ai_mechanical_copyediting_candidates(monkeypatch):
    async def fake_proofread_with_ai(*args, **kwargs):
        return AIProofreadResult(
            response_id="resp-1",
            issues=[
                ProofreadIssue(
                    id="ai-issue-punctuation",
                    category="punctuation",
                    severity="low",
                    original="中文,逗号",
                    replacement="中文，逗号",
                    suggestion="替换符号。",
                ),
                ProofreadIssue(
                    id="ai-issue-space",
                    category="style",
                    severity="low",
                    original="A B",
                    replacement="AB",
                    suggestion="删除空格。",
                ),
            ],
        )

    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setattr(workspace_agent.proofread_service, "proofread_with_ai", fake_proofread_with_ai)
    create_response = client.post(
        "/api/v2/projects/selection",
        json={
            "text": "这里有中文,逗号，也有 A B。",
            "book": BOOK,
            "review_goal": "检查当前选区。",
        },
    )
    assert create_response.status_code == 200
    project_id = create_response.json()["project_id"]

    run_response = client.post(f"/api/v2/projects/{project_id}/runs", json={})
    assert run_response.status_code == 200
    run_status_response = client.get(f"/api/v2/projects/{project_id}/runs/{run_response.json()['run_id']}")
    assert run_status_response.status_code == 200
    run = run_status_response.json()
    assert run["status"] == "succeeded"
    assert run["candidate_count"] == 0
    assert client.get(f"/api/v2/projects/{project_id}/candidates").json()["candidates"] == []


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
