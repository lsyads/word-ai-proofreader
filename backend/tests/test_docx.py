import io
import json
import zipfile
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import TypeAlias

from fastapi.testclient import TestClient
from app.main import app
from app.schemas import ProofreadIssue
from app.services import docx as docx_service
from app.services import docx_store
from app.services import docx_tasks as docx_task_service


client = TestClient(app)
BOOK = {"title": "测试书名", "introduction": "这是一部测试图书。"}
ParagraphSpec: TypeAlias = str | tuple[str, str]


def setup_function():
    docx_task_service.clear_tasks_for_tests()


def make_docx(
    paragraphs: list[ParagraphSpec],
    table_text: str = "",
    textbox_text: str = "",
    extra_document_namespaces: str = "",
    extra_document_attributes: str = "",
) -> bytes:
    body_parts = []
    for paragraph in paragraphs:
        if isinstance(paragraph, tuple):
            text, style = paragraph
            style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
        else:
            text = paragraph
            style_xml = ""

        body_parts.append(
            f"""
            <w:p>
              {style_xml}
              <w:r><w:t>{text}</w:t></w:r>
            </w:p>
            """
        )

    if table_text:
        body_parts.append(
            f"""
            <w:tbl><w:tr><w:tc>
              <w:p><w:r><w:t>{table_text}</w:t></w:r></w:p>
            </w:tc></w:tr></w:tbl>
            """
        )

    if textbox_text:
        body_parts.append(
            f"""
            <w:p>
              <w:r>
                <w:drawing>
                  <w:txbxContent>
                    <w:p><w:r><w:t>{textbox_text}</w:t></w:r></w:p>
                  </w:txbxContent>
                </w:drawing>
              </w:r>
            </w:p>
            """
        )

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      {extra_document_namespaces}
      {extra_document_attributes}>
      <w:body>
        {''.join(body_parts)}
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


def test_parse_docx_extracts_body_table_and_textbox_text():
    document = docx_service.parse_docx(
        make_docx(["目录", "第一章 开始", "正文内容"], table_text="表格文字", textbox_text="文本框文字")
    )

    assert "目录" in document.text
    assert "正文内容" in document.text
    assert "表格文字" in document.text
    assert "文本框文字" in document.text
    assert document.text.count("文本框文字") == 1


def test_split_docx_prefers_chapter_then_section_boundaries():
    long_section = "甲" * 7200
    document = docx_service.parse_docx(
        make_docx(["第一章 开始", "第一节 小节", long_section, "第二章 继续", "正文"])
    )

    chunks = docx_service.split_docx_into_chunks(document)

    assert len(chunks) >= 3
    assert chunks[0].text.startswith("第一章")
    assert any(chunk.text.startswith("第二章") for chunk in chunks)
    assert all(len(chunk.text) <= 7600 for chunk in chunks)


def test_split_docx_uses_toc_titles_only_after_chapter_and_section_are_still_large():
    long_body = "甲" * 3600
    document = docx_service.parse_docx(
        make_docx(
            [
                ("目录小标题一 1", "TOC1"),
                ("目录小标题二 2", "TOC1"),
                "第一章 开始",
                "第一节 小节",
                "目录小标题一",
                long_body,
                "目录小标题二",
                long_body,
            ]
        )
    )

    chunks = docx_service.split_docx_into_chunks(document)

    assert any(chunk.text.startswith("第一节 小节") for chunk in chunks)
    assert any(chunk.text.startswith("目录小标题一") for chunk in chunks)
    assert any(chunk.text.startswith("目录小标题二") for chunk in chunks)


def test_split_docx_does_not_use_heading3_as_toc_without_toc_entries():
    long_body = "甲" * 3600
    document = docx_service.parse_docx(
        make_docx(
            [
                "第一章 开始",
                "第一节 小节",
                ("普通三级标题一", "Heading3"),
                long_body,
                ("普通三级标题二", "Heading3"),
                long_body,
            ]
        )
    )

    chunks = docx_service.split_docx_into_chunks(document)

    assert not any(chunk.text.startswith("普通三级标题一") for chunk in chunks)
    assert not any(chunk.text.startswith("普通三级标题二") for chunk in chunks)


def test_write_docx_result_inserts_comment(tmp_path: Path):
    source = make_docx(["第一章 开始", "这里有错字。"])
    issue = ProofreadIssue(
        id="issue-1",
        category="typo",
        severity="high",
        original="错字",
        suggestion="修正错字。",
        start=3,
        end=5,
    )
    chunked_issue = docx_service.chunking.globalize_issues(
        docx_service.ProofreadChunk(index=0, start=0, end=20, text="这里有错字。"),
        [issue],
    )[0]
    chunked_issue.global_start = docx_service.parse_docx(source).text.index("错字")
    chunked_issue.global_end = chunked_issue.global_start + 2
    output = tmp_path / "out.docx"

    summary = docx_service.write_docx_result(source, [chunked_issue], "comment", output)

    assert summary.comment_count == 1
    with zipfile.ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml").decode()
        comments_xml = archive.read("word/comments.xml").decode()
    assert "commentRangeStart" in document_xml
    assert "修正错字" in comments_xml


def test_write_docx_result_inserts_revision(tmp_path: Path):
    source = make_docx(["这里有错字。"])
    document = docx_service.parse_docx(source)
    start = document.text.index("错字")
    issue = ProofreadIssue(
        id="issue-1",
        category="typo",
        severity="high",
        original="错字",
        replacement="正字",
        suggestion="修正错字。",
    )
    chunked_issue = docx_service.ChunkedProofreadIssue(
        **issue.model_dump(),
        chunk_index=0,
        global_start=start,
        global_end=start + 2,
    )
    output = tmp_path / "out.docx"

    summary = docx_service.write_docx_result(source, [chunked_issue], "revision", output)

    assert summary.revision_count == 1
    with zipfile.ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml").decode()
    assert "<w:del" in document_xml
    assert "<w:ins" in document_xml
    assert "正字" in document_xml


def test_write_docx_result_preserves_ignorable_namespace_declarations(tmp_path: Path):
    source = make_docx(
        ["这里有错字。"],
        extra_document_namespaces='xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"',
        extra_document_attributes='mc:Ignorable="w14 wp14"',
    )
    document = docx_service.parse_docx(source)
    start = document.text.index("错字")
    issue = ProofreadIssue(
        id="issue-1",
        category="typo",
        severity="high",
        original="错字",
        suggestion="修正错字。",
    )
    chunked_issue = docx_service.ChunkedProofreadIssue(
        **issue.model_dump(),
        chunk_index=0,
        global_start=start,
        global_end=start + 2,
    )
    output = tmp_path / "out.docx"

    docx_service.write_docx_result(source, [chunked_issue], "comment", output)

    with zipfile.ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml").decode()
    assert 'mc:Ignorable="w14 wp14"' in document_xml
    assert 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"' in document_xml
    assert 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"' in document_xml


def test_build_output_filename_uses_utc_plus_8_timestamp(monkeypatch):
    class FixedDatetime:
        @classmethod
        def now(cls, tz=None):
            assert tz is docx_service.OUTPUT_FILENAME_TIMEZONE
            return datetime(2026, 5, 2, 9, 2, 3, tzinfo=tz)

    monkeypatch.setattr(docx_service, "datetime", FixedDatetime)

    assert docx_service.build_output_filename("书稿.docx", "comment") == "书稿-AI审校-批注-20260502090203.docx"


def test_docx_task_api_uploads_generates_and_downloads(monkeypatch, tmp_path: Path):
    async def fake_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
        reasoning_enabled=False,
    ):
        assert "错字" in text
        return [
            ProofreadIssue(
                id="issue-1",
                category="typo",
                severity="high",
                original="错字",
                replacement="正字",
                suggestion="修正错字。",
            )
        ]

    monkeypatch.setattr(docx_task_service, "proofread_text", fake_proofread_text)
    monkeypatch.setenv("DOCX_OUTPUT_DIR", str(tmp_path / "docx-results"))
    source = make_docx(["这里有错字。"])
    response = client.post(
        "/api/proofread/docx/tasks",
        params={
            "filename": "书稿.docx",
            "book": json.dumps(BOOK, ensure_ascii=False),
            "application_mode": "revision",
        },
        content=source,
        headers={"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
    )

    assert response.status_code == 200
    task_id = response.json()["task_id"]

    with client.stream("GET", f"/api/proofread/docx/tasks/{task_id}/events") as stream:
        body = stream.read().decode()
    assert "completed" in body

    final = client.get(f"/api/proofread/docx/tasks/{task_id}")
    assert final.status_code == 200
    payload = final.json()
    assert payload["status"] == "succeeded"
    assert payload["issue_count"] == 1
    assert payload["output_filename"].endswith(".docx")
    assert payload["retention_days"] == 7
    assert datetime.fromisoformat(payload["expires_at"]) - datetime.now(UTC) > timedelta(days=6)

    download = client.get(f"/api/proofread/docx/tasks/{task_id}/download")
    assert download.status_code == 200
    with zipfile.ZipFile(io.BytesIO(download.content)) as archive:
        assert "word/document.xml" in archive.namelist()


def test_docx_download_survives_in_memory_task_restart(monkeypatch, tmp_path: Path):
    async def fake_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
        reasoning_enabled=False,
    ):
        return []

    monkeypatch.setattr(docx_task_service, "proofread_text", fake_proofread_text)
    monkeypatch.setenv("DOCX_OUTPUT_DIR", str(tmp_path / "docx-results"))
    source = make_docx(["第一章 开始", "没有问题。"])
    response = client.post(
        "/api/proofread/docx/tasks",
        params={
            "filename": "书稿.docx",
            "book": json.dumps(BOOK, ensure_ascii=False),
            "application_mode": "comment",
        },
        content=source,
        headers={"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
    )
    assert response.status_code == 200
    task_id = response.json()["task_id"]

    with client.stream("GET", f"/api/proofread/docx/tasks/{task_id}/events") as stream:
        assert "completed" in stream.read().decode()

    docx_task_service.clear_tasks_for_tests()

    final = client.get(f"/api/proofread/docx/tasks/{task_id}")
    assert final.status_code == 200
    assert final.json()["output_filename"].endswith(".docx")

    download = client.get(f"/api/proofread/docx/tasks/{task_id}/download")
    assert download.status_code == 200
    with zipfile.ZipFile(io.BytesIO(download.content)) as archive:
        assert "word/document.xml" in archive.namelist()


def test_docx_store_cleanup_deletes_only_expired_results(monkeypatch, tmp_path: Path):
    monkeypatch.setenv("DOCX_OUTPUT_DIR", str(tmp_path / "docx-results"))
    root = docx_store.output_dir()
    expired_file = root / "expired-task" / "expired.docx"
    active_file = root / "active-task" / "active.docx"
    expired_file.parent.mkdir(parents=True)
    active_file.parent.mkdir(parents=True)
    expired_file.write_bytes(b"expired")
    active_file.write_bytes(b"active")
    now = datetime.now(UTC)

    docx_store.save_result(
        task_id="expired-task",
        source_filename="源.docx",
        output_filename="expired.docx",
        application_mode="comment",
        status="succeeded",
        total_chunks=1,
        completed_chunks=1,
        failed_chunks=0,
        issue_count=0,
        relative_path="expired-task/expired.docx",
        created_at=(now - timedelta(days=8)).isoformat(),
        updated_at=(now - timedelta(days=8)).isoformat(),
        expires_at=(now - timedelta(seconds=1)).isoformat(),
    )
    docx_store.save_result(
        task_id="active-task",
        source_filename="源.docx",
        output_filename="active.docx",
        application_mode="revision",
        status="partial_succeeded",
        total_chunks=2,
        completed_chunks=1,
        failed_chunks=1,
        issue_count=3,
        relative_path="active-task/active.docx",
        created_at=now.isoformat(),
        updated_at=now.isoformat(),
        expires_at=(now + timedelta(days=7)).isoformat(),
    )

    assert docx_store.cleanup_expired() == 1
    assert not expired_file.exists()
    assert active_file.exists()
    assert docx_store.get_result("expired-task") is None
    assert docx_store.get_result("active-task") is not None


def test_docx_download_reports_expired_or_missing_result(monkeypatch, tmp_path: Path):
    monkeypatch.setenv("DOCX_OUTPUT_DIR", str(tmp_path / "docx-results"))
    now = datetime.now(UTC)
    expired_file = docx_store.output_dir() / "expired-task" / "expired.docx"
    expired_file.parent.mkdir(parents=True)
    expired_file.write_bytes(b"expired")
    docx_store.save_result(
        task_id="expired-task",
        source_filename="源.docx",
        output_filename="expired.docx",
        application_mode="comment",
        status="succeeded",
        total_chunks=1,
        completed_chunks=1,
        failed_chunks=0,
        issue_count=0,
        relative_path="expired-task/expired.docx",
        created_at=(now - timedelta(days=8)).isoformat(),
        updated_at=(now - timedelta(days=8)).isoformat(),
        expires_at=(now - timedelta(seconds=1)).isoformat(),
    )
    docx_store.save_result(
        task_id="missing-task",
        source_filename="源.docx",
        output_filename="missing.docx",
        application_mode="comment",
        status="succeeded",
        total_chunks=1,
        completed_chunks=1,
        failed_chunks=0,
        issue_count=0,
        relative_path="missing-task/missing.docx",
        created_at=now.isoformat(),
        updated_at=now.isoformat(),
        expires_at=(now + timedelta(days=7)).isoformat(),
    )

    expired = client.get("/api/proofread/docx/tasks/expired-task/download")
    missing = client.get("/api/proofread/docx/tasks/missing-task/download")

    assert expired.status_code == 409
    assert "expired" in expired.json()["detail"]
    assert missing.status_code == 409
    assert "missing" in missing.json()["detail"]


def test_docx_upload_rejects_doc_extension():
    response = client.post(
        "/api/proofread/docx/tasks",
        params={"filename": "旧稿.doc", "book": json.dumps(BOOK, ensure_ascii=False)},
        content=b"not-docx",
    )

    assert response.status_code == 400
    assert "另存为 .docx" in response.json()["detail"]
