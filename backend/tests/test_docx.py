import io
import json
import zipfile
from pathlib import Path

from fastapi.testclient import TestClient
from app.main import app
from app.schemas import ProofreadIssue
from app.services import docx as docx_service
from app.services import docx_tasks as docx_task_service


client = TestClient(app)
BOOK = {"title": "测试书名", "introduction": "这是一部测试图书。"}


def setup_function():
    docx_task_service.clear_tasks_for_tests()


def make_docx(paragraphs: list[str], table_text: str = "", textbox_text: str = "") -> bytes:
    body_parts = []
    for text in paragraphs:
        body_parts.append(
            f"""
            <w:p>
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
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
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


def test_docx_task_api_uploads_generates_and_downloads(monkeypatch):
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

    download = client.get(f"/api/proofread/docx/tasks/{task_id}/download")
    assert download.status_code == 200
    with zipfile.ZipFile(io.BytesIO(download.content)) as archive:
        assert "word/document.xml" in archive.namelist()


def test_docx_upload_rejects_doc_extension():
    response = client.post(
        "/api/proofread/docx/tasks",
        params={"filename": "旧稿.doc", "book": json.dumps(BOOK, ensure_ascii=False)},
        content=b"not-docx",
    )

    assert response.status_code == 400
    assert "另存为 .docx" in response.json()["detail"]
