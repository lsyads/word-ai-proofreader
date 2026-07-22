import io
import json
import zipfile
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import TypeAlias
from xml.etree import ElementTree as ET

import pytest
from fastapi.testclient import TestClient
from app.agents import trace as agent_trace
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
    agent_trace.clear_traces_for_tests()


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


def replace_docx_entry(document_bytes: bytes, name: str, data: bytes) -> bytes:
    source = io.BytesIO(document_bytes)
    output = io.BytesIO()
    with zipfile.ZipFile(source, "r") as source_archive, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as output_archive:
        for entry_name in source_archive.namelist():
            output_archive.writestr(data=data if entry_name == name else source_archive.read(entry_name), zinfo_or_arcname=entry_name)
    return output.getvalue()


def make_docx_with_internal_entity() -> bytes:
    document_xml = b"""<?xml version="1.0" encoding="UTF-8"?>
    <!DOCTYPE document [
      <!ENTITY expanded "internal entity content">
    ]>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p><w:r><w:t>&expanded;</w:t></w:r></w:p>
        <w:sectPr/>
      </w:body>
    </w:document>
    """
    return replace_docx_entry(make_docx(["normal content"]), "word/document.xml", document_xml)


def test_parse_docx_extracts_body_table_and_textbox_text():
    document = docx_service.parse_docx(
        make_docx(["目录", "第一章 开始", "正文内容"], table_text="表格文字", textbox_text="文本框文字")
    )

    assert "目录" in document.text
    assert "正文内容" in document.text
    assert "表格文字" in document.text
    assert "文本框文字" in document.text
    assert document.text.count("文本框文字") == 1


def test_parse_docx_rejects_internal_xml_entities():
    with pytest.raises(docx_service.DocxError, match="无法安全解析"):
        docx_service.parse_docx(make_docx_with_internal_entity())


def test_docx_task_api_rejects_internal_xml_entities():
    response = client.post(
        "/api/proofread/docx/tasks",
        params={"filename": "malicious.docx", "book": json.dumps(BOOK, ensure_ascii=False)},
        content=make_docx_with_internal_entity(),
    )

    assert response.status_code == 400
    assert "无法安全解析" in response.json()["detail"]


def test_writeback_rejects_internal_entities_in_related_xml_parts(tmp_path: Path):
    relationships_xml = b"""<?xml version="1.0" encoding="UTF-8"?>
    <!DOCTYPE Relationships [
      <!ENTITY expanded "comments.xml">
    ]>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="unsafe" Target="&expanded;"/>
    </Relationships>
    """
    source = replace_docx_entry(
        make_docx(["normal content"]),
        "word/_rels/document.xml.rels",
        relationships_xml,
    )

    with pytest.raises(docx_service.DocxError, match="无法安全解析"):
        docx_service.write_docx_result(source, [], "comment", tmp_path / "out.docx")


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
    assert 'w:author="Word Proofreader"' in comments_xml
    assert "Word AI Proofreader" not in comments_xml
    assert "修正错字" in comments_xml


def test_write_docx_result_inserts_multiple_comments_in_same_text_node(tmp_path: Path):
    source = make_docx(["这里有甲字，也有乙字。"])
    document = docx_service.parse_docx(source)
    issues = []
    for issue_id, original in [("issue-1", "甲字"), ("issue-2", "乙字")]:
        start = document.text.index(original)
        issues.append(
            docx_service.ChunkedProofreadIssue(
                id=issue_id,
                category="typo",
                severity="high",
                original=original,
                suggestion=f"修正{original}。",
                chunk_index=0,
                global_start=start,
                global_end=start + len(original),
            )
        )
    output = tmp_path / "out.docx"

    summary = docx_service.write_docx_result(source, issues, "comment", output)

    assert summary.comment_count == 2
    assert summary.fallback_count == 0
    with zipfile.ZipFile(output) as archive:
        comments_xml = archive.read("word/comments.xml").decode()
    assert comments_xml.count("<w:comment ") == 2
    assert "修正甲字" in comments_xml
    assert "修正乙字" in comments_xml


def test_write_docx_result_relocates_issue_inside_bound_chunk(tmp_path: Path):
    source = make_docx(["第一章 开始", "这里有甲字。", "第二章 继续", "这里有乙字。"])
    document = docx_service.parse_docx(source)
    chunks = docx_service.split_docx_into_chunks(document)
    target_chunk = next(chunk for chunk in chunks if "乙字" in chunk.text)
    issue = docx_service.ChunkedProofreadIssue(
        id="issue-1",
        category="typo",
        severity="high",
        original="乙字",
        suggestion="修正乙字。",
        chunk_index=target_chunk.index,
        global_start=None,
        global_end=None,
    )
    output = tmp_path / "out.docx"

    summary = docx_service.write_docx_result(source, [issue], "comment", output)

    assert summary.comment_count == 1
    assert summary.fallback_count == 0
    with zipfile.ZipFile(output) as archive:
        comments_xml = archive.read("word/comments.xml").decode()
    assert "修正乙字" in comments_xml


def test_write_docx_result_inserts_comment_across_split_runs(tmp_path: Path):
    source = make_docx(["这里有错字。"])
    with zipfile.ZipFile(io.BytesIO(source)) as archive:
        document_xml = archive.read("word/document.xml").decode()
    document_xml = document_xml.replace(
        "<w:r><w:t>这里有错字。</w:t></w:r>",
        "<w:r><w:t>这里有错</w:t></w:r><w:r><w:t>字。</w:t></w:r>",
    )
    source = replace_docx_entry(source, "word/document.xml", document_xml.encode())
    document = docx_service.parse_docx(source)
    start = document.text.index("错字")
    issue = docx_service.ChunkedProofreadIssue(
        id="issue-1",
        category="typo",
        severity="high",
        original="错字",
        suggestion="修正错字。",
        chunk_index=0,
        global_start=start,
        global_end=start + 2,
    )
    output = tmp_path / "out.docx"

    summary = docx_service.write_docx_result(source, [issue], "comment", output)

    assert summary.comment_count == 1
    assert summary.fallback_count == 0
    with zipfile.ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml").decode()
        comments_xml = archive.read("word/comments.xml").decode()
    assert "commentRangeStart" in document_xml
    assert "commentRangeEnd" in document_xml
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

    summary = docx_service.write_docx_result(
        source,
        [chunked_issue],
        "revision",
        output,
        author="责任编辑",
    )

    assert summary.comment_count == 1
    assert summary.revision_count == 1
    with zipfile.ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml").decode()
        comments_xml = archive.read("word/comments.xml").decode()
    assert "commentRangeStart" in document_xml
    assert "<w:del" in document_xml
    assert "<w:ins" in document_xml
    assert 'w:author="责任编辑"' in comments_xml
    assert 'w:author="责任编辑"' in document_xml
    assert "Word AI Proofreader" not in comments_xml
    assert "Word AI Proofreader" not in document_xml
    assert "正字" in document_xml
    assert "修正错字" in comments_xml
    assert document_xml.index("<w:del") < document_xml.index("commentRangeStart")
    assert document_xml.index("commentRangeStart") < document_xml.index("<w:ins")
    assert document_xml.index("<w:ins") < document_xml.index("commentRangeEnd")


def test_write_docx_result_inserts_revision_across_split_runs(tmp_path: Path):
    source = make_docx(["这里有错字。"])
    with zipfile.ZipFile(io.BytesIO(source)) as archive:
        document_xml = archive.read("word/document.xml").decode()
    document_xml = document_xml.replace(
        "<w:r><w:t>这里有错字。</w:t></w:r>",
        "<w:r><w:t>这里有错</w:t></w:r><w:r><w:t>字。</w:t></w:r>",
    )
    source = replace_docx_entry(source, "word/document.xml", document_xml.encode())
    document = docx_service.parse_docx(source)
    start = document.text.index("错字")
    issue = docx_service.ChunkedProofreadIssue(
        id="issue-1",
        category="typo",
        severity="high",
        original="错字",
        replacement="正字",
        suggestion="修正错字。",
        chunk_index=0,
        global_start=start,
        global_end=start + 2,
    )
    output = tmp_path / "out.docx"

    summary = docx_service.write_docx_result(source, [issue], "revision", output)

    assert summary.comment_count == 1
    assert summary.revision_count == 1
    assert summary.fallback_count == 0
    with zipfile.ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml").decode()
        comments_xml = archive.read("word/comments.xml").decode()
    root = ET.fromstring(document_xml)
    namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    assert len(root.findall(".//w:del", namespaces)) == 2
    assert len(root.findall(".//w:ins", namespaces)) == 1
    assert "正字" in document_xml
    assert "修正错字" in comments_xml


def test_write_docx_result_does_not_insert_revision_across_paragraphs(tmp_path: Path):
    source = make_docx(["甲", "乙"])
    issue = docx_service.ChunkedProofreadIssue(
        id="issue-1",
        category="typo",
        severity="high",
        original="甲乙",
        replacement="甲丙",
        suggestion="跨段落不应直接修订。",
        chunk_index=0,
        global_start=0,
        global_end=2,
    )
    output = tmp_path / "out.docx"

    summary = docx_service.write_docx_result(
        source,
        [issue],
        "revision",
        output,
        fallback_summary_truncate_enabled=False,
    )

    assert summary.comment_count == 0
    assert summary.revision_count == 0
    assert summary.fallback_count == 1
    with zipfile.ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml").decode()
        comments_xml = archive.read("word/comments.xml").decode()
    assert "<w:del" not in document_xml
    assert "<w:ins" not in document_xml
    assert "跨段落不应直接修订" in comments_xml


def test_write_docx_result_comment_mode_does_not_insert_revision(tmp_path: Path):
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

    summary = docx_service.write_docx_result(source, [chunked_issue], "comment", output)

    assert summary.comment_count == 1
    assert summary.revision_count == 0
    with zipfile.ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml").decode()
        comments_xml = archive.read("word/comments.xml").decode()
    assert "commentRangeStart" in document_xml
    assert "<w:del" not in document_xml
    assert "<w:ins" not in document_xml
    assert "修正错字" in comments_xml


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


def test_write_docx_result_truncates_summary_comments_by_default(tmp_path: Path):
    source = make_docx(["第一章 开始", "正文。"])
    issues = [
        docx_service.ChunkedProofreadIssue(
            id=f"issue-{index}",
            category="typo",
            severity="high",
            original=f"未定位 {index}",
            suggestion=f"marker-{index} " + ("长" * 1180),
            chunk_index=0,
            global_start=None,
            global_end=None,
        )
        for index in range(1, 12)
    ]
    output = tmp_path / "out.docx"

    summary = docx_service.write_docx_result(source, issues, "comment", output)

    assert summary.fallback_count == 10
    with zipfile.ZipFile(output) as archive:
        comments_xml = archive.read("word/comments.xml").decode()
    assert comments_xml.count("<w:comment ") == 10
    assert "marker-10" in comments_xml
    assert "marker-11" not in comments_xml
    assert "另有 1 条未定位问题未写入汇总批注" in comments_xml


def test_write_docx_result_unlimited_summary_comments_when_truncate_disabled(tmp_path: Path):
    source = make_docx(["第一章 开始", "正文。"])
    issues = [
        docx_service.ChunkedProofreadIssue(
            id=f"issue-{index}",
            category="typo",
            severity="high",
            original=f"未定位 {index}",
            suggestion=f"marker-{index} " + ("长" * 1180),
            chunk_index=0,
            global_start=None,
            global_end=None,
        )
        for index in range(1, 12)
    ]
    output = tmp_path / "out.docx"

    summary = docx_service.write_docx_result(
        source,
        issues,
        "comment",
        output,
        fallback_summary_truncate_enabled=False,
    )

    assert summary.fallback_count == 11
    with zipfile.ZipFile(output) as archive:
        comments_xml = archive.read("word/comments.xml").decode()
    assert comments_xml.count("<w:comment ") == 11
    assert "marker-11" in comments_xml
    assert "未写入汇总批注" not in comments_xml


def test_write_docx_result_splits_single_long_summary_issue_without_truncating(tmp_path: Path):
    source = make_docx(["第一章 开始", "正文。"])
    issue = docx_service.ChunkedProofreadIssue(
        id="issue-long",
        category="typo",
        severity="high",
        original="未定位长问题",
        suggestion=("长建议" * 900) + "tail-marker",
        chunk_index=0,
        global_start=None,
        global_end=None,
    )
    output = tmp_path / "out.docx"

    summary = docx_service.write_docx_result(
        source,
        [issue],
        "comment",
        output,
        fallback_summary_truncate_enabled=False,
    )

    assert summary.fallback_count == 1
    with zipfile.ZipFile(output) as archive:
        comments_xml = archive.read("word/comments.xml").decode()
    assert comments_xml.count("<w:comment ") > 1
    assert "tail-marker" in comments_xml
    assert "已截断" not in comments_xml


def test_build_output_filename_uses_utc_plus_8_timestamp(monkeypatch):
    class FixedDatetime:
        @classmethod
        def now(cls, tz=None):
            assert tz is docx_service.OUTPUT_FILENAME_TIMEZONE
            return datetime(2026, 5, 2, 9, 2, 3, tzinfo=tz)

    monkeypatch.setattr(docx_service, "datetime", FixedDatetime)

    assert docx_service.build_output_filename("书稿.docx", "comment") == "书稿-AI审校-批注-20260502090203.docx"
    assert docx_service.build_output_filename("书稿.docx", "revision") == "书稿-AI审校-修订批注-20260502090203.docx"


def test_docx_task_api_uploads_generates_and_downloads(monkeypatch, tmp_path: Path):
    async def fake_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
        reasoning_enabled=False,
        temperature=0.2,
    ):
        assert "错字" in text
        assert temperature == 0.8
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
            "temperature": 0.8,
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


def test_docx_task_api_passes_summary_truncate_setting_and_author(monkeypatch, tmp_path: Path):
    async def fake_proofread_text(
        text,
        book,
        session_id=None,
        provider_api=None,
        proofread_mode="fast",
        reasoning_enabled=False,
        temperature=0.2,
    ):
        return []

    captured: dict[str, object] = {}
    original_write_docx_result = docx_service.write_docx_result

    def spy_write_docx_result(
        source_bytes,
        issues,
        application_mode,
        output_path,
        fallback_summary_truncate_enabled=True,
        author=docx_service.DEFAULT_WRITEBACK_AUTHOR,
    ):
        captured["fallback_summary_truncate_enabled"] = fallback_summary_truncate_enabled
        captured["author"] = author
        return original_write_docx_result(
            source_bytes,
            issues,
            application_mode,
            output_path,
            fallback_summary_truncate_enabled=fallback_summary_truncate_enabled,
            author=author,
        )

    monkeypatch.setattr(docx_task_service, "proofread_text", fake_proofread_text)
    monkeypatch.setattr(docx_task_service.docx_service, "write_docx_result", spy_write_docx_result)
    monkeypatch.setenv("DOCX_OUTPUT_DIR", str(tmp_path / "docx-results"))
    response = client.post(
        "/api/proofread/docx/tasks",
        params={
            "filename": "书稿.docx",
            "book": json.dumps(BOOK, ensure_ascii=False),
            "application_mode": "comment",
            "fallback_summary_truncate_enabled": "false",
            "author": "责任编辑",
        },
        content=make_docx(["第一章 开始", "没有问题。"]),
        headers={"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
    )

    assert response.status_code == 200
    task_id = response.json()["task_id"]
    with client.stream("GET", f"/api/proofread/docx/tasks/{task_id}/events") as stream:
        assert "completed" in stream.read().decode()
    assert captured == {"fallback_summary_truncate_enabled": False, "author": "责任编辑"}


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
    run_id = response.json()["run_id"]
    assert run_id.startswith("agent_run_")

    with client.stream("GET", f"/api/proofread/docx/tasks/{task_id}/events") as stream:
        assert "completed" in stream.read().decode()

    docx_task_service.clear_tasks_for_tests()

    final = client.get(f"/api/proofread/docx/tasks/{task_id}")
    assert final.status_code == 200
    final_payload = final.json()
    assert final_payload["run_id"] == run_id
    assert final_payload["output_filename"].endswith(".docx")

    trace_response = client.get(f"/api/agent/runs/{run_id}/trace")
    assert trace_response.status_code == 200
    trace_payload = trace_response.json()
    assert trace_payload["run_id"] == run_id
    assert any(node["node_name"] == "proofread_chunk" for node in trace_payload["nodes"])
    assert any(node["node_name"] == "write_docx_output" for node in trace_payload["nodes"])
    assert trace_payload["chunks"]

    download = client.get(f"/api/proofread/docx/tasks/{task_id}/download")
    assert download.status_code == 200
    with zipfile.ZipFile(io.BytesIO(download.content)) as archive:
        assert "word/document.xml" in archive.namelist()


def test_docx_store_migrates_legacy_results_without_run_id(monkeypatch, tmp_path: Path):
    monkeypatch.setenv("DOCX_OUTPUT_DIR", str(tmp_path / "docx-results"))
    root = docx_store.output_dir()
    root.mkdir(parents=True)
    legacy_file = root / "legacy-task" / "legacy.docx"
    legacy_file.parent.mkdir(parents=True)
    legacy_file.write_bytes(b"legacy")
    db_path = root / "results.sqlite3"
    now = datetime.now(UTC)

    import sqlite3

    with sqlite3.connect(db_path) as connection:
        connection.execute(
            """
            CREATE TABLE docx_results (
                task_id TEXT PRIMARY KEY,
                source_filename TEXT NOT NULL,
                output_filename TEXT NOT NULL,
                application_mode TEXT NOT NULL,
                status TEXT NOT NULL,
                total_chunks INTEGER NOT NULL,
                completed_chunks INTEGER NOT NULL,
                failed_chunks INTEGER NOT NULL,
                issue_count INTEGER NOT NULL,
                relative_path TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                expires_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            INSERT INTO docx_results (
                task_id, source_filename, output_filename, application_mode, status,
                total_chunks, completed_chunks, failed_chunks, issue_count,
                relative_path, created_at, updated_at, expires_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "legacy-task",
                "源.docx",
                "legacy.docx",
                "comment",
                "succeeded",
                1,
                1,
                0,
                0,
                "legacy-task/legacy.docx",
                now.isoformat(),
                now.isoformat(),
                (now + timedelta(days=7)).isoformat(),
            ),
        )
        connection.commit()

    stored = docx_store.get_result("legacy-task")

    assert stored is not None
    assert stored.run_id is None
    assert docx_store.resolve_download("legacy-task")[1].run_id is None


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
