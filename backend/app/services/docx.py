from __future__ import annotations

import copy
import re
import zipfile
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Literal
from xml.etree import ElementTree as ET
from xml.sax.saxutils import quoteattr

from app.schemas import ChunkedProofreadIssue, ProofreadChunk, ProofreadIssue
from app.services import chunking

ApplicationMode = Literal["comment", "revision"]

WORD_DOCUMENT_PATH = "word/document.xml"
WORD_RELS_PATH = "word/_rels/document.xml.rels"
COMMENTS_PATH = "word/comments.xml"
CONTENT_TYPES_PATH = "[Content_Types].xml"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"

COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
OUTPUT_FILENAME_TIMEZONE = timezone(timedelta(hours=8))

CHAPTER_RE = re.compile(r"^\s*第[一二三四五六七八九十百千万零〇\d]+[章篇部卷]\b")
SECTION_RE = re.compile(r"^\s*第[一二三四五六七八九十百千万零〇\d]+[节回]\b")
TOC_STYLE_RE = re.compile(r"^toc\d*|^TOC\d*|目录", re.IGNORECASE)
TRAILING_PAGE_NUMBER_RE = re.compile(r"[\s.\u00a0·…\t]+[ivxlcdmIVXLCDM\d一二三四五六七八九十百千万零〇]+$")

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("rel", REL_NS)
ET.register_namespace("", CT_NS)


@dataclass
class TextSpan:
    start: int
    end: int
    node: ET.Element
    paragraph: ET.Element
    block_index: int
    in_textbox: bool


@dataclass
class DocxBlock:
    index: int
    start: int
    end: int
    text: str
    paragraph: ET.Element
    style: str | None = None
    in_textbox: bool = False


@dataclass
class DocxDocument:
    entries: dict[str, bytes]
    document_root: ET.Element
    document_namespaces: list[tuple[str, str]]
    text: str
    blocks: list[DocxBlock]
    spans: list[TextSpan]


@dataclass
class WritebackSummary:
    comment_count: int = 0
    revision_count: int = 0
    fallback_count: int = 0
    failed_count: int = 0


class DocxError(ValueError):
    """Raised when a DOCX package cannot be parsed or written safely."""


def parse_docx(document_bytes: bytes) -> DocxDocument:
    try:
        with zipfile.ZipFile(BytesIO(document_bytes), "r") as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
    except zipfile.BadZipFile as exc:
        raise DocxError("上传文件不是有效的 .docx 文件。") from exc

    if WORD_DOCUMENT_PATH not in entries:
        raise DocxError("上传文件缺少 word/document.xml，无法作为 .docx 审校。")

    document_namespaces = _collect_namespaces(entries[WORD_DOCUMENT_PATH])
    _register_namespaces(document_namespaces)
    document_root = ET.fromstring(entries[WORD_DOCUMENT_PATH])
    parent_map = _build_parent_map(document_root)
    body = document_root.find(f".//{_w('body')}")
    if body is None:
        raise DocxError("DOCX 正文为空，无法审校。")

    full_text_parts: list[str] = []
    blocks: list[DocxBlock] = []
    spans: list[TextSpan] = []
    position = 0

    for paragraph in body.iter(_w("p")):
        text_nodes = [
            node
            for node in paragraph.iter(_w("t"))
            if node.text and _ancestor(node, parent_map, _w("p")) is paragraph
        ]
        paragraph_text = "".join(node.text or "" for node in text_nodes)
        if not paragraph_text.strip():
            continue

        if full_text_parts:
            full_text_parts.append("\n")
            position += 1

        block_index = len(blocks)
        block_start = position
        in_textbox = any(_has_ancestor(node, parent_map, _w("txbxContent")) for node in text_nodes)

        for node in text_nodes:
            node_text = node.text or ""
            span_start = position
            full_text_parts.append(node_text)
            position += len(node_text)
            spans.append(
                TextSpan(
                    start=span_start,
                    end=position,
                    node=node,
                    paragraph=paragraph,
                    block_index=block_index,
                    in_textbox=in_textbox,
                )
            )

        blocks.append(
            DocxBlock(
                index=block_index,
                start=block_start,
                end=position,
                text=paragraph_text,
                paragraph=paragraph,
                style=_paragraph_style(paragraph),
                in_textbox=in_textbox,
            )
        )

    text = "".join(full_text_parts)
    if not text.strip():
        raise DocxError("DOCX 未提取到可审校文字。")

    return DocxDocument(
        entries=entries,
        document_root=document_root,
        document_namespaces=document_namespaces,
        text=text,
        blocks=blocks,
        spans=spans,
    )


def split_docx_into_chunks(document: DocxDocument, chunk_size: int = chunking.DEFAULT_CHUNK_SIZE) -> list[ProofreadChunk]:
    toc_titles = _extract_toc_titles(document.blocks)
    block_groups = _split_blocks_by_boundaries(document.blocks, "chapter")
    refined_groups: list[list[DocxBlock]] = []

    for group in block_groups:
        refined_groups.extend(_split_large_group(group, "section"))

    toc_refined_groups: list[list[DocxBlock]] = []
    for group in refined_groups:
        toc_refined_groups.extend(_split_large_group_by_toc_titles(group, toc_titles))

    chunks: list[ProofreadChunk] = []
    for group in toc_refined_groups:
        chunks.extend(_chunks_for_group(document.text, group, chunk_size))

    return [
        ProofreadChunk(index=index, start=chunk.start, end=chunk.end, text=chunk.text)
        for index, chunk in enumerate(chunks)
    ]


def write_docx_result(
    source_bytes: bytes,
    issues: list[ChunkedProofreadIssue],
    application_mode: ApplicationMode,
    output_path: Path,
) -> WritebackSummary:
    document = parse_docx(source_bytes)
    package = _DocxPackage(document.entries, document.document_root)
    summary = WritebackSummary()
    fallback_issues: list[ChunkedProofreadIssue] = []

    sorted_issues = sorted(
        issues,
        key=lambda issue: issue.global_start if issue.global_start is not None else -1,
        reverse=True,
    )

    for issue in sorted_issues:
        if issue.global_start is None or issue.global_end is None:
            fallback_issues.append(issue)
            continue

        if application_mode == "revision" and issue.replacement:
            if _try_insert_revision(document, issue.global_start, issue.global_end, issue, package):
                summary.revision_count += 1
                continue

        if _try_insert_comment(document, issue.global_start, issue.global_end, issue, package):
            summary.comment_count += 1
            continue

        fallback_issues.append(issue)

    if fallback_issues:
        if _insert_summary_comment(document, fallback_issues, package):
            summary.fallback_count = len(fallback_issues)
        else:
            summary.failed_count = len(fallback_issues)

    package.save(output_path)
    return summary


def build_output_filename(source_filename: str, application_mode: ApplicationMode) -> str:
    stem = Path(source_filename).stem.strip() or "审校文件"
    suffix = "修订" if application_mode == "revision" else "批注"
    timestamp = datetime.now(OUTPUT_FILENAME_TIMEZONE).strftime("%Y%m%d%H%M%S")
    safe_stem = re.sub(r'[\\/:*?"<>|]+', "_", stem)
    return f"{safe_stem}-AI审校-{suffix}-{timestamp}.docx"


def _split_large_group(group: list[DocxBlock], boundary_kind: str) -> list[list[DocxBlock]]:
    if _group_length(group) <= chunking.SELECTION_CHUNK_THRESHOLD:
        return [group]

    split_groups = _split_blocks_by_boundaries(group, boundary_kind)
    if len(split_groups) == 1:
        return split_groups

    return split_groups


def _split_large_group_by_toc_titles(
    group: list[DocxBlock],
    toc_titles: set[str],
) -> list[list[DocxBlock]]:
    if _group_length(group) <= chunking.SELECTION_CHUNK_THRESHOLD or not toc_titles:
        return [group]

    split_groups = _split_blocks_by_toc_titles(group, toc_titles)
    if len(split_groups) == 1:
        return [group]

    return split_groups


def _split_blocks_by_boundaries(blocks: list[DocxBlock], boundary_kind: str) -> list[list[DocxBlock]]:
    groups: list[list[DocxBlock]] = []
    current: list[DocxBlock] = []

    for block in blocks:
        is_boundary = _is_boundary_block(block, boundary_kind)
        if current and is_boundary:
            groups.append(current)
            current = []
        current.append(block)

    if current:
        groups.append(current)

    return groups or [blocks]


def _split_blocks_by_toc_titles(blocks: list[DocxBlock], toc_titles: set[str]) -> list[list[DocxBlock]]:
    groups: list[list[DocxBlock]] = []
    current: list[DocxBlock] = []

    for block in blocks:
        is_boundary = _normalized_toc_title(block.text) in toc_titles and not _is_toc_entry_block(block)
        if current and is_boundary:
            groups.append(current)
            current = []
        current.append(block)

    if current:
        groups.append(current)

    return groups or [blocks]


def _chunks_for_group(text: str, group: list[DocxBlock], chunk_size: int) -> list[ProofreadChunk]:
    if not group:
        return []

    start = group[0].start
    end = group[-1].end
    group_text = text[start:end]

    if len(group_text) <= chunking.SELECTION_CHUNK_THRESHOLD:
        return [ProofreadChunk(index=0, start=start, end=end, text=group_text)]

    local_chunks = chunking.split_text_into_chunks(group_text, "selection", chunk_size)
    return [
        ProofreadChunk(index=chunk.index, start=start + chunk.start, end=start + chunk.end, text=chunk.text)
        for chunk in local_chunks
    ]


def _group_length(group: list[DocxBlock]) -> int:
    if not group:
        return 0
    return group[-1].end - group[0].start


def _is_boundary_block(block: DocxBlock, boundary_kind: str) -> bool:
    text = block.text.strip()
    style = block.style or ""

    if boundary_kind == "chapter":
        return bool(CHAPTER_RE.match(text)) or style.lower() in {"heading1", "1", "标题1"}

    if boundary_kind == "section":
        return bool(SECTION_RE.match(text)) or style.lower() in {"heading2", "2", "标题2"}

    return False


def _extract_toc_titles(blocks: list[DocxBlock]) -> set[str]:
    titles: set[str] = set()

    for block in blocks:
        if not _is_toc_entry_block(block):
            continue

        title = _normalized_toc_title(block.text)
        if title:
            titles.add(title)

    return titles


def _is_toc_entry_block(block: DocxBlock) -> bool:
    style = block.style or ""
    return bool(TOC_STYLE_RE.search(style))


def _normalized_toc_title(text: str) -> str:
    title = text.strip()
    if not title:
        return ""

    title = title.replace("\u00a0", " ")
    title = re.sub(r"\s+", " ", title)
    title = TRAILING_PAGE_NUMBER_RE.sub("", title).strip()
    title = re.sub(r"^[\d一二三四五六七八九十百千万零〇]+(?:[.、．]\d+)*[.、．]?\s*", "", title)
    return title.strip()


def _try_insert_comment(
    document: DocxDocument,
    start: int,
    end: int,
    issue: ProofreadIssue,
    package: "_DocxPackage",
) -> bool:
    span = _single_span_for_range(document, start, end)
    if span is None:
        return False

    comment_id = package.add_comment(_format_issue_comment(issue, prefix="文本框" if span.in_textbox else None))
    return _wrap_single_text_span(document, span, start, end, package.comment_markers(comment_id))


def _try_insert_revision(
    document: DocxDocument,
    start: int,
    end: int,
    issue: ProofreadIssue,
    package: "_DocxPackage",
) -> bool:
    if not issue.replacement:
        return False

    span = _single_span_for_range(document, start, end)
    if span is None:
        return False

    parent_map = _build_parent_map(document.document_root)
    run = _ancestor(span.node, parent_map, _w("r"))
    if run is None:
        return False

    run_parent = parent_map.get(run)
    if run_parent is None:
        return False

    original_text = span.node.text or ""
    relative_start = start - span.start
    relative_end = end - span.start
    before = original_text[:relative_start]
    target = original_text[relative_start:relative_end]
    after = original_text[relative_end:]

    if target != issue.original:
        return False

    replacement_nodes: list[ET.Element] = []
    rpr = run.find(_w("rPr"))
    if before:
        replacement_nodes.append(_make_run(before, rpr))
    replacement_nodes.append(package.revision_delete(target, rpr))
    replacement_nodes.append(package.revision_insert(issue.replacement, rpr))
    if after:
        replacement_nodes.append(_make_run(after, rpr))

    _replace_child(run_parent, run, replacement_nodes)
    return True


def _insert_summary_comment(
    document: DocxDocument,
    issues: list[ChunkedProofreadIssue],
    package: "_DocxPackage",
) -> bool:
    if not document.blocks:
        return False

    comment_id = package.add_comment(_format_summary_comment(issues))
    paragraph = document.blocks[0].paragraph
    children = list(paragraph)
    first_run = next((child for child in children if child.tag == _w("r")), None)
    start_marker, end_marker, reference_run = package.comment_markers(comment_id)

    if first_run is None:
        paragraph.extend([start_marker, end_marker, reference_run])
        return True

    index = list(paragraph).index(first_run)
    paragraph.insert(index, start_marker)
    paragraph.insert(index + 2, end_marker)
    paragraph.insert(index + 3, reference_run)
    return True


def _single_span_for_range(document: DocxDocument, start: int, end: int) -> TextSpan | None:
    for span in document.spans:
        if span.start <= start and end <= span.end:
            return span
    return None


def _wrap_single_text_span(
    document: DocxDocument,
    span: TextSpan,
    start: int,
    end: int,
    markers: tuple[ET.Element, ET.Element, ET.Element],
) -> bool:
    parent_map = _build_parent_map(document.document_root)
    run = _ancestor(span.node, parent_map, _w("r"))
    if run is None:
        return False

    run_parent = parent_map.get(run)
    if run_parent is None:
        return False

    original_text = span.node.text or ""
    relative_start = start - span.start
    relative_end = end - span.start
    before = original_text[:relative_start]
    target = original_text[relative_start:relative_end]
    after = original_text[relative_end:]

    if not target:
        return False

    rpr = run.find(_w("rPr"))
    start_marker, end_marker, reference_run = markers
    replacement_nodes: list[ET.Element] = []

    if before:
        replacement_nodes.append(_make_run(before, rpr))
    replacement_nodes.append(start_marker)
    replacement_nodes.append(_make_run(target, rpr))
    replacement_nodes.append(end_marker)
    replacement_nodes.append(reference_run)
    if after:
        replacement_nodes.append(_make_run(after, rpr))

    _replace_child(run_parent, run, replacement_nodes)
    return True


def _format_issue_comment(issue: ProofreadIssue, prefix: str | None = None) -> str:
    lines: list[str] = []
    if prefix:
        lines.append(f"{prefix}内容审校建议")
    lines.append(f"[{issue.severity}] {issue.category}")
    if issue.replacement:
        lines.append(f"替换为：{issue.replacement}")
    lines.append(f"建议：{issue.suggestion}")
    return "\n".join(lines)


def _format_summary_comment(issues: list[ChunkedProofreadIssue]) -> str:
    lines = ["AI 审校未能精准定位的问题："]
    for index, issue in enumerate(issues[:20], start=1):
        lines.append("")
        lines.append(f"{index}. [{issue.severity}] {issue.category}")
        lines.append(f"原文：{issue.original or '未提供'}")
        if issue.replacement:
            lines.append(f"替换为：{issue.replacement}")
        lines.append(f"建议：{issue.suggestion}")
    if len(issues) > 20:
        lines.append("")
        lines.append(f"另有 {len(issues) - 20} 条未定位问题，请在任务结果中查看。")
    return "\n".join(lines)


def _replace_child(parent: ET.Element, old_child: ET.Element, new_children: list[ET.Element]) -> None:
    children = list(parent)
    index = children.index(old_child)
    parent.remove(old_child)
    for offset, child in enumerate(new_children):
        parent.insert(index + offset, child)


def _make_run(text: str, rpr: ET.Element | None = None, tag: str = "t") -> ET.Element:
    run = ET.Element(_w("r"))
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    text_node = ET.SubElement(run, _w(tag))
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set(f"{{{XML_NS}}}space", "preserve")
    text_node.text = text
    return run


def _paragraph_style(paragraph: ET.Element) -> str | None:
    p_style = paragraph.find(f"{_w('pPr')}/{_w('pStyle')}")
    if p_style is None:
        return None
    return p_style.get(_w("val"))


def _build_parent_map(root: ET.Element) -> dict[ET.Element, ET.Element]:
    return {child: parent for parent in root.iter() for child in parent}


def _has_ancestor(node: ET.Element, parent_map: dict[ET.Element, ET.Element], tag: str) -> bool:
    current = parent_map.get(node)
    while current is not None:
        if current.tag == tag:
            return True
        current = parent_map.get(current)
    return False


def _ancestor(node: ET.Element, parent_map: dict[ET.Element, ET.Element], tag: str) -> ET.Element | None:
    current = parent_map.get(node)
    while current is not None:
        if current.tag == tag:
            return current
        current = parent_map.get(current)
    return None


def _w(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def _r(local: str) -> str:
    return f"{{{R_NS}}}{local}"


def _collect_namespaces(xml_bytes: bytes) -> list[tuple[str, str]]:
    namespaces: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()

    for _, namespace in ET.iterparse(BytesIO(xml_bytes), events=("start-ns",)):
        prefix, uri = namespace
        key = (prefix or "", uri)
        if key in seen:
            continue
        seen.add(key)
        namespaces.append(key)

    return namespaces


def _register_namespaces(namespaces: list[tuple[str, str]]) -> None:
    for prefix, uri in namespaces:
        if not prefix or prefix.lower().startswith("xml") or re.match(r"^ns\d+$", prefix):
            continue
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            continue


def _restore_root_namespace_declarations(xml_bytes: bytes, namespaces: list[tuple[str, str]]) -> bytes:
    xml = xml_bytes.decode("utf-8")
    root_start = xml.find("<", xml.find("?>") + 2 if "?>" in xml[:120] else 0)
    if root_start == -1:
        return xml_bytes

    root_end = xml.find(">", root_start)
    if root_end == -1:
        return xml_bytes

    root_tag = xml[root_start:root_end]
    declarations: list[str] = []
    for prefix, uri in namespaces:
        if not prefix or prefix.lower().startswith("xml"):
            continue
        if f"xmlns:{prefix}=" in root_tag:
            continue
        declarations.append(f" xmlns:{prefix}={quoteattr(uri)}")

    if not declarations:
        return xml_bytes

    return f"{xml[:root_end]}{''.join(declarations)}{xml[root_end:]}".encode("utf-8")


class _DocxPackage:
    def __init__(self, entries: dict[str, bytes], document_root: ET.Element) -> None:
        self.entries = dict(entries)
        self.document_root = document_root
        self.document_namespaces = _collect_namespaces(entries[WORD_DOCUMENT_PATH])
        _register_namespaces(self.document_namespaces)
        self.comments_root = self._ensure_comments_root()
        self._ensure_comments_relationship()
        self._ensure_comments_content_type()
        self._next_comment_id = self._find_next_comment_id()
        self._next_revision_id = self._find_next_revision_id()

    def add_comment(self, text: str) -> int:
        comment_id = self._next_comment_id
        self._next_comment_id += 1

        comment = ET.SubElement(
            self.comments_root,
            _w("comment"),
            {
                _w("id"): str(comment_id),
                _w("author"): "Word AI Proofreader",
                _w("date"): datetime.now(UTC).isoformat(),
            },
        )
        for line in text.splitlines() or [""]:
            paragraph = ET.SubElement(comment, _w("p"))
            paragraph.append(_make_run(line))
        return comment_id

    def comment_markers(self, comment_id: int) -> tuple[ET.Element, ET.Element, ET.Element]:
        start_marker = ET.Element(_w("commentRangeStart"), {_w("id"): str(comment_id)})
        end_marker = ET.Element(_w("commentRangeEnd"), {_w("id"): str(comment_id)})
        reference_run = ET.Element(_w("r"))
        ET.SubElement(reference_run, _w("commentReference"), {_w("id"): str(comment_id)})
        return start_marker, end_marker, reference_run

    def revision_delete(self, text: str, rpr: ET.Element | None = None) -> ET.Element:
        revision_id = self._next_revision_id
        self._next_revision_id += 1
        deleted = ET.Element(
            _w("del"),
            {
                _w("id"): str(revision_id),
                _w("author"): "Word AI Proofreader",
                _w("date"): datetime.now(UTC).isoformat(),
            },
        )
        deleted.append(_make_run(text, rpr, tag="delText"))
        return deleted

    def revision_insert(self, text: str, rpr: ET.Element | None = None) -> ET.Element:
        revision_id = self._next_revision_id
        self._next_revision_id += 1
        inserted = ET.Element(
            _w("ins"),
            {
                _w("id"): str(revision_id),
                _w("author"): "Word AI Proofreader",
                _w("date"): datetime.now(UTC).isoformat(),
            },
        )
        inserted.append(_make_run(text, rpr))
        return inserted

    def save(self, output_path: Path) -> None:
        self.entries[WORD_DOCUMENT_PATH] = ET.tostring(
            self.document_root,
            encoding="utf-8",
            xml_declaration=True,
        )
        self.entries[WORD_DOCUMENT_PATH] = _restore_root_namespace_declarations(
            self.entries[WORD_DOCUMENT_PATH],
            self.document_namespaces,
        )
        self.entries[COMMENTS_PATH] = ET.tostring(
            self.comments_root,
            encoding="utf-8",
            xml_declaration=True,
        )
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, data in self.entries.items():
                archive.writestr(name, data)

    def _ensure_comments_root(self) -> ET.Element:
        if COMMENTS_PATH in self.entries:
            return ET.fromstring(self.entries[COMMENTS_PATH])
        return ET.Element(_w("comments"))

    def _ensure_comments_relationship(self) -> None:
        if WORD_RELS_PATH in self.entries:
            root = ET.fromstring(self.entries[WORD_RELS_PATH])
        else:
            root = ET.Element(f"{{{REL_NS}}}Relationships")

        existing = [
            relationship
            for relationship in root.findall(f"{{{REL_NS}}}Relationship")
            if relationship.get("Type") == COMMENTS_REL_TYPE
        ]
        if not existing:
            next_id = _next_relationship_id(root)
            ET.SubElement(
                root,
                f"{{{REL_NS}}}Relationship",
                {"Id": next_id, "Type": COMMENTS_REL_TYPE, "Target": "comments.xml"},
            )

        self.entries[WORD_RELS_PATH] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def _ensure_comments_content_type(self) -> None:
        if CONTENT_TYPES_PATH in self.entries:
            root = ET.fromstring(self.entries[CONTENT_TYPES_PATH])
        else:
            root = ET.Element(f"{{{CT_NS}}}Types")

        exists = any(
            override.get("PartName") == f"/{COMMENTS_PATH}"
            for override in root.findall(f"{{{CT_NS}}}Override")
        )
        if not exists:
            ET.SubElement(
                root,
                f"{{{CT_NS}}}Override",
                {"PartName": f"/{COMMENTS_PATH}", "ContentType": COMMENTS_CONTENT_TYPE},
            )

        self.entries[CONTENT_TYPES_PATH] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def _find_next_comment_id(self) -> int:
        ids = [
            int(comment.get(_w("id")))
            for comment in self.comments_root.findall(_w("comment"))
            if comment.get(_w("id"), "").isdigit()
        ]
        return max(ids, default=-1) + 1

    def _find_next_revision_id(self) -> int:
        ids: list[int] = []
        for element in self.document_root.iter():
            value = element.get(_w("id"))
            if value and value.isdigit():
                ids.append(int(value))
        return max(ids, default=0) + 1


def _next_relationship_id(root: ET.Element) -> str:
    used_ids = {relationship.get("Id") for relationship in root.findall(f"{{{REL_NS}}}Relationship")}
    index = 1
    while f"rId{index}" in used_ids:
        index += 1
    return f"rId{index}"
