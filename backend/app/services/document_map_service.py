from __future__ import annotations

from app.schemas import V2DocumentBlock, V2DocumentChunk, V2DocumentMapResponse
from app.services import chunking
from app.services import docx as docx_service


PREVIEW_LENGTH = 80


def build_document_map(project_id: str, source_bytes: bytes) -> V2DocumentMapResponse:
    document = docx_service.parse_docx(source_bytes)
    chunks = docx_service.split_docx_into_chunks(document)
    return V2DocumentMapResponse(
        project_id=project_id,
        text_len=len(document.text),
        block_count=len(document.blocks),
        chunk_count=len(chunks),
        blocks=[
            V2DocumentBlock(
                index=block.index,
                start=block.start,
                end=block.end,
                text_preview=_preview(block.text),
                style=block.style,
                in_textbox=block.in_textbox,
            )
            for block in document.blocks
        ],
        chunks=[
            V2DocumentChunk(
                index=chunk.index,
                start=chunk.start,
                end=chunk.end,
                chunk_len=len(chunk.text),
                text_preview=_preview(chunk.text),
            )
            for chunk in chunks
        ],
    )


def build_selection_document_map(project_id: str, text: str) -> V2DocumentMapResponse:
    chunks = chunking.split_text_into_chunks(text, "selection", chunking.DEFAULT_CHUNK_SIZE)
    blocks: list[V2DocumentBlock] = []
    position = 0
    for paragraph in text.splitlines():
        block_text = paragraph.strip()
        raw_start = text.find(paragraph, position)
        if raw_start == -1:
            raw_start = position
        position = raw_start + len(paragraph)
        if not block_text:
            continue
        blocks.append(
            V2DocumentBlock(
                index=len(blocks),
                start=raw_start,
                end=raw_start + len(paragraph),
                text_preview=_preview(block_text),
                style=None,
                in_textbox=False,
            )
        )
    if not blocks and text.strip():
        blocks.append(
            V2DocumentBlock(
                index=0,
                start=0,
                end=len(text),
                text_preview=_preview(text),
                style=None,
                in_textbox=False,
            )
        )

    return V2DocumentMapResponse(
        project_id=project_id,
        text_len=len(text),
        block_count=len(blocks),
        chunk_count=len(chunks),
        blocks=blocks,
        chunks=[
            V2DocumentChunk(
                index=chunk.index,
                start=chunk.start,
                end=chunk.end,
                chunk_len=len(chunk.text),
                text_preview=_preview(chunk.text),
            )
            for chunk in chunks
        ],
    )


def _preview(text: str) -> str:
    normalized = " ".join(text.split())
    if len(normalized) <= PREVIEW_LENGTH:
        return normalized
    return f"{normalized[:PREVIEW_LENGTH]}..."
