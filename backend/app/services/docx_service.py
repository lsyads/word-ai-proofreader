from __future__ import annotations

from pathlib import Path

from app.schemas import ChunkedProofreadIssue, ProofreadChunk
from app.services import docx


def parse_docx(document_bytes: bytes) -> docx.DocxDocument:
    """Parse a DOCX package and extract its text model."""
    return docx.parse_docx(document_bytes)


def split_docx_into_chunks(
    document: docx.DocxDocument,
    chunk_size: int = docx.chunking.DEFAULT_CHUNK_SIZE,
) -> list[ProofreadChunk]:
    """Split an extracted DOCX document into proofreading chunks."""
    return docx.split_docx_into_chunks(document, chunk_size)


def write_docx_result(
    source_bytes: bytes,
    issues: list[ChunkedProofreadIssue],
    application_mode: docx.ApplicationMode,
    output_path: Path,
    *,
    fallback_summary_truncate_enabled: bool = True,
) -> docx.WritebackSummary:
    """Write proofreading issues back into a new DOCX result file."""
    return docx.write_docx_result(
        source_bytes,
        issues,
        application_mode,
        output_path,
        fallback_summary_truncate_enabled=fallback_summary_truncate_enabled,
    )

