from __future__ import annotations

from app.schemas import ChunkedProofreadIssue, ProofreadChunk, ProofreadIssue, ProofreadScope
from app.services import chunking


def split_text_into_chunks(
    text: str,
    scope: ProofreadScope = "selection",
    chunk_size: int = chunking.DEFAULT_CHUNK_SIZE,
) -> list[ProofreadChunk]:
    """Split plain text into proofread chunks using the existing boundary rules."""
    return chunking.split_text_into_chunks(text, scope, chunk_size)


def globalize_issues(chunk: ProofreadChunk, issues: list[ProofreadIssue]) -> list[ChunkedProofreadIssue]:
    """Convert chunk-local issue offsets into document-level offsets."""
    return chunking.globalize_issues(chunk, issues)

