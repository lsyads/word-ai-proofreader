from __future__ import annotations

from pathlib import Path
from typing import Literal

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

from app.schemas import BookInfo, ChunkedProofreadIssue, ProofreadChunk, ProofreadIssue, ProofreadScope
from app.services import chunk_service, docx_service, locator_service
from app.services.ai_client import DEFAULT_TEMPERATURE
from app.services.ai_provider_service import proofread_text_with_provider
from app.services.proofread import ProofreadMode, ProviderAPI


class SplitTextInput(BaseModel):
    text: str = Field(..., min_length=1)
    scope: ProofreadScope = "selection"
    chunk_size: int = Field(default=5000, ge=500, le=10000)


class ProofreadTextInput(BaseModel):
    text: str = Field(..., min_length=1)
    book: BookInfo
    session_id: str | None = None
    ai_profile_id: str | None = None
    provider_api: ProviderAPI | None = None
    proofread_mode: ProofreadMode = "fast"
    reasoning_enabled: bool = False
    temperature: float = Field(default=DEFAULT_TEMPERATURE, ge=0, le=1.5)


class GlobalizeIssuesInput(BaseModel):
    chunk: ProofreadChunk
    issues: list[ProofreadIssue]


class LocateIssuesInput(BaseModel):
    text: str = Field(..., min_length=1)
    issues: list[ProofreadIssue]


class ParseDocxInput(BaseModel):
    content: bytes = Field(..., min_length=1)


class WriteDocxInput(BaseModel):
    source_bytes: bytes = Field(..., min_length=1)
    issues: list[ChunkedProofreadIssue]
    application_mode: Literal["comment", "revision"] = "comment"
    output_path: Path
    fallback_summary_truncate_enabled: bool = True


def split_text(text: str, scope: ProofreadScope = "selection", chunk_size: int = 5000) -> list[ProofreadChunk]:
    """Split plain text into proofread chunks. Input is source text, scope, and chunk size; output is ordered chunks without storing full text in trace."""
    return chunk_service.split_text_into_chunks(text, scope, chunk_size)


async def proofread_with_model(
    text: str,
    book: BookInfo,
    session_id: str | None = None,
    ai_profile_id: str | None = None,
    provider_api: ProviderAPI | None = None,
    proofread_mode: ProofreadMode = "fast",
    reasoning_enabled: bool = False,
    temperature: float = DEFAULT_TEMPERATURE,
) -> list[ProofreadIssue]:
    """Run AI proofreading for one text chunk. Input is chunk text plus model options; output is normalized issues located inside that chunk."""
    return await proofread_text_with_provider(
        text,
        book,
        session_id=session_id,
        ai_profile_id=ai_profile_id,
        provider_api=provider_api,
        proofread_mode=proofread_mode,
        reasoning_enabled=reasoning_enabled,
        temperature=temperature,
    )


def globalize_issue_offsets(chunk: ProofreadChunk, issues: list[ProofreadIssue]) -> list[ChunkedProofreadIssue]:
    """Translate chunk-local issue offsets to document-level offsets. Input is one chunk and its issues; output is issues carrying chunk_index/global_start/global_end."""
    return chunk_service.globalize_issues(chunk, issues)


def locate_issue_offsets(text: str, issues: list[ProofreadIssue]) -> list[ProofreadIssue]:
    """Locate issue originals in text and build replayable locators. Input is source text plus issues; output is the same issues with start/end/locator when found."""
    return locator_service.locate_issues(text, issues)


def parse_docx_document(content: bytes) -> dict[str, int]:
    """Parse a DOCX package and return safe metadata only. Input is DOCX bytes; output is extracted character/block counts, not full document text."""
    document = docx_service.parse_docx(content)
    return {"text_len": len(document.text), "block_count": len(document.blocks), "span_count": len(document.spans)}


def write_docx_output(
    source_bytes: bytes,
    issues: list[ChunkedProofreadIssue],
    output_path: Path,
    application_mode: Literal["comment", "revision"] = "comment",
    fallback_summary_truncate_enabled: bool = True,
) -> dict[str, int]:
    """Write proofread issues into a DOCX result file. Input is source bytes, issues, output path, mode, and fallback policy; output is writeback counts."""
    summary = docx_service.write_docx_result(
        source_bytes,
        issues,
        application_mode,
        output_path,
        fallback_summary_truncate_enabled=fallback_summary_truncate_enabled,
    )
    return {
        "comment_count": summary.comment_count,
        "revision_count": summary.revision_count,
        "fallback_count": summary.fallback_count,
        "failed_count": summary.failed_count,
    }


proofread_tools = [
    StructuredTool.from_function(
        split_text,
        name="split_text",
        description=split_text.__doc__ or "",
        args_schema=SplitTextInput,
    ),
    StructuredTool.from_function(
        coroutine=proofread_with_model,
        name="proofread_with_model",
        description=proofread_with_model.__doc__ or "",
        args_schema=ProofreadTextInput,
    ),
    StructuredTool.from_function(
        globalize_issue_offsets,
        name="globalize_issue_offsets",
        description=globalize_issue_offsets.__doc__ or "",
        args_schema=GlobalizeIssuesInput,
    ),
    StructuredTool.from_function(
        locate_issue_offsets,
        name="locate_issue_offsets",
        description=locate_issue_offsets.__doc__ or "",
        args_schema=LocateIssuesInput,
    ),
    StructuredTool.from_function(
        parse_docx_document,
        name="parse_docx_document",
        description=parse_docx_document.__doc__ or "",
        args_schema=ParseDocxInput,
    ),
    StructuredTool.from_function(
        write_docx_output,
        name="write_docx_output",
        description=write_docx_output.__doc__ or "",
        args_schema=WriteDocxInput,
    ),
]


def get_proofread_tools() -> list[StructuredTool]:
    """Return the LangChain tool wrappers used by the Agent V1 workflow."""
    return proofread_tools
