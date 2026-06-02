from __future__ import annotations

from collections.abc import Awaitable, Callable
from typing import Any, Literal, NotRequired, TypedDict

from pydantic import BaseModel, Field

from app.schemas import (
    ApplicationMode,
    BookInfo,
    ChunkedProofreadIssue,
    ChunkedTaskStatus,
    ProofreadChunk,
    ProofreadIssue,
    ProofreadScope,
)
from app.services.ai_client import DEFAULT_TEMPERATURE
from app.services.proofread import ProofreadMode, ProviderAPI


AgentFlow = Literal["selection", "selection_stream", "chunked", "chunked_task", "docx_task"]
NodeTraceStatus = Literal["running", "succeeded", "failed"]

ProofreadChunkCallable = Callable[..., Awaitable[list[ProofreadIssue]]]


class AgentOptions(BaseModel):
    session_id: str | None = None
    ai_profile_id: str | None = None
    provider_api: ProviderAPI | None = None
    proofread_mode: ProofreadMode = "fast"
    reasoning_enabled: bool = False
    temperature: float = DEFAULT_TEMPERATURE
    scope: ProofreadScope = "selection"
    chunk_size: int = 5000
    application_mode: ApplicationMode = "comment"
    fallback_summary_truncate_enabled: bool = True


class AgentRunSummary(BaseModel):
    run_id: str
    flow: AgentFlow
    task_id: str | None = None
    status: ChunkedTaskStatus | NodeTraceStatus
    total_chunks: int = 0
    completed_chunks: int = 0
    failed_chunks: int = 0
    issue_count: int = 0
    error_message: str | None = None


class AgentNodeTrace(BaseModel):
    node_name: str
    status: NodeTraceStatus
    started_at: str
    ended_at: str | None = None
    elapsed_seconds: float | None = None
    error_message: str | None = None


class AgentChunkTrace(BaseModel):
    chunk_index: int = Field(..., ge=0)
    chunk_start: int = Field(..., ge=0)
    chunk_end: int = Field(..., ge=0)
    chunk_len: int = Field(..., ge=0)
    status: NodeTraceStatus
    issue_count: int = Field(default=0, ge=0)
    retry_count: int = Field(default=0, ge=0)
    error_message: str | None = None
    started_at: str
    ended_at: str | None = None
    elapsed_seconds: float | None = None


class AgentRunTrace(BaseModel):
    run_id: str
    flow: AgentFlow
    task_id: str | None = None
    status: ChunkedTaskStatus | NodeTraceStatus
    created_at: str
    updated_at: str
    total_chunks: int
    completed_chunks: int
    failed_chunks: int
    issue_count: int
    error_message: str | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)
    nodes: list[AgentNodeTrace] = Field(default_factory=list)
    chunks: list[AgentChunkTrace] = Field(default_factory=list)


class AgentState(TypedDict):
    run_id: str
    flow: AgentFlow
    text: str
    book: BookInfo
    options: AgentOptions
    proofread_callable: ProofreadChunkCallable
    continue_on_chunk_error: bool
    chunks: NotRequired[list[ProofreadChunk]]
    issues: NotRequired[list[ChunkedProofreadIssue]]
    completed_chunks: NotRequired[int]
    failed_chunks: NotRequired[int]
    status: NotRequired[ChunkedTaskStatus]
    error_message: NotRequired[str | None]

