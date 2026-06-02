from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator


class BookInfo(BaseModel):
    title: str = Field(..., min_length=1)
    introduction: str | None = None

    @field_validator("title")
    @classmethod
    def title_must_not_be_blank(cls, value: str) -> str:
        normalized = value.strip()
        if not normalized:
            raise ValueError("book title must not be blank")
        return normalized

    @field_validator("introduction")
    @classmethod
    def empty_introduction_to_none(cls, value: str | None) -> str | None:
        if value is None:
            return None

        normalized = value.strip()
        return normalized or None


class ProofreadRequest(BaseModel):
    text: str = Field(..., min_length=1)
    book: BookInfo
    session_id: str | None = None
    ai_profile_id: str | None = None
    provider_api: Literal["responses", "chat"] | None = None
    proofread_mode: Literal["fast", "thinking"] = "fast"
    reasoning_enabled: bool = False
    temperature: float = Field(default=0.2, ge=0, le=1.5)
    context: dict[str, Any] | None = None

    @field_validator("text")
    @classmethod
    def text_must_not_be_blank(cls, value: str) -> str:
        normalized = value.strip()
        if not normalized:
            raise ValueError("text must not be blank")
        return normalized


class ProofreadLocator(BaseModel):
    key: str = Field(..., min_length=1)
    key_start: int = Field(..., ge=0)
    key_end: int = Field(..., ge=0)
    original_start_in_key: int = Field(..., ge=0)
    original_end_in_key: int = Field(..., ge=0)
    strategy: Literal["original", "context"]
    key_occurrence_index: int | None = Field(default=None, ge=0)


class ProofreadIssue(BaseModel):
    id: str
    category: str
    severity: Literal["low", "medium", "high"]
    original: str
    replacement: str | None = None
    suggestion: str
    start: int | None = Field(default=None, ge=0)
    end: int | None = Field(default=None, ge=0)
    locator: ProofreadLocator | None = None

    @field_validator("replacement")
    @classmethod
    def empty_replacement_to_none(cls, value: str | None) -> str | None:
        if value is None:
            return None

        normalized = value.strip()
        return normalized or None


class ProofreadResponse(BaseModel):
    issues: list[ProofreadIssue]
    run_id: str | None = None


ProofreadScope = Literal["selection", "document"]
ApplicationMode = Literal["comment", "revision"]
ChunkedTaskStatus = Literal[
    "queued",
    "running",
    "succeeded",
    "partial_succeeded",
    "failed",
    "cancelled",
]


class ProofreadChunk(BaseModel):
    index: int = Field(..., ge=0)
    start: int = Field(..., ge=0)
    end: int = Field(..., ge=0)
    text: str = Field(..., min_length=1)


class ChunkedProofreadIssue(ProofreadIssue):
    chunk_index: int = Field(..., ge=0)
    global_start: int | None = Field(default=None, ge=0)
    global_end: int | None = Field(default=None, ge=0)


class ChunkedProofreadRequest(ProofreadRequest):
    scope: ProofreadScope = "selection"
    chunk_size: int = Field(default=5000, ge=500, le=10000)


class ChunkedProofreadResult(BaseModel):
    task_id: str | None = None
    run_id: str | None = None
    scope: ProofreadScope
    status: ChunkedTaskStatus
    total_chunks: int
    completed_chunks: int
    failed_chunks: int
    issues: list[ChunkedProofreadIssue]
    error_message: str | None = None


class DocxProofreadResult(BaseModel):
    task_id: str
    run_id: str | None = None
    status: ChunkedTaskStatus
    total_chunks: int
    completed_chunks: int
    failed_chunks: int
    issue_count: int
    source_filename: str
    application_mode: ApplicationMode
    output_filename: str | None = None
    download_url: str | None = None
    expires_at: str | None = None
    retention_days: int | None = None
    error_message: str | None = None


class SessionResponse(BaseModel):
    session_id: str
    created_at: str


class AIProfileResponse(BaseModel):
    id: str
    label: str
    model: str
    default_api: Literal["responses", "chat"]
    supported_apis: list[Literal["responses", "chat"]]
    configured: bool


class AgentNodeTraceResponse(BaseModel):
    node_name: str
    status: Literal["running", "succeeded", "failed"]
    started_at: str
    ended_at: str | None = None
    elapsed_seconds: float | None = None
    error_message: str | None = None


class AgentChunkTraceResponse(BaseModel):
    chunk_index: int
    chunk_start: int
    chunk_end: int
    chunk_len: int
    status: Literal["running", "succeeded", "failed"]
    issue_count: int
    retry_count: int
    error_message: str | None = None
    started_at: str
    ended_at: str | None = None
    elapsed_seconds: float | None = None


class AgentRunTraceResponse(BaseModel):
    run_id: str
    flow: str
    task_id: str | None = None
    status: str
    created_at: str
    updated_at: str
    total_chunks: int
    completed_chunks: int
    failed_chunks: int
    issue_count: int
    error_message: str | None = None
    metadata: dict[str, Any]
    nodes: list[AgentNodeTraceResponse]
    chunks: list[AgentChunkTraceResponse]


V2ProjectStatus = Literal["created", "running", "waiting_for_approval", "written", "failed", "cancelled"]
V2RunStatus = Literal["queued", "running", "waiting_for_approval", "succeeded", "partial_succeeded", "failed", "cancelled"]
V2CandidateStatus = Literal["pending", "approved", "rejected", "deferred", "written"]
V2ProjectSourceType = Literal["selection", "docx"]
V2RunEventName = Literal[
    "project_created",
    "document_map_built",
    "plan_created",
    "tool_started",
    "tool_completed",
    "candidate_found",
    "waiting_for_approval",
    "writeback_completed",
    "report_ready",
    "error",
]


class V2ProjectCreateRequest(BaseModel):
    book: BookInfo
    review_goal: str = Field(default="完成全书出版审校，输出候选问题、证据、写回结果和审校报告。", min_length=1)


class V2SelectionProjectCreateRequest(V2ProjectCreateRequest):
    text: str = Field(..., min_length=1)
    session_id: str | None = None

    @field_validator("text")
    @classmethod
    def selection_text_must_not_be_blank(cls, value: str) -> str:
        normalized = value.strip()
        if not normalized:
            raise ValueError("selection text must not be blank")
        return normalized


class V2ProjectResponse(BaseModel):
    project_id: str
    source_type: V2ProjectSourceType
    status: V2ProjectStatus
    source_filename: str
    text_preview: str | None = None
    book: BookInfo
    review_goal: str
    created_at: str
    updated_at: str
    run_count: int = 0
    candidate_count: int = 0
    pending_count: int = 0
    approved_count: int = 0
    output_filename: str | None = None
    download_url: str | None = None


class V2DocumentBlock(BaseModel):
    index: int
    start: int
    end: int
    text_preview: str
    style: str | None = None
    in_textbox: bool = False


class V2DocumentChunk(BaseModel):
    index: int
    start: int
    end: int
    chunk_len: int
    text_preview: str


class V2DocumentMapResponse(BaseModel):
    project_id: str
    text_len: int
    block_count: int
    chunk_count: int
    blocks: list[V2DocumentBlock]
    chunks: list[V2DocumentChunk]


class V2ReviewPlanStep(BaseModel):
    step_id: str
    title: str
    tool_name: str
    status: Literal["pending", "running", "succeeded", "failed"] = "pending"
    description: str


class V2ReviewPlanResponse(BaseModel):
    project_id: str
    run_id: str | None = None
    steps: list[V2ReviewPlanStep]


class V2RunCreateRequest(BaseModel):
    session_id: str | None = None
    ai_profile_id: str | None = None
    provider_api: Literal["responses", "chat"] | None = None
    proofread_mode: Literal["fast", "thinking"] = "fast"
    reasoning_enabled: bool = False
    temperature: float = Field(default=0.2, ge=0, le=1.5)


class V2RunResponse(BaseModel):
    project_id: str
    run_id: str
    status: V2RunStatus
    stage: str
    total_chunks: int
    completed_chunks: int
    failed_chunks: int
    candidate_count: int
    error_message: str | None = None
    created_at: str
    updated_at: str


class V2CandidateIssue(BaseModel):
    candidate_id: str
    project_id: str
    run_id: str
    status: V2CandidateStatus = "pending"
    category: str
    severity: Literal["low", "medium", "high"]
    original: str
    replacement: str | None = None
    suggestion: str
    evidence: str
    chunk_index: int
    global_start: int | None = None
    global_end: int | None = None
    locator: ProofreadLocator | None = None
    self_check: str | None = None
    created_at: str
    updated_at: str


class V2CandidateListResponse(BaseModel):
    project_id: str
    candidates: list[V2CandidateIssue]


class V2ApprovalDecision(BaseModel):
    candidate_id: str
    status: Literal["approved", "rejected", "deferred"]


class V2ApprovalDecisionRequest(BaseModel):
    decisions: list[V2ApprovalDecision] = Field(..., min_length=1)


class V2ApprovalDecisionResponse(BaseModel):
    project_id: str
    updated_count: int
    candidates: list[V2CandidateIssue]


class V2MarkWrittenRequest(BaseModel):
    candidate_ids: list[str] = Field(..., min_length=1)


class V2MarkWrittenResponse(BaseModel):
    project_id: str
    updated_count: int
    candidates: list[V2CandidateIssue]


class V2WritebackRequest(BaseModel):
    application_mode: ApplicationMode = "comment"
    fallback_summary_truncate_enabled: bool = True


class V2WritebackResponse(BaseModel):
    project_id: str
    output_filename: str
    download_url: str
    comment_count: int
    revision_count: int
    fallback_count: int
    failed_count: int
    written_count: int


class V2RunEventResponse(BaseModel):
    event: V2RunEventName | str
    data: dict[str, Any]
    created_at: str


class V2RunTraceResponse(BaseModel):
    project_id: str
    run_id: str
    status: V2RunStatus
    events: list[V2RunEventResponse]


class V2ReviewReportResponse(BaseModel):
    project_id: str
    status: V2ProjectStatus
    source_filename: str
    book: BookInfo
    review_goal: str
    issue_count: int
    pending_count: int
    approved_count: int
    rejected_count: int
    deferred_count: int
    written_count: int
    severity_counts: dict[str, int]
    category_counts: dict[str, int]
    unresolved_items: list[str]
    generated_at: str
