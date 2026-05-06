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
    scope: ProofreadScope
    status: ChunkedTaskStatus
    total_chunks: int
    completed_chunks: int
    failed_chunks: int
    issues: list[ChunkedProofreadIssue]
    error_message: str | None = None


class DocxProofreadResult(BaseModel):
    task_id: str
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
