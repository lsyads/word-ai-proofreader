from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator


class ProofreadRequest(BaseModel):
    text: str = Field(..., min_length=1)
    session_id: str | None = None
    provider_api: Literal["responses", "chat"] | None = None
    proofread_mode: Literal["fast", "thinking"] = "fast"
    context: dict[str, Any] | None = None

    @field_validator("text")
    @classmethod
    def text_must_not_be_blank(cls, value: str) -> str:
        normalized = value.strip()
        if not normalized:
            raise ValueError("text must not be blank")
        return normalized


class ProofreadIssue(BaseModel):
    id: str
    category: str
    severity: Literal["low", "medium", "high"]
    original: str
    replacement: str | None = None
    suggestion: str
    comment: str
    start: int | None = Field(default=None, ge=0)
    end: int | None = Field(default=None, ge=0)

    @field_validator("replacement")
    @classmethod
    def empty_replacement_to_none(cls, value: str | None) -> str | None:
        if value is None:
            return None

        normalized = value.strip()
        return normalized or None


class ProofreadResponse(BaseModel):
    issues: list[ProofreadIssue]


class SessionResponse(BaseModel):
    session_id: str
    created_at: str
