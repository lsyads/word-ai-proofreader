from __future__ import annotations

from collections import Counter
from datetime import UTC, datetime

from app.schemas import BookInfo, V2CandidateIssue, V2ProjectStatus, V2ReviewReportResponse


def build_review_report(
    *,
    project_id: str,
    status: V2ProjectStatus,
    source_filename: str,
    book: BookInfo,
    review_goal: str,
    candidates: list[V2CandidateIssue],
) -> V2ReviewReportResponse:
    severity_counts = Counter(candidate.severity for candidate in candidates)
    category_counts = Counter(candidate.category for candidate in candidates)
    pass_counts = Counter(candidate.pass_name for candidate in candidates)
    unresolved = [
        f"{candidate.severity}/{candidate.category}: {candidate.original[:40]} - {candidate.suggestion[:80]}"
        for candidate in candidates
        if candidate.status in {"pending", "deferred"}
    ]
    return V2ReviewReportResponse(
        project_id=project_id,
        status=status,
        source_filename=source_filename,
        book=book,
        review_goal=review_goal,
        issue_count=len(candidates),
        pending_count=sum(1 for candidate in candidates if candidate.status == "pending"),
        approved_count=sum(1 for candidate in candidates if candidate.status == "approved"),
        rejected_count=sum(1 for candidate in candidates if candidate.status == "rejected"),
        deferred_count=sum(1 for candidate in candidates if candidate.status == "deferred"),
        written_count=sum(1 for candidate in candidates if candidate.status == "written"),
        severity_counts=dict(severity_counts),
        category_counts=dict(category_counts),
        pass_counts=dict(pass_counts),
        unresolved_items=unresolved,
        generated_at=datetime.now(UTC).isoformat(),
    )
