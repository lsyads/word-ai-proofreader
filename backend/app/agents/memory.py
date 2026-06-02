from __future__ import annotations

from pydantic import BaseModel

from app.schemas import V2CandidateIssue


class V2MemoryItem(BaseModel):
    kind: str = "observation"
    key: str
    value: str
    source: str
    confidence: float = 0.7


def derive_memory_items(candidates: list[V2CandidateIssue]) -> list[V2MemoryItem]:
    """Build safe project memory hints from candidate metadata without storing full document text."""
    categories = sorted({candidate.category for candidate in candidates})
    severities = sorted({candidate.severity for candidate in candidates})
    passes = sorted({candidate.pass_name for candidate in candidates})
    items: list[V2MemoryItem] = []
    if categories:
        items.append(V2MemoryItem(key="observed_categories", value=", ".join(categories), source="candidate_summary"))
    if severities:
        items.append(V2MemoryItem(key="observed_severities", value=", ".join(severities), source="candidate_summary"))
    if passes:
        items.append(V2MemoryItem(key="observed_passes", value=", ".join(passes), source="candidate_summary"))
    approved_categories = sorted({candidate.category for candidate in candidates if candidate.status == "approved"})
    if approved_categories:
        items.append(
            V2MemoryItem(
                kind="preference",
                key="approved_issue_categories",
                value=", ".join(approved_categories),
                source="editor_decision",
                confidence=0.8,
            )
        )
    return items
