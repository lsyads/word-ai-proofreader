from __future__ import annotations

from pydantic import BaseModel

from app.schemas import V2CandidateIssue


class V2MemoryItem(BaseModel):
    key: str
    value: str
    source: str


def derive_memory_items(candidates: list[V2CandidateIssue]) -> list[V2MemoryItem]:
    """Build safe project memory hints from candidate metadata without storing full document text."""
    categories = sorted({candidate.category for candidate in candidates})
    severities = sorted({candidate.severity for candidate in candidates})
    items: list[V2MemoryItem] = []
    if categories:
        items.append(V2MemoryItem(key="observed_categories", value=", ".join(categories), source="candidate_summary"))
    if severities:
        items.append(V2MemoryItem(key="observed_severities", value=", ".join(severities), source="candidate_summary"))
    return items
