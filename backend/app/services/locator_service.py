from __future__ import annotations

from app.schemas import ProofreadIssue, ProofreadLocator
from app.services import proofread


def locate_issues(text: str, issues: list[ProofreadIssue]) -> list[ProofreadIssue]:
    """Locate normalized proofreading issues in the supplied source text."""
    return proofread.locate_issues(text, issues)


def build_locator(
    text: str,
    original: str,
    start: int | None,
    end: int | None,
) -> ProofreadLocator | None:
    """Build a replayable locator for an already located issue range."""
    return proofread.build_locator(text, original, start, end)

