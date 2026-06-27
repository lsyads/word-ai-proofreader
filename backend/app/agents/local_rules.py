from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


PassName = Literal["terminology_pass", "style_rule_pass", "consistency_pass"]
Severity = Literal["low", "medium", "high"]
EvidenceKind = Literal["rule", "document_map"]


@dataclass(frozen=True)
class LocalRuleMatch:
    rule_id: str
    pass_name: PassName
    category: str
    severity: Severity
    original: str
    replacement: str | None
    suggestion: str
    start: int
    end: int
    confidence: float
    evidence_kind: EvidenceKind


def run_terminology_rules(source_text: str) -> list[LocalRuleMatch]:
    """Keep the terminology pass callable, but emit no default non-AI local candidates."""
    return []


def run_style_rules(source_text: str) -> list[LocalRuleMatch]:
    """Keep the style-rule pass callable, but emit no default non-AI local candidates."""
    return []


def run_consistency_rules(source_text: str, *, source_type: str) -> list[LocalRuleMatch]:
    """Keep the consistency pass callable, but emit no default non-AI local candidates."""
    return []
