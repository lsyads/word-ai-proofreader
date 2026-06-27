from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass
from typing import Literal


PassName = Literal["terminology_pass", "style_rule_pass", "consistency_pass"]
Severity = Literal["low", "medium", "high"]
EvidenceKind = Literal["rule", "document_map"]
ReplacementBuilder = Callable[[re.Match[str]], str]
HIGH_CONFIDENCE_LOCAL_RULE_MIN = 0.85


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


@dataclass(frozen=True)
class RegexLocalRule:
    rule_id: str
    pass_name: PassName
    category: str
    severity: Severity
    pattern: re.Pattern[str]
    suggestion: str
    confidence: float
    replacement_builder: ReplacementBuilder
    evidence_kind: EvidenceKind = "rule"


STYLE_RULES = [
    RegexLocalRule(
        rule_id="style_consecutive_punctuation",
        pass_name="style_rule_pass",
        category="style",
        severity="low",
        pattern=re.compile(r"([。！？!?])\1+"),
        suggestion="连续同类句末标点可规范为单个标点。",
        confidence=0.9,
        replacement_builder=lambda match: match.group(1),
    ),
    RegexLocalRule(
        rule_id="style_ascii_comma",
        pass_name="style_rule_pass",
        category="style",
        severity="low",
        pattern=re.compile(r"[\u4e00-\u9fff],[\u4e00-\u9fff]"),
        suggestion="中文语境中出现英文逗号，建议改为中文逗号。",
        confidence=0.92,
        replacement_builder=lambda match: match.group(0).replace(",", "，"),
    ),
    RegexLocalRule(
        rule_id="style_halfwidth_parenthesis",
        pass_name="style_rule_pass",
        category="style",
        severity="low",
        pattern=re.compile(r"[\u4e00-\u9fff]\([^)]+\)"),
        suggestion="中文正文中的半角括号建议改为全角括号。",
        confidence=0.9,
        replacement_builder=lambda match: match.group(0).replace("(", "（").replace(")", "）"),
    ),
]


def run_terminology_rules(source_text: str) -> list[LocalRuleMatch]:
    """Keep the terminology pass callable, but do not emit low-confidence local candidates by default."""
    return []


def run_style_rules(source_text: str) -> list[LocalRuleMatch]:
    """Find all non-overlapping local style-rule matches in project text."""
    matches: list[LocalRuleMatch] = []
    for rule in STYLE_RULES:
        if rule.confidence < HIGH_CONFIDENCE_LOCAL_RULE_MIN:
            continue
        for match in rule.pattern.finditer(source_text):
            replacement = rule.replacement_builder(match)
            if replacement == match.group(0):
                continue
            matches.append(
                LocalRuleMatch(
                    rule_id=rule.rule_id,
                    pass_name=rule.pass_name,
                    category=rule.category,
                    severity=rule.severity,
                    original=match.group(0),
                    replacement=replacement,
                    suggestion=rule.suggestion,
                    start=match.start(),
                    end=match.end(),
                    confidence=rule.confidence,
                    evidence_kind=rule.evidence_kind,
                )
            )
    return matches


def run_consistency_rules(source_text: str, *, source_type: str) -> list[LocalRuleMatch]:
    """Keep the consistency pass callable, but do not emit coarse repeated-number candidates by default."""
    return []
