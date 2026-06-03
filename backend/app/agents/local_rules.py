from __future__ import annotations

import re
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
    evidence_kind: EvidenceKind = "rule"


TERMINOLOGY_VARIANT_PAIRS = [
    ("AI", "人工智能"),
    ("责任编辑", "责编"),
    ("DOCX", "docx"),
]

STYLE_RULES = [
    RegexLocalRule(
        rule_id="style_consecutive_punctuation",
        pass_name="style_rule_pass",
        category="style",
        severity="low",
        pattern=re.compile(r"[。！？!?]{2,}"),
        suggestion="连续标点可能不符合出版体例，请核对。",
        confidence=0.74,
    ),
    RegexLocalRule(
        rule_id="style_ascii_comma",
        pass_name="style_rule_pass",
        category="style",
        severity="low",
        pattern=re.compile(r"[\u4e00-\u9fff],[\u4e00-\u9fff]"),
        suggestion="中文语境中出现英文逗号，请确认是否应改为中文逗号。",
        confidence=0.74,
    ),
    RegexLocalRule(
        rule_id="style_halfwidth_parenthesis",
        pass_name="style_rule_pass",
        category="style",
        severity="low",
        pattern=re.compile(r"[\u4e00-\u9fff]\([^)]+\)"),
        suggestion="中文正文中的半角括号可能不符合体例，请核对。",
        confidence=0.74,
    ),
]


def run_terminology_rules(source_text: str) -> list[LocalRuleMatch]:
    """Find configured terminology variants when both forms appear in the same project text."""
    matches: list[LocalRuleMatch] = []
    for left, right in TERMINOLOGY_VARIANT_PAIRS:
        if left not in source_text or right not in source_text:
            continue
        for start in _find_all(source_text, right):
            matches.append(
                LocalRuleMatch(
                    rule_id="terminology_variant_pair",
                    pass_name="terminology_pass",
                    category="consistency",
                    severity="medium",
                    original=right,
                    suggestion=f"发现“{left}”与“{right}”并用，请确认本书术语或称谓是否统一。",
                    start=start,
                    end=start + len(right),
                    confidence=0.68,
                    evidence_kind="rule",
                )
            )
    return matches


def run_style_rules(source_text: str) -> list[LocalRuleMatch]:
    """Find all non-overlapping local style-rule matches in project text."""
    matches: list[LocalRuleMatch] = []
    for rule in STYLE_RULES:
        for match in rule.pattern.finditer(source_text):
            matches.append(
                LocalRuleMatch(
                    rule_id=rule.rule_id,
                    pass_name=rule.pass_name,
                    category=rule.category,
                    severity=rule.severity,
                    original=match.group(0),
                    suggestion=rule.suggestion,
                    start=match.start(),
                    end=match.end(),
                    confidence=rule.confidence,
                    evidence_kind=rule.evidence_kind,
                )
            )
    return matches


def run_consistency_rules(source_text: str, *, source_type: str) -> list[LocalRuleMatch]:
    """Find coarse cross-document numeric consistency risks for DOCX projects."""
    if source_type != "docx":
        return []

    numeric_tokens = re.findall(r"\d+(?:\.\d+)?(?:年|月|日|%|％|页|章|节)?", source_text)
    if len(numeric_tokens) < 3:
        return []

    repeated_numbers = sorted({token for token in numeric_tokens if numeric_tokens.count(token) > 1})
    matches: list[LocalRuleMatch] = []
    for token in repeated_numbers:
        start = source_text.find(token)
        matches.append(
            LocalRuleMatch(
                rule_id="cross_chapter_numeric_consistency",
                pass_name="consistency_pass",
                category="consistency",
                severity="medium",
                original=token,
                suggestion="全书出现多个数字/时间表达，请结合上下文核对统计口径、单位和前后一致性。",
                start=start,
                end=start + len(token),
                confidence=0.62,
                evidence_kind="document_map",
            )
        )
    return matches


def _find_all(source_text: str, needle: str) -> list[int]:
    starts: list[int] = []
    cursor = 0
    while True:
        start = source_text.find(needle, cursor)
        if start < 0:
            return starts
        starts.append(start)
        cursor = start + len(needle)
