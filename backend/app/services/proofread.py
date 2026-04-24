from __future__ import annotations

import os

from app.schemas import ProofreadIssue
from app.services.ai_client import proofread_with_ai


async def proofread_text(text: str) -> list[ProofreadIssue]:
    if os.getenv("AI_API_KEY"):
        return await proofread_with_ai(text)

    return build_mock_issues(text)


def build_mock_issues(text: str) -> list[ProofreadIssue]:
    sample = text[: min(len(text), 12)]
    end = len(sample)

    return [
        ProofreadIssue(
            id="mock-issue-1",
            category="style",
            severity="medium",
            original=sample,
            suggestion="请结合上下文检查该表述是否准确、简洁，并确认是否符合出版物体例。",
            comment="本地 mock 审校结果：当前未配置 AI_API_KEY，已返回一条示例审校建议，用于验证 Word 批注闭环。",
            start=0,
            end=end,
        )
    ]
