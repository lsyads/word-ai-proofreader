from __future__ import annotations

from app.schemas import V2DocumentMapResponse, V2ReviewPlanResponse, V2ReviewPlanStep


def create_review_plan(
    project_id: str,
    run_id: str | None = None,
    *,
    source_type: str = "docx",
    review_goal: str = "",
    document_map: V2DocumentMapResponse | None = None,
) -> V2ReviewPlanResponse:
    """Create the deterministic V2.2 publishing-review plan for one project run."""
    return V2ReviewPlanResponse(
        project_id=project_id,
        run_id=run_id,
        steps=[
            V2ReviewPlanStep(
                step_id="plan_review",
                title="生成审校计划",
                tool_name="create_review_plan",
                description="根据来源类型和文档地图建立本次 Agent run 的执行阶段。",
                status="succeeded" if run_id else "pending",
                reason="V2.2 每次 run 都重新构建文档地图和审校计划。",
            ),
            V2ReviewPlanStep(
                step_id="proofread_pass",
                title="基础语言审校",
                tool_name="proofread_document_chunk",
                description="复用 V1 分块、AI 审校和定位能力，按统一出版审校 prompt 执行。",
                reason="明显错别字、漏字、多字、语病、事实和逻辑风险始终启用；机械校对项默认过滤。",
            ),
            V2ReviewPlanStep(
                step_id="merge_candidates",
                title="候选问题归并",
                tool_name="merge_candidate_issues",
                description="按原文、位置和建议归并重复候选问题。",
            ),
            V2ReviewPlanStep(
                step_id="evaluate_candidates",
                title="二次复核与证据绑定",
                tool_name="evaluate_candidate_issues",
                description="过滤低价值候选，标注置信度、证据类型和人工重点判断理由。",
            ),
            V2ReviewPlanStep(
                step_id="human_approval",
                title="等待编辑确认",
                tool_name="human_approval_queue",
                description="Agent 不静默改正文，等待编辑批准、拒绝或暂缓。",
            ),
        ],
    )
