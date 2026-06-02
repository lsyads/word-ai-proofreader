from __future__ import annotations

from app.schemas import V2ReviewPlanResponse, V2ReviewPlanStep


def create_review_plan(project_id: str, run_id: str | None = None) -> V2ReviewPlanResponse:
    """Create the fixed V2 publishing-review plan used by the first workspace Agent."""
    return V2ReviewPlanResponse(
        project_id=project_id,
        run_id=run_id,
        steps=[
            V2ReviewPlanStep(
                step_id="document_map",
                title="建立文档地图",
                tool_name="build_document_map",
                description="抽取章节、块和分块摘要，建立可追踪位置映射。",
            ),
            V2ReviewPlanStep(
                step_id="proofread_chunks",
                title="分块审校",
                tool_name="proofread_document_chunk",
                description="复用 V1 AI 审校和定位能力，逐块生成候选问题。",
            ),
            V2ReviewPlanStep(
                step_id="merge_candidates",
                title="候选问题归并",
                tool_name="merge_candidate_issues",
                description="按原文、位置和建议归并重复候选问题。",
            ),
            V2ReviewPlanStep(
                step_id="evaluate_candidates",
                title="自检与证据绑定",
                tool_name="evaluate_candidate_issues",
                description="给候选问题补充证据和自检说明，进入编辑确认队列。",
            ),
            V2ReviewPlanStep(
                step_id="human_approval",
                title="等待编辑确认",
                tool_name="human_approval_queue",
                description="Agent 不静默改正文，等待编辑批准、拒绝或暂缓。",
            ),
        ],
    )
