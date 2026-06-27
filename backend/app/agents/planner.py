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
    normalized_goal = review_goal.lower()
    text_len = document_map.text_len if document_map else 0
    is_book_scope = source_type == "docx" or text_len >= 7000
    wants_terms = any(keyword in normalized_goal for keyword in ("术语", "专名", "人名", "地名", "机构", "一致"))
    wants_style = any(keyword in normalized_goal for keyword in ("体例", "标点", "数字", "格式", "出版", "全半角"))
    wants_consistency = is_book_scope or any(keyword in normalized_goal for keyword in ("全书", "跨章节", "一致", "前后"))

    return V2ReviewPlanResponse(
        project_id=project_id,
        run_id=run_id,
        steps=[
            V2ReviewPlanStep(
                step_id="plan_review",
                title="生成审校计划",
                tool_name="create_review_plan",
                description="根据审校目标、来源类型和文档地图决定本次 Agent run 的专项 pass。",
                status="succeeded" if run_id else "pending",
                reason="V2.2 每次 run 都重新绑定审校目标和文档地图。",
            ),
            V2ReviewPlanStep(
                step_id="proofread_pass",
                title="基础语言审校",
                tool_name="proofread_document_chunk",
                description="复用 V1 分块、AI 审校和定位能力，但使用 V2.2 审校目标上下文 prompt。",
                reason="基础错别字、语病、标点和明确体例问题始终启用。",
            ),
            V2ReviewPlanStep(
                step_id="terminology_pass",
                title="术语与专名一致性",
                tool_name="check_terminology_consistency",
                description="保留术语与专名一致性阶段入口；默认不由本地规则生成低置信人工核查候选。",
                enabled=is_book_scope or wants_terms,
                reason="全书项目或目标提到术语/专名/一致性时保留阶段可观测性，候选主要来自 AI 或编辑记忆。",
            ),
            V2ReviewPlanStep(
                step_id="style_rule_pass",
                title="出版体例规则",
                tool_name="check_style_rules",
                description="检查高置信机械体例规则，例如英文逗号、半角括号和连续同类标点。",
                enabled=wants_style or bool(review_goal),
                reason="审校目标会转化为本项目体例检查约束；本地规则只保留可直接替换的高置信候选。",
            ),
            V2ReviewPlanStep(
                step_id="consistency_pass",
                title="跨章节一致性",
                tool_name="check_cross_chapter_consistency",
                description="保留跨章节一致性阶段入口；默认不由重复数字等粗粒度规则生成低置信候选。",
                enabled=wants_consistency,
                reason="DOCX/长文本或目标提到全书、跨章节、前后一致时保留阶段可观测性，避免默认制造大量人工核查项。",
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
