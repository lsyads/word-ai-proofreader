from __future__ import annotations

import logging
import uuid

from app.agents import planner
from app.schemas import (
    BookInfo,
    ChunkedProofreadIssue,
    V2CandidateIssue,
    V2ProjectResponse,
    V2RunCreateRequest,
    V2RunResponse,
    V2WritebackRequest,
    V2WritebackResponse,
)
from app.services import document_map_service, project_store, report_service
from app.services import docx as docx_service
from app.services import proofread as proofread_service
from app.services import chunk_service
from app.services import chunking

logger = logging.getLogger(__name__)


class V2WorkspaceConflict(RuntimeError):
    """Raised when a V2 workspace action is not valid for the current project state."""


class AgentWorkspaceRunner:
    def create_project(
        self,
        *,
        source_filename: str,
        source_bytes: bytes,
        book: BookInfo,
        review_goal: str,
    ) -> V2ProjectResponse:
        project = project_store.create_project(
            source_filename=source_filename,
            source_bytes=source_bytes,
            book=book,
            review_goal=review_goal,
            source_type="docx",
        )
        document_map = document_map_service.build_document_map(project.project_id, source_bytes)
        project_store.save_document_map(document_map)
        project_store.save_review_plan(planner.create_review_plan(project.project_id))
        return project_store.project_response(project.project_id)

    def create_selection_project(
        self,
        *,
        text: str,
        book: BookInfo,
        review_goal: str,
    ) -> V2ProjectResponse:
        text_preview = _preview(text)
        project = project_store.create_project(
            source_type="selection",
            source_filename="当前选区",
            source_bytes=text.encode("utf-8"),
            book=book,
            review_goal=review_goal,
            text_preview=text_preview,
        )
        document_map = document_map_service.build_selection_document_map(project.project_id, text)
        project_store.save_document_map(document_map)
        project_store.save_review_plan(planner.create_review_plan(project.project_id))
        return project_store.project_response(project.project_id)

    async def run_project(self, project_id: str, request: V2RunCreateRequest) -> V2RunResponse:
        project = project_store.require_project(project_id)
        if project.source_type == "selection":
            source_text = project.source_bytes.decode("utf-8")
            chunks = chunking.split_text_into_chunks(source_text, "selection", chunking.DEFAULT_CHUNK_SIZE)
        else:
            document = docx_service.parse_docx(project.source_bytes)
            source_text = document.text
            chunks = docx_service.split_docx_into_chunks(document)
        run = project_store.create_run(project_id, total_chunks=len(chunks))
        plan = planner.create_review_plan(project_id, run.run_id)
        project_store.save_review_plan(plan)
        project_store.update_project_status(project_id, "running")
        project_store.update_run(project_id, run.run_id, status="running", stage="plan_created")
        project_store.add_run_event(
            project_id,
            run.run_id,
            "plan_created",
            {"step_count": len(plan.steps), "message": "V2 审校计划已生成。"},
        )

        candidates: list[V2CandidateIssue] = []
        completed_chunks = 0
        failed_chunks = 0
        try:
            for chunk in chunks:
                project_store.update_run(
                    project_id,
                    run.run_id,
                    status="running",
                    stage="proofread_chunks",
                    completed_chunks=completed_chunks,
                    failed_chunks=failed_chunks,
                    candidate_count=len(candidates),
                )
                project_store.add_run_event(
                    project_id,
                    run.run_id,
                    "tool_started",
                    {"tool_name": "proofread_document_chunk", "chunk_index": chunk.index},
                )
                try:
                    issues = await proofread_service.proofread_text(
                        chunk.text,
                        project.book,
                        session_id=request.session_id,
                        ai_profile_id=request.ai_profile_id,
                        provider_api=request.provider_api,
                        proofread_mode=request.proofread_mode,
                        reasoning_enabled=request.reasoning_enabled,
                        temperature=request.temperature,
                    )
                except Exception as exc:
                    failed_chunks += 1
                    project_store.add_run_event(
                        project_id,
                        run.run_id,
                        "error",
                        {
                            "tool_name": "proofread_document_chunk",
                            "chunk_index": chunk.index,
                            "message": str(exc),
                        },
                    )
                    logger.exception("V2 project chunk failed project_id=%s run_id=%s chunk_index=%s", project_id, run.run_id, chunk.index)
                    continue

                completed_chunks += 1
                global_issues = chunk_service.globalize_issues(chunk, issues)
                chunk_candidates = [
                    _candidate_from_issue(project_id, run.run_id, issue, source_text)
                    for issue in global_issues
                ]
                candidates.extend(chunk_candidates)
                for candidate in chunk_candidates:
                    project_store.add_run_event(
                        project_id,
                        run.run_id,
                        "candidate_found",
                        {
                            "candidate_id": candidate.candidate_id,
                            "chunk_index": candidate.chunk_index,
                            "category": candidate.category,
                            "severity": candidate.severity,
                        },
                    )
                project_store.add_run_event(
                    project_id,
                    run.run_id,
                    "tool_completed",
                    {
                        "tool_name": "proofread_document_chunk",
                        "chunk_index": chunk.index,
                        "issue_count": len(chunk_candidates),
                    },
                )

            candidates = _dedupe_candidates(candidates)
            project_store.save_candidates(candidates)
            status = "waiting_for_approval" if completed_chunks > 0 else "failed"
            error_message = "All chunks failed to proofread." if status == "failed" else None
            project_store.update_run(
                project_id,
                run.run_id,
                status=status,
                stage=status,
                completed_chunks=completed_chunks,
                failed_chunks=failed_chunks,
                candidate_count=len(candidates),
                error_message=error_message,
            )
            project_store.update_project_status(project_id, status)
            project_store.add_run_event(
                project_id,
                run.run_id,
                "waiting_for_approval" if status == "waiting_for_approval" else "error",
                {"candidate_count": len(candidates), "message": "候选问题已进入编辑确认队列。" if status == "waiting_for_approval" else error_message},
            )
        except Exception as exc:
            project_store.update_run(project_id, run.run_id, status="failed", stage="failed", error_message=str(exc))
            project_store.update_project_status(project_id, "failed")
            project_store.add_run_event(project_id, run.run_id, "error", {"message": str(exc)})
            raise

        return project_store.require_run(project_id, run.run_id)

    def write_approved(self, project_id: str, request: V2WritebackRequest) -> V2WritebackResponse:
        project = project_store.require_project(project_id)
        if project.source_type == "selection":
            raise V2WorkspaceConflict("Selection project writeback must be completed by the Word add-in.")
        approved = [candidate for candidate in project_store.list_candidates(project_id) if candidate.status == "approved"]
        if not approved:
            raise V2WorkspaceConflict("No approved candidate issues are available to write back.")

        output_filename = docx_service.build_output_filename(project.source_filename, request.application_mode)
        output_path = project_store.project_output_path(project_id, output_filename)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        summary = docx_service.write_docx_result(
            project.source_bytes,
            [_chunked_issue_from_candidate(candidate) for candidate in approved],
            request.application_mode,
            output_path,
            fallback_summary_truncate_enabled=request.fallback_summary_truncate_enabled,
        )
        project_store.save_project_output(project_id, output_filename=output_filename, output_path=output_path)
        project_store.mark_candidates_written(project_id, [candidate.candidate_id for candidate in approved])
        refreshed_project = project_store.require_project(project_id)
        report = report_service.build_review_report(
            project_id=project_id,
            status=refreshed_project.status,
            source_filename=refreshed_project.source_filename,
            book=refreshed_project.book,
            review_goal=refreshed_project.review_goal,
            candidates=project_store.list_candidates(project_id),
        )
        project_store.save_report(report)
        return V2WritebackResponse(
            project_id=project_id,
            output_filename=output_filename,
            download_url=f"/api/v2/projects/{project_id}/download",
            comment_count=summary.comment_count,
            revision_count=summary.revision_count,
            fallback_count=summary.fallback_count,
            failed_count=summary.failed_count,
            written_count=len(approved),
        )


def _candidate_from_issue(
    project_id: str,
    run_id: str,
    issue: ChunkedProofreadIssue,
    document_text: str,
) -> V2CandidateIssue:
    now = project_store._now_iso()
    evidence = _evidence_for_issue(issue, document_text)
    return V2CandidateIssue(
        candidate_id=f"candidate_{uuid.uuid4().hex}",
        project_id=project_id,
        run_id=run_id,
        status="pending",
        category=issue.category,
        severity=issue.severity,
        original=issue.original,
        replacement=issue.replacement,
        suggestion=issue.suggestion,
        evidence=evidence,
        chunk_index=issue.chunk_index,
        global_start=issue.global_start,
        global_end=issue.global_end,
        locator=issue.locator,
        self_check="已由 V2 evaluator 绑定文档位置和上下文证据，等待编辑确认。",
        created_at=now,
        updated_at=now,
    )


def _dedupe_candidates(candidates: list[V2CandidateIssue]) -> list[V2CandidateIssue]:
    seen: set[tuple] = set()
    deduped: list[V2CandidateIssue] = []
    for candidate in candidates:
        key = (candidate.original, candidate.replacement, candidate.suggestion, candidate.global_start, candidate.global_end)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(candidate)
    return deduped


def _evidence_for_issue(issue: ChunkedProofreadIssue, document_text: str) -> str:
    if issue.global_start is None or issue.global_end is None:
        return issue.original[:120]
    start = max(0, issue.global_start - 30)
    end = min(len(document_text), issue.global_end + 30)
    return document_text[start:end]


def _chunked_issue_from_candidate(candidate: V2CandidateIssue) -> ChunkedProofreadIssue:
    return ChunkedProofreadIssue(
        id=candidate.candidate_id,
        category=candidate.category,
        severity=candidate.severity,
        original=candidate.original,
        replacement=candidate.replacement,
        suggestion=candidate.suggestion,
        start=None,
        end=None,
        locator=candidate.locator,
        chunk_index=candidate.chunk_index,
        global_start=candidate.global_start,
        global_end=candidate.global_end,
    )


def _preview(text: str) -> str:
    normalized = " ".join(text.split())
    return normalized[:120] + ("..." if len(normalized) > 120 else "")


workspace_runner = AgentWorkspaceRunner()
