from __future__ import annotations

import logging
import re
import uuid
from dataclasses import dataclass
from typing import Any

from app.agents import memory, planner
from app.schemas import (
    BookInfo,
    ChunkedProofreadIssue,
    ProofreadIssue,
    V2CandidateIssue,
    V2DocumentMapResponse,
    V2ProjectResponse,
    V2RunCreateRequest,
    V2RunResponse,
    V2WritebackRequest,
    V2WritebackResponse,
)
from app.services import chunk_service, chunking, document_map_service, project_store, report_service
from app.services import docx as docx_service
from app.services import proofread as proofread_service

logger = logging.getLogger(__name__)
_SECRET_PATTERNS = (
    re.compile(r"(Authorization\s*[:=]\s*Bearer\s+)[^\s,;}]+", re.IGNORECASE),
    re.compile(r"(Bearer\s+)[A-Za-z0-9._\-]+", re.IGNORECASE),
    re.compile(r"(api[_-]?key\s*[:=]\s*)[^\s,;}]+", re.IGNORECASE),
)
_MAX_EVENT_ERROR_LENGTH = 1000


class V2WorkspaceConflict(RuntimeError):
    """Raised when a V2 workspace action is not valid for the current project state."""


@dataclass(frozen=True)
class PreparedSource:
    source_text: str
    chunks: list
    document_map: V2DocumentMapResponse


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
        project_store.save_review_plan(
            planner.create_review_plan(
                project.project_id,
                source_type="docx",
                document_map=document_map,
            )
        )
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
        project_store.save_review_plan(
            planner.create_review_plan(
                project.project_id,
                source_type="selection",
                document_map=document_map,
            )
        )
        return project_store.project_response(project.project_id)

    def start_project_run(self, project_id: str, request: V2RunCreateRequest) -> V2RunResponse:
        project = project_store.require_project_summary(project_id)
        active = project_store.active_run(project_id)
        if active:
            raise V2WorkspaceConflict("This project already has a queued or running review run.")
        document_map = project_store.get_document_map(project_id)
        run = project_store.create_run(project_id, total_chunks=document_map.chunk_count)
        plan = planner.create_review_plan(
            project_id,
            run.run_id,
            source_type=project.source_type,
            document_map=document_map,
        )
        project_store.save_review_plan(plan)
        project_store.update_project_status(project_id, "running")
        project_store.add_run_event(
            project_id,
            run.run_id,
            "plan_created",
            {
                "step_count": len(plan.steps),
                "enabled_steps": [step.step_id for step in plan.steps if step.enabled],
                "message": "V2.2 审校计划已生成。",
            },
        )
        project_store.update_run(project_id, run.run_id, status="queued", stage="queued")
        return project_store.require_run(project_id, run.run_id)

    async def run_project_job(self, project_id: str, run_id: str, request: V2RunCreateRequest) -> None:
        try:
            await self._execute_project_run(project_id, run_id, request)
        except Exception:
            logger.exception("V2.2 project run failed project_id=%s run_id=%s", project_id, run_id)
            raise

    async def run_project(self, project_id: str, request: V2RunCreateRequest) -> V2RunResponse:
        run = self.start_project_run(project_id, request)
        await self._execute_project_run(project_id, run.run_id, request)
        return project_store.require_run(project_id, run.run_id)

    async def _execute_project_run(self, project_id: str, run_id: str, request: V2RunCreateRequest) -> None:
        project = project_store.require_project(project_id)
        prepared = self._prepare_source(project)
        candidates: list[V2CandidateIssue] = []
        completed_chunks = 0
        failed_chunks = 0
        project_store.update_run(project_id, run_id, status="running", stage="proofread_pass")

        try:
            candidates, completed_chunks, failed_chunks = await self._run_proofread_pass(
                project=project,
                run_id=run_id,
                request=request,
                prepared=prepared,
            )
            project_store.update_run(project_id, run_id, status="running", stage="merge_candidates")
            project_store.add_run_event(project_id, run_id, "pass_started", {"pass_name": "merge_candidates"})
            before_merge = len(candidates)
            candidates = _dedupe_candidates(candidates)
            project_store.add_run_event(
                project_id,
                run_id,
                "candidate_merged",
                {"before_count": before_merge, "after_count": len(candidates)},
            )
            project_store.add_run_event(project_id, run_id, "pass_completed", {"pass_name": "merge_candidates"})

            project_store.update_run(project_id, run_id, status="running", stage="evaluate_candidates")
            project_store.add_run_event(project_id, run_id, "pass_started", {"pass_name": "evaluate_candidates"})
            candidates = [_evaluate_candidate(candidate) for candidate in candidates]
            for candidate in candidates:
                project_store.add_run_event(
                    project_id,
                    run_id,
                    "candidate_evaluated",
                    {
                        "candidate_id": candidate.candidate_id,
                        "pass_name": candidate.pass_name,
                        "confidence": candidate.confidence,
                        "needs_human_review": candidate.needs_human_review,
                    },
                )
            project_store.add_run_event(
                project_id,
                run_id,
                "pass_completed",
                {"pass_name": "evaluate_candidates", "candidate_count": len(candidates)},
            )

            project_store.save_candidates(candidates)
            self._save_safe_memory(project_id, run_id, candidates)
            if completed_chunks == 0:
                status = "failed"
            elif candidates:
                status = "waiting_for_approval"
            else:
                status = "succeeded"
            error_message = "All chunks failed to proofread." if status == "failed" else None
            project_store.update_run(
                project_id,
                run_id,
                status=status,
                stage=status,
                completed_chunks=completed_chunks,
                failed_chunks=failed_chunks,
                candidate_count=len(candidates),
                error_message=error_message,
            )
            project_store.update_project_status(project_id, status)
            if status == "waiting_for_approval":
                completion_event = "waiting_for_approval"
                completion_message = "候选问题已进入编辑确认队列。"
            elif status == "succeeded":
                completion_event = "review_completed"
                completion_message = "审校完成，未发现需要确认的问题。"
            else:
                completion_event = "error"
                completion_message = error_message
            project_store.add_run_event(
                project_id,
                run_id,
                completion_event,
                {"candidate_count": len(candidates), "message": completion_message},
            )
            refreshed_project = project_store.require_project(project_id)
            report = report_service.build_review_report(
                project_id=project_id,
                status=refreshed_project.status,
                source_filename=refreshed_project.source_filename,
                book=refreshed_project.book,
                review_goal=refreshed_project.review_goal,
                candidates=project_store.list_candidates(project_id, run_id=run_id),
            )
            project_store.save_report(report)
            project_store.add_run_event(project_id, run_id, "report_ready", {"candidate_count": len(candidates)})
        except Exception as exc:
            project_store.update_run(project_id, run_id, status="failed", stage="failed", error_message=str(exc))
            project_store.update_project_status(project_id, "failed")
            project_store.add_run_event(project_id, run_id, "error", {"message": str(exc)})
            raise

    async def _run_proofread_pass(
        self,
        *,
        project: project_store.StoredProject,
        run_id: str,
        request: V2RunCreateRequest,
        prepared: PreparedSource,
    ) -> tuple[list[V2CandidateIssue], int, int]:
        candidates: list[V2CandidateIssue] = []
        completed_chunks = 0
        failed_chunks = 0
        project_store.add_run_event(
            project.project_id,
            run_id,
            "pass_started",
            {"pass_name": "proofread_pass", "chunk_count": len(prepared.chunks)},
        )
        for chunk in prepared.chunks:
            project_store.update_run(
                project.project_id,
                run_id,
                status="running",
                stage="proofread_pass",
                completed_chunks=completed_chunks,
                failed_chunks=failed_chunks,
                candidate_count=len(candidates),
            )
            project_store.add_run_event(
                project.project_id,
                run_id,
                "tool_started",
                _chunk_event_data(
                    chunk_index=chunk.index,
                    total_chunks=len(prepared.chunks),
                    completed_chunks=completed_chunks,
                    failed_chunks=failed_chunks,
                    candidate_count=len(candidates),
                ),
            )
            try:
                issues = await proofread_service.proofread_text_with_context(
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
                project_store.update_run(
                    project.project_id,
                    run_id,
                    status="running",
                    stage="proofread_pass",
                    completed_chunks=completed_chunks,
                    failed_chunks=failed_chunks,
                    candidate_count=len(candidates),
                )
                project_store.add_run_event(
                    project.project_id,
                    run_id,
                    "error",
                    _chunk_event_data(
                        chunk_index=chunk.index,
                        total_chunks=len(prepared.chunks),
                        completed_chunks=completed_chunks,
                        failed_chunks=failed_chunks,
                        candidate_count=len(candidates),
                        message=_safe_event_error(str(exc)),
                    ),
                )
                logger.exception("V2.2 chunk failed project_id=%s run_id=%s chunk_index=%s", project.project_id, run_id, chunk.index)
                continue

            completed_chunks += 1
            global_issues = chunk_service.globalize_issues(chunk, issues)
            chunk_candidates = [
                _candidate_from_issue(project.project_id, run_id, issue, prepared.source_text, pass_name="proofread_pass")
                for issue in global_issues
            ]
            candidates.extend(chunk_candidates)
            for candidate in chunk_candidates:
                self._add_candidate_found_event(candidate)
            project_store.update_run(
                project.project_id,
                run_id,
                status="running",
                stage="proofread_pass",
                completed_chunks=completed_chunks,
                failed_chunks=failed_chunks,
                candidate_count=len(candidates),
            )
            project_store.add_run_event(
                project.project_id,
                run_id,
                "tool_completed",
                {
                    "tool_name": "proofread_document_chunk",
                    "pass_name": "proofread_pass",
                    "chunk_index": chunk.index,
                    "current_chunk": chunk.index + 1,
                    "total_chunks": len(prepared.chunks),
                    "completed_chunks": completed_chunks,
                    "failed_chunks": failed_chunks,
                    "candidate_count": len(candidates),
                    "issue_count": len(chunk_candidates),
                },
            )
        project_store.add_run_event(
            project.project_id,
            run_id,
            "pass_completed",
            {
                "pass_name": "proofread_pass",
                "completed_chunks": completed_chunks,
                "failed_chunks": failed_chunks,
                "candidate_count": len(candidates),
            },
        )
        return candidates, completed_chunks, failed_chunks

    def _prepare_source(self, project: project_store.StoredProject) -> PreparedSource:
        if project.source_type == "selection":
            source_text = project.source_bytes.decode("utf-8")
            chunks = chunking.split_text_into_chunks(source_text, "selection", chunking.DEFAULT_CHUNK_SIZE)
            document_map = project_store.get_document_map(project.project_id)
            return PreparedSource(source_text=source_text, chunks=chunks, document_map=document_map)

        document = docx_service.parse_docx(project.source_bytes)
        source_text = document.text
        chunks = docx_service.split_docx_into_chunks(document)
        document_map = project_store.get_document_map(project.project_id)
        return PreparedSource(source_text=source_text, chunks=chunks, document_map=document_map)

    def _add_candidate_found_event(self, candidate: V2CandidateIssue) -> None:
        project_store.add_run_event(
            candidate.project_id,
            candidate.run_id,
            "candidate_found",
            {
                "candidate_id": candidate.candidate_id,
                "chunk_index": candidate.chunk_index,
                "category": candidate.category,
                "severity": candidate.severity,
                "pass_name": candidate.pass_name,
                "confidence": candidate.confidence,
            },
        )

    def _save_safe_memory(self, project_id: str, run_id: str, candidates: list[V2CandidateIssue]) -> None:
        items = memory.derive_memory_items(candidates)
        for item in items:
            saved = project_store.save_memory_item(
                project_id,
                kind=item.kind,
                key=item.key,
                value=item.value,
                source=item.source,
                confidence=item.confidence,
            )
            project_store.add_run_event(
                project_id,
                run_id,
                "memory_updated",
                {"memory_id": saved.memory_id, "kind": saved.kind, "key": saved.key, "source": saved.source},
            )

    def write_approved(self, project_id: str, request: V2WritebackRequest) -> V2WritebackResponse:
        project = project_store.require_project(project_id)
        if project.source_type == "selection":
            raise V2WorkspaceConflict("Selection project writeback must be completed by the Word add-in.")
        latest = project_store.latest_run(project_id)
        if not latest:
            raise V2WorkspaceConflict("No review run is available to write back.")
        approved = project_store.list_candidates(project_id, run_id=latest.run_id, status="approved")
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
        project_store.mark_candidates_written(
            project_id,
            [candidate.candidate_id for candidate in approved],
            run_id=latest.run_id,
        )
        refreshed_project = project_store.require_project(project_id)
        report = report_service.build_review_report(
            project_id=project_id,
            status=refreshed_project.status,
            source_filename=refreshed_project.source_filename,
            book=refreshed_project.book,
            review_goal=refreshed_project.review_goal,
            candidates=project_store.list_candidates(project_id, run_id=latest.run_id),
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
    *,
    pass_name: str,
    confidence: float = 0.72,
    evidence_kind: str | None = None,
    rule_id: str | None = None,
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
        self_check="已由 V2.2 evaluator 绑定位置和证据，等待编辑确认。",
        pass_name=pass_name,
        confidence=confidence,
        evidence_kind=evidence_kind or ("locator" if issue.locator else "context"),
        rule_id=rule_id,
        needs_human_review=True,
        evaluation_note=None,
        created_at=now,
        updated_at=now,
    )


def _evaluate_candidate(candidate: V2CandidateIssue) -> V2CandidateIssue:
    confidence = candidate.confidence
    if candidate.locator:
        confidence = min(1, confidence + 0.08)
    if candidate.replacement is None:
        confidence = max(0.45, confidence - 0.08)
    if len(candidate.original.strip()) <= 1:
        confidence = max(0.35, confidence - 0.2)
    needs_human_review = candidate.replacement is None or candidate.severity == "high" or confidence < 0.7
    note = "证据充分，可进入编辑确认。"
    if needs_human_review:
        note = "建议人工重点判断：候选涉及核查、无直接替换或置信度偏低。"
    return candidate.model_copy(
        update={
            "confidence": round(confidence, 2),
            "needs_human_review": needs_human_review,
            "evaluation_note": note,
            "self_check": f"V2.2 复核：{note}",
            "updated_at": project_store._now_iso(),
        }
    )


def _dedupe_candidates(candidates: list[V2CandidateIssue]) -> list[V2CandidateIssue]:
    seen: set[tuple] = set()
    deduped: list[V2CandidateIssue] = []
    for candidate in candidates:
        key = (candidate.original, candidate.replacement, candidate.suggestion, candidate.global_start, candidate.global_end, candidate.pass_name)
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


def _chunk_event_data(
    *,
    chunk_index: int,
    total_chunks: int,
    completed_chunks: int,
    failed_chunks: int,
    candidate_count: int,
    message: str | None = None,
) -> dict[str, Any]:
    data: dict[str, Any] = {
        "tool_name": "proofread_document_chunk",
        "pass_name": "proofread_pass",
        "chunk_index": chunk_index,
        "current_chunk": chunk_index + 1,
        "total_chunks": total_chunks,
        "completed_chunks": completed_chunks,
        "failed_chunks": failed_chunks,
        "candidate_count": candidate_count,
    }
    if message is not None:
        data["message"] = message
    return data


def _safe_event_error(value: str) -> str:
    sanitized = value
    for pattern in _SECRET_PATTERNS:
        sanitized = pattern.sub(r"\1[REDACTED]", sanitized)
    if len(sanitized) > _MAX_EVENT_ERROR_LENGTH:
        return sanitized[:_MAX_EVENT_ERROR_LENGTH] + "...[truncated]"
    return sanitized


workspace_runner = AgentWorkspaceRunner()
