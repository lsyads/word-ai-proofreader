from __future__ import annotations

import logging
import re
import time
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
        run = project_store.create_run(
            project_id,
            total_chunks=document_map.chunk_count,
            run_settings=request.model_dump(),
        )
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

    def start_failed_chunk_retry(
        self,
        project_id: str,
        run_id: str,
        request: V2RunCreateRequest,
    ) -> V2RunResponse:
        project_store.require_project(project_id)
        latest = project_store.latest_run(project_id)
        if not latest or latest.run_id != run_id:
            raise V2WorkspaceConflict("Only the latest review run can retry failed chunks.")
        if latest.status in {"queued", "running"}:
            raise V2WorkspaceConflict("The review run is still running.")

        failed_indices, _completed_indices, retry_counts = _chunk_state_from_events(
            project_store.list_run_events(project_id, run_id)
        )
        if not failed_indices:
            raise V2WorkspaceConflict("This review run has no failed chunks to retry.")

        if project_store.get_run_settings(project_id, run_id) is None:
            project_store.save_run_settings(project_id, run_id, request.model_dump())
        project_store.update_run(project_id, run_id, status="queued", stage="retry_failed_chunks")
        project_store.update_project_status(project_id, "running")
        project_store.add_run_event(
            project_id,
            run_id,
            "retry_queued",
            {
                "tool_name": "proofread_document_chunk",
                "pass_name": "proofread_pass",
                "retry_chunk_indices": sorted(failed_indices),
                "retry_counts": {
                    str(index): retry_counts.get(index, 0) + 1
                    for index in sorted(failed_indices)
                },
                "message": f"已创建失败分块重试任务，共 {len(failed_indices)} 块。",
            },
        )
        return project_store.require_run(project_id, run_id)

    async def retry_failed_chunks_job(
        self,
        project_id: str,
        run_id: str,
        request: V2RunCreateRequest,
    ) -> None:
        try:
            await self._execute_failed_chunk_retry(project_id, run_id, request)
        except Exception:
            logger.exception("V2.2 failed chunk retry crashed project_id=%s run_id=%s", project_id, run_id)
            raise

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

    async def _execute_failed_chunk_retry(
        self,
        project_id: str,
        run_id: str,
        request: V2RunCreateRequest,
    ) -> None:
        project = project_store.require_project(project_id)
        prepared = self._prepare_source(project)
        stored_settings = project_store.get_run_settings(project_id, run_id)
        retry_request = V2RunCreateRequest.model_validate(stored_settings or request.model_dump())
        failed_indices, completed_indices, retry_counts = _chunk_state_from_events(
            project_store.list_run_events(project_id, run_id)
        )
        retry_chunks = [chunk for chunk in prepared.chunks if chunk.index in failed_indices]
        if not retry_chunks:
            project_store.update_run(project_id, run_id, status="waiting_for_approval", stage="waiting_for_approval")
            project_store.update_project_status(project_id, "waiting_for_approval")
            return

        existing_candidates = project_store.list_candidates(project_id, run_id=run_id)
        existing_keys = {_candidate_key(candidate) for candidate in existing_candidates}
        new_candidates: list[V2CandidateIssue] = []
        project_store.update_run(
            project_id,
            run_id,
            status="running",
            stage="retry_failed_chunks",
            completed_chunks=len(completed_indices),
            failed_chunks=len(failed_indices),
            candidate_count=len(existing_candidates),
        )

        for chunk in retry_chunks:
            retry_count = retry_counts.get(chunk.index, 0) + 1
            started_at = time.monotonic()
            project_store.add_run_event(
                project_id,
                run_id,
                "chunk_retrying",
                _chunk_event_data(
                    chunk_index=chunk.index,
                    total_chunks=len(prepared.chunks),
                    completed_chunks=len(completed_indices),
                    failed_chunks=len(failed_indices),
                    candidate_count=len(existing_candidates) + len(new_candidates),
                    retry_count=retry_count,
                    message=f"正在重新审校第 {chunk.index + 1}/{len(prepared.chunks)} 块。",
                ),
            )
            try:
                issues = await proofread_service.proofread_text_with_context(
                    chunk.text,
                    project.book,
                    session_id=retry_request.session_id,
                    ai_profile_id=retry_request.ai_profile_id,
                    provider_api=retry_request.provider_api,
                    proofread_mode=retry_request.proofread_mode,
                    reasoning_enabled=retry_request.reasoning_enabled,
                    temperature=retry_request.temperature,
                )
            except Exception as exc:
                failed_indices.add(chunk.index)
                completed_indices.discard(chunk.index)
                project_store.update_run(
                    project_id,
                    run_id,
                    status="running",
                    stage="retry_failed_chunks",
                    completed_chunks=len(completed_indices),
                    failed_chunks=len(failed_indices),
                    candidate_count=len(existing_candidates) + len(new_candidates),
                )
                project_store.add_run_event(
                    project_id,
                    run_id,
                    "error",
                    _chunk_event_data(
                        chunk_index=chunk.index,
                        total_chunks=len(prepared.chunks),
                        completed_chunks=len(completed_indices),
                        failed_chunks=len(failed_indices),
                        candidate_count=len(existing_candidates) + len(new_candidates),
                        retry_count=retry_count,
                        elapsed_seconds=_elapsed_seconds(started_at),
                        error_type=type(exc).__name__,
                        message=_safe_event_error(str(exc) or type(exc).__name__),
                    ),
                )
                logger.exception("V2.2 retry chunk failed project_id=%s run_id=%s chunk_index=%s", project_id, run_id, chunk.index)
                continue

            failed_indices.discard(chunk.index)
            completed_indices.add(chunk.index)
            global_issues = chunk_service.globalize_issues(chunk, issues)
            chunk_candidates = [
                _evaluate_candidate(
                    _candidate_from_issue(project.project_id, run_id, issue, prepared.source_text, pass_name="proofread_pass")
                )
                for issue in global_issues
            ]
            chunk_candidates = [
                candidate
                for candidate in _dedupe_candidates(chunk_candidates)
                if _candidate_key(candidate) not in existing_keys
            ]
            for candidate in chunk_candidates:
                existing_keys.add(_candidate_key(candidate))
                self._add_candidate_found_event(candidate)
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
            new_candidates.extend(chunk_candidates)
            if chunk_candidates:
                project_store.save_candidates(chunk_candidates)
            project_store.update_run(
                project_id,
                run_id,
                status="running",
                stage="retry_failed_chunks",
                completed_chunks=len(completed_indices),
                failed_chunks=len(failed_indices),
                candidate_count=len(existing_candidates) + len(new_candidates),
            )
            project_store.add_run_event(
                project_id,
                run_id,
                "tool_completed",
                {
                    "tool_name": "proofread_document_chunk",
                    "pass_name": "proofread_pass",
                    "chunk_index": chunk.index,
                    "current_chunk": chunk.index + 1,
                    "total_chunks": len(prepared.chunks),
                    "completed_chunks": len(completed_indices),
                    "failed_chunks": len(failed_indices),
                    "candidate_count": len(existing_candidates) + len(new_candidates),
                    "issue_count": len(chunk_candidates),
                    "retry_count": retry_count,
                    "elapsed_seconds": _elapsed_seconds(started_at),
                },
            )

        all_candidates = project_store.list_candidates(project_id, run_id=run_id)
        if not completed_indices and failed_indices:
            status = "failed"
            error_message = "All chunks failed to proofread."
        elif all_candidates:
            status = "waiting_for_approval"
            error_message = None
        else:
            status = "succeeded" if not failed_indices else "partial_succeeded"
            error_message = None
        project_store.update_run(
            project_id,
            run_id,
            status=status,
            stage=status,
            completed_chunks=len(completed_indices),
            failed_chunks=len(failed_indices),
            candidate_count=len(all_candidates),
            error_message=error_message,
        )
        if new_candidates:
            project_status = "waiting_for_approval"
        elif project.output_filename:
            project_status = "written"
        elif status == "partial_succeeded":
            project_status = "succeeded"
        else:
            project_status = status
        project_store.update_project_status(project_id, project_status)
        if new_candidates and project.output_filename:
            project_store.mark_project_output_stale(project_id, True)
        project_store.add_run_event(
            project_id,
            run_id,
            "pass_completed",
            {
                "pass_name": "retry_failed_chunks",
                "completed_chunks": len(completed_indices),
                "failed_chunks": len(failed_indices),
                "candidate_count": len(all_candidates),
                "new_candidate_count": len(new_candidates),
            },
        )
        refreshed_project = project_store.require_project(project_id)
        report = report_service.build_review_report(
            project_id=project_id,
            status=refreshed_project.status,
            source_filename=refreshed_project.source_filename,
            book=refreshed_project.book,
            review_goal=refreshed_project.review_goal,
            candidates=all_candidates,
        )
        project_store.save_report(report)
        project_store.add_run_event(project_id, run_id, "report_ready", {"candidate_count": len(all_candidates)})

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
        already_written = project_store.list_candidates(project_id, run_id=latest.run_id, status="written")
        included = [*already_written, *approved]
        if not approved:
            raise V2WorkspaceConflict("No approved candidate issues are available to write back.")

        output_filename = docx_service.build_output_filename(project.source_filename, request.application_mode)
        output_path = project_store.project_output_path(project_id, output_filename)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        summary = docx_service.write_docx_result(
            project.source_bytes,
            [_chunked_issue_from_candidate(candidate) for candidate in included],
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
            included_count=len(included),
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
        key = _candidate_key(candidate)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(candidate)
    return deduped


def _candidate_key(candidate: V2CandidateIssue) -> tuple:
    return (
        candidate.original,
        candidate.replacement,
        candidate.suggestion,
        candidate.global_start,
        candidate.global_end,
        candidate.pass_name,
    )


def _chunk_state_from_events(events: list[Any]) -> tuple[set[int], set[int], dict[int, int]]:
    failed: set[int] = set()
    completed: set[int] = set()
    retry_counts: dict[int, int] = {}
    for event in events:
        data = event.data
        chunk_index = data.get("chunk_index")
        if not isinstance(chunk_index, int):
            continue
        if event.event == "chunk_retrying":
            retry_counts[chunk_index] = max(retry_counts.get(chunk_index, 0), int(data.get("retry_count") or 1))
            continue
        if data.get("tool_name") != "proofread_document_chunk":
            continue
        if event.event == "tool_completed":
            completed.add(chunk_index)
            failed.discard(chunk_index)
        elif event.event == "error":
            failed.add(chunk_index)
            completed.discard(chunk_index)
    return failed, completed, retry_counts


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
    retry_count: int | None = None,
    elapsed_seconds: float | None = None,
    error_type: str | None = None,
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
    if retry_count is not None:
        data["retry_count"] = retry_count
    if elapsed_seconds is not None:
        data["elapsed_seconds"] = elapsed_seconds
    if error_type is not None:
        data["error_type"] = error_type
    return data


def _elapsed_seconds(started_at: float) -> float:
    return round(time.monotonic() - started_at, 3)


def _safe_event_error(value: str) -> str:
    sanitized = value
    for pattern in _SECRET_PATTERNS:
        sanitized = pattern.sub(r"\1[REDACTED]", sanitized)
    if len(sanitized) > _MAX_EVENT_ERROR_LENGTH:
        return sanitized[:_MAX_EVENT_ERROR_LENGTH] + "...[truncated]"
    return sanitized


workspace_runner = AgentWorkspaceRunner()
