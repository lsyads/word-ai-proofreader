from __future__ import annotations

import json
import logging
from contextlib import asynccontextmanager
from collections.abc import AsyncIterator
from pathlib import Path
from typing import Any

from fastapi import BackgroundTasks, FastAPI, HTTPException, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse

from app.schemas import (
    AgentRunTraceResponse,
    AIProfileResponse,
    ApplicationMode,
    BookInfo,
    ChunkedProofreadRequest,
    ChunkedProofreadResult,
    DocxProofreadResult,
    ProofreadRequest,
    ProofreadResponse,
    SessionResponse,
    V2ApprovalDecisionRequest,
    V2ApprovalDecisionResponse,
    V2CandidateListResponse,
    V2DocumentMapResponse,
    V2MemoryCreateRequest,
    V2MemoryListResponse,
    V2MarkWrittenRequest,
    V2MarkWrittenResponse,
    V2ProjectListResponse,
    V2ProjectResponse,
    V2ReviewPlanResponse,
    V2ReviewReportResponse,
    V2RunCreateRequest,
    V2RunResponse,
    V2RunTraceResponse,
    V2SelectionProjectCreateRequest,
    V2WritebackRequest,
    V2WritebackResponse,
)
from app.agents import trace as agent_trace
from app.agents import memory as agent_memory
from app.agents.service import AgentOptions, agent_runner
from app.agents.workspace import V2WorkspaceConflict, workspace_runner
import app.services.proofread as proofread_service
from app.services import chunking
from app.services import docx as docx_service
from app.services import docx_tasks as docx_task_service
from app.services import project_store
from app.services import report_service
from app.services import tasks as task_service
from app.services.ai_client import AIClientError
from app.services.ai_profiles import AIProfileError, list_public_ai_profiles
from app.services.sessions import create_session
from app.settings import get_settings


def configure_logging(level_name: str) -> None:
    level = getattr(logging, level_name.upper(), logging.INFO)
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
    )
    logging.getLogger().setLevel(level)
    for handler in logging.getLogger().handlers:
        handler.setLevel(level)
    logging.getLogger("app").setLevel(level)


@asynccontextmanager
async def lifespan(_: FastAPI) -> AsyncIterator[None]:
    docx_task_service.cleanup_expired_results()
    yield


app = FastAPI(title="Word AI Proofreader", version="0.1.0", lifespan=lifespan)
settings = get_settings()
configure_logging(settings.backend_log_level)
logger = logging.getLogger(__name__)

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.backend_cors_origin_list,
    allow_credentials=False,
    allow_methods=["GET", "POST", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/sessions", response_model=SessionResponse)
async def create_ai_session() -> SessionResponse:
    session = create_session()
    logger.info("created AI session session_id=%s", _mask_session_id(session.session_id))
    return session


@app.get("/api/ai-profiles", response_model=list[AIProfileResponse])
async def get_ai_profiles() -> list[AIProfileResponse]:
    try:
        profiles = list_public_ai_profiles()
    except AIProfileError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return [
        AIProfileResponse(
            id=profile.id,
            label=profile.label,
            model=profile.model,
            default_api=profile.default_api,
            supported_apis=list(profile.supported_apis),
            configured=profile.configured,
        )
        for profile in profiles
    ]


@app.post("/api/proofread", response_model=ProofreadResponse)
async def proofread(request: ProofreadRequest) -> ProofreadResponse:
    logger.info(
        "proofread request received text_len=%s session_id=%s ai_profile_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s temperature=%s",
        len(request.text),
        _mask_session_id(request.session_id),
        request.ai_profile_id or "default",
        request.provider_api or "default",
        request.proofread_mode,
        request.reasoning_enabled,
        request.temperature,
    )
    _debug_log_json("proofread request body", request.model_dump())
    try:
        response = await agent_runner.proofread_selection(
            request.text,
            request.book,
            options=AgentOptions(
                session_id=request.session_id,
                ai_profile_id=request.ai_profile_id,
                provider_api=request.provider_api,
                proofread_mode=request.proofread_mode,
                reasoning_enabled=request.reasoning_enabled,
                temperature=request.temperature,
                scope="selection",
                chunk_size=chunking.DEFAULT_CHUNK_SIZE,
            ),
            proofread_callable=proofread_service.proofread_text,
        )
    except AIProfileError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except AIClientError as exc:
        logger.warning(
            "proofread request failed text_len=%s session_id=%s error=%s",
            len(request.text),
            _mask_session_id(request.session_id),
            exc,
            exc_info=True,
        )
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    logger.info(
        "proofread request completed issue_count=%s located_issue_count=%s",
        len(response.issues),
        _count_located_issues(response.issues),
    )
    _debug_log_json("proofread response body", response.model_dump())
    return response


@app.post("/api/proofread/stream")
async def proofread_stream(request: ProofreadRequest) -> StreamingResponse:
    try:
        provider_api = proofread_service.resolve_provider_api(request.provider_api, request.ai_profile_id)
    except AIProfileError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    if provider_api == "chat":
        raise HTTPException(
            status_code=400,
            detail="Chat mode uses /api/proofread with standard Chat Completions, not SSE.",
        )

    logger.info(
        "proofread stream accepted text_len=%s session_id=%s ai_profile_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s temperature=%s",
        len(request.text),
        _mask_session_id(request.session_id),
        request.ai_profile_id or "default",
        provider_api,
        request.proofread_mode,
        request.reasoning_enabled,
        request.temperature,
    )
    _debug_log_json("proofread stream request body", request.model_dump())
    run_id = agent_trace.create_run(
        "selection_stream",
        total_chunks=1,
        metadata={
            "session_id_suffix": request.session_id[-8:] if request.session_id else None,
            "ai_profile_id": request.ai_profile_id or "default",
            "provider_api": provider_api,
            "proofread_mode": request.proofread_mode,
            "reasoning_enabled": request.reasoning_enabled,
            "temperature": request.temperature,
            "scope": "selection",
        },
    )
    return StreamingResponse(
        _proofread_event_stream(request, run_id),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache, no-transform",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@app.post("/api/proofread/chunked", response_model=ChunkedProofreadResult)
async def proofread_chunked(request: ChunkedProofreadRequest) -> ChunkedProofreadResult:
    logger.info(
        "chunked proofread request received text_len=%s scope=%s chunk_size=%s session_id=%s ai_profile_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s temperature=%s",
        len(request.text),
        request.scope,
        request.chunk_size,
        _mask_session_id(request.session_id),
        request.ai_profile_id or "default",
        request.provider_api or "default",
        request.proofread_mode,
        request.reasoning_enabled,
        request.temperature,
    )
    _debug_log_json("chunked proofread request body", request.model_dump())
    try:
        proofread_service.resolve_provider_api(request.provider_api, request.ai_profile_id)
    except AIProfileError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    response = await agent_runner.proofread_chunked_request(
        request,
        proofread_callable=proofread_service.proofread_text,
    )
    logger.info(
        "chunked proofread request completed total_chunks=%s completed_chunks=%s failed_chunks=%s issue_count=%s",
        response.total_chunks,
        response.completed_chunks,
        response.failed_chunks,
        len(response.issues),
    )
    _debug_log_json("chunked proofread response body", response.model_dump())
    return response


@app.get("/api/agent/runs/{run_id}/trace", response_model=AgentRunTraceResponse)
async def get_agent_run_trace(run_id: str) -> AgentRunTraceResponse:
    try:
        return AgentRunTraceResponse.model_validate(agent_trace.get_trace(run_id).model_dump())
    except agent_trace.AgentTraceNotFound as exc:
        raise HTTPException(status_code=404, detail="Agent run trace not found") from exc


@app.post("/api/v2/projects", response_model=V2ProjectResponse)
async def create_v2_project(
    request: Request,
    filename: str = Query(..., min_length=1),
    book: str = Query(..., min_length=1),
    review_goal: str = Query(default="完成全书出版审校，输出候选问题、证据、写回结果和审校报告。", min_length=1),
) -> V2ProjectResponse:
    if Path(filename).suffix.lower() == ".doc":
        raise HTTPException(status_code=400, detail=".doc 是旧二进制格式，请先另存为 .docx 后再上传。")
    if Path(filename).suffix.lower() != ".docx":
        raise HTTPException(status_code=400, detail="V2 审校项目当前仅支持 .docx。")
    try:
        book_info = BookInfo.model_validate(json.loads(book))
    except Exception as exc:
        raise HTTPException(status_code=422, detail="book 参数必须是有效的书籍信息 JSON。") from exc
    content = await request.body()
    if not content:
        raise HTTPException(status_code=400, detail="上传的 .docx 文件为空。")
    try:
        return workspace_runner.create_project(
            source_filename=filename,
            source_bytes=content,
            book=book_info,
            review_goal=review_goal,
        )
    except docx_service.DocxError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/v2/projects/selection", response_model=V2ProjectResponse)
async def create_v2_selection_project(request: V2SelectionProjectCreateRequest) -> V2ProjectResponse:
    return workspace_runner.create_selection_project(
        text=request.text,
        book=request.book,
        review_goal=request.review_goal,
    )


@app.get("/api/v2/projects", response_model=V2ProjectListResponse)
async def list_v2_projects(limit: int = Query(default=20, ge=1, le=100)) -> V2ProjectListResponse:
    return V2ProjectListResponse(projects=project_store.list_projects(limit=limit))


@app.get("/api/v2/projects/{project_id}", response_model=V2ProjectResponse)
async def get_v2_project(project_id: str) -> V2ProjectResponse:
    try:
        return project_store.project_response(project_id)
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 project not found") from exc


@app.get("/api/v2/projects/{project_id}/document-map", response_model=V2DocumentMapResponse)
async def get_v2_document_map(project_id: str) -> V2DocumentMapResponse:
    try:
        return project_store.get_document_map(project_id)
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 document map not found") from exc


@app.get("/api/v2/projects/{project_id}/plan", response_model=V2ReviewPlanResponse)
async def get_v2_review_plan(project_id: str) -> V2ReviewPlanResponse:
    try:
        return project_store.get_review_plan(project_id)
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 review plan not found") from exc


@app.post("/api/v2/projects/{project_id}/runs", response_model=V2RunResponse)
async def create_v2_run(
    project_id: str,
    request: V2RunCreateRequest,
    background_tasks: BackgroundTasks,
) -> V2RunResponse:
    try:
        proofread_service.resolve_provider_api(request.provider_api, request.ai_profile_id)
        run = workspace_runner.start_project_run(project_id, request)
        background_tasks.add_task(workspace_runner.run_project_job, project_id, run.run_id, request)
        return run
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 project not found") from exc
    except AIProfileError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except docx_service.DocxError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/api/v2/projects/{project_id}/runs/{run_id}", response_model=V2RunResponse)
async def get_v2_run(project_id: str, run_id: str) -> V2RunResponse:
    try:
        return project_store.require_run(project_id, run_id)
    except project_store.V2RunNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 run not found") from exc


@app.get("/api/v2/projects/{project_id}/runs/{run_id}/trace", response_model=V2RunTraceResponse)
async def get_v2_run_trace(project_id: str, run_id: str) -> V2RunTraceResponse:
    try:
        run = project_store.require_run(project_id, run_id)
        return V2RunTraceResponse(
            project_id=project_id,
            run_id=run_id,
            status=run.status,
            events=project_store.list_run_events(project_id, run_id),
        )
    except project_store.V2RunNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 run not found") from exc


@app.get("/api/v2/projects/{project_id}/runs/{run_id}/events")
async def v2_run_events(project_id: str, run_id: str) -> StreamingResponse:
    try:
        project_store.require_run(project_id, run_id)
    except project_store.V2RunNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 run not found") from exc
    return StreamingResponse(
        _v2_event_stream(project_store.list_run_events(project_id, run_id)),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache, no-transform",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@app.get("/api/v2/projects/{project_id}/candidates", response_model=V2CandidateListResponse)
async def get_v2_candidates(project_id: str) -> V2CandidateListResponse:
    try:
        project_store.require_project(project_id)
        return V2CandidateListResponse(project_id=project_id, candidates=project_store.list_candidates(project_id))
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 project not found") from exc


@app.post("/api/v2/projects/{project_id}/candidates/decisions", response_model=V2ApprovalDecisionResponse)
async def decide_v2_candidates(project_id: str, request: V2ApprovalDecisionRequest) -> V2ApprovalDecisionResponse:
    try:
        updated_count = project_store.update_candidate_statuses(
            project_id,
            {decision.candidate_id: decision.status for decision in request.decisions},
        )
        _refresh_v2_memory_from_decisions(project_id)
        return V2ApprovalDecisionResponse(
            project_id=project_id,
            updated_count=updated_count,
            candidates=project_store.list_candidates(project_id),
        )
    except project_store.V2CandidateNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 candidate not found") from exc
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 project not found") from exc


@app.get("/api/v2/projects/{project_id}/memory", response_model=V2MemoryListResponse)
async def get_v2_project_memory(project_id: str) -> V2MemoryListResponse:
    try:
        return V2MemoryListResponse(project_id=project_id, memory=project_store.list_memory_items(project_id))
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 project not found") from exc


@app.post("/api/v2/projects/{project_id}/memory", response_model=V2MemoryListResponse)
async def create_v2_project_memory(project_id: str, request: V2MemoryCreateRequest) -> V2MemoryListResponse:
    try:
        saved = project_store.save_memory_item(
            project_id,
            kind=request.kind,
            key=request.key,
            value=request.value,
            source=request.source,
            confidence=request.confidence,
        )
        latest = project_store.latest_run(project_id)
        if latest:
            project_store.add_run_event(
                project_id,
                latest.run_id,
                "memory_updated",
                {"memory_id": saved.memory_id, "kind": saved.kind, "key": saved.key, "source": saved.source},
            )
        return V2MemoryListResponse(project_id=project_id, memory=project_store.list_memory_items(project_id))
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 project not found") from exc


@app.delete("/api/v2/projects/{project_id}/memory/{memory_id}", response_model=V2MemoryListResponse)
async def delete_v2_project_memory(project_id: str, memory_id: str) -> V2MemoryListResponse:
    try:
        project_store.delete_memory_item(project_id, memory_id)
        return V2MemoryListResponse(project_id=project_id, memory=project_store.list_memory_items(project_id))
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 project not found") from exc
    except project_store.V2MemoryNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 memory item not found") from exc


@app.post("/api/v2/projects/{project_id}/candidates/mark-written", response_model=V2MarkWrittenResponse)
async def mark_v2_candidates_written(project_id: str, request: V2MarkWrittenRequest) -> V2MarkWrittenResponse:
    try:
        updated_count = project_store.mark_candidates_written(project_id, request.candidate_ids)
        project_store.update_project_status(project_id, "written")
        project = project_store.require_project(project_id)
        report = report_service.build_review_report(
            project_id=project_id,
            status=project.status,
            source_filename=project.source_filename,
            book=project.book,
            review_goal=project.review_goal,
            candidates=project_store.list_candidates(project_id),
        )
        project_store.save_report(report)
        return V2MarkWrittenResponse(
            project_id=project_id,
            updated_count=updated_count,
            candidates=project_store.list_candidates(project_id),
        )
    except project_store.V2CandidateNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 candidate not found") from exc
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 project not found") from exc


@app.post("/api/v2/projects/{project_id}/writeback", response_model=V2WritebackResponse)
async def writeback_v2_project(project_id: str, request: V2WritebackRequest) -> V2WritebackResponse:
    try:
        response = workspace_runner.write_approved(project_id, request)
        project = project_store.require_project(project_id)
        project_store.add_run_event(
            project_id,
            "writeback",
            "writeback_completed",
            {"output_filename": response.output_filename, "written_count": response.written_count},
        )
        report = report_service.build_review_report(
            project_id=project_id,
            status=project.status,
            source_filename=project.source_filename,
            book=project.book,
            review_goal=project.review_goal,
            candidates=project_store.list_candidates(project_id),
        )
        project_store.save_report(report)
        return response
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 project not found") from exc
    except V2WorkspaceConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc


@app.get("/api/v2/projects/{project_id}/report", response_model=V2ReviewReportResponse)
async def get_v2_report(project_id: str) -> V2ReviewReportResponse:
    try:
        project = project_store.require_project(project_id)
        report = project_store.get_report(project_id)
        if report:
            return report
        return report_service.build_review_report(
            project_id=project_id,
            status=project.status,
            source_filename=project.source_filename,
            book=project.book,
            review_goal=project.review_goal,
            candidates=project_store.list_candidates(project_id),
        )
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 project not found") from exc


@app.get("/api/v2/projects/{project_id}/download")
async def download_v2_project_output(project_id: str) -> FileResponse:
    try:
        output_path, output_filename = project_store.output_download_path(project_id)
    except project_store.V2ProjectNotFound as exc:
        raise HTTPException(status_code=404, detail="V2 project output not found") from exc
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=output_filename,
    )


@app.post("/api/proofread/tasks", response_model=ChunkedProofreadResult)
async def create_proofread_task(request: ChunkedProofreadRequest) -> ChunkedProofreadResult:
    logger.info(
        "proofread task create requested text_len=%s scope=%s chunk_size=%s session_id=%s ai_profile_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s temperature=%s",
        len(request.text),
        request.scope,
        request.chunk_size,
        _mask_session_id(request.session_id),
        request.ai_profile_id or "default",
        request.provider_api or "default",
        request.proofread_mode,
        request.reasoning_enabled,
        request.temperature,
    )
    _debug_log_json("proofread task request body", request.model_dump())
    try:
        proofread_service.resolve_provider_api(request.provider_api, request.ai_profile_id)
    except AIProfileError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return task_service.create_task(request)


@app.get("/api/proofread/tasks/{task_id}", response_model=ChunkedProofreadResult)
async def get_proofread_task(task_id: str) -> ChunkedProofreadResult:
    try:
        return task_service.get_task(task_id)
    except task_service.ProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="Proofread task not found") from exc


@app.delete("/api/proofread/tasks/{task_id}", response_model=ChunkedProofreadResult)
async def cancel_proofread_task(task_id: str) -> ChunkedProofreadResult:
    try:
        return task_service.cancel_task(task_id)
    except task_service.ProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="Proofread task not found") from exc


@app.post("/api/proofread/tasks/{task_id}/retry-current", response_model=ChunkedProofreadResult)
async def retry_current_proofread_chunk(task_id: str) -> ChunkedProofreadResult:
    try:
        return task_service.retry_current_chunk(task_id)
    except task_service.ProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="Proofread task not found") from exc
    except task_service.ProofreadTaskConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc


@app.post("/api/proofread/tasks/{task_id}/retry-failed", response_model=ChunkedProofreadResult)
async def retry_failed_proofread_chunks(task_id: str) -> ChunkedProofreadResult:
    try:
        return task_service.retry_failed_chunks(task_id)
    except task_service.ProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="Proofread task not found") from exc
    except task_service.ProofreadTaskConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc


@app.get("/api/proofread/tasks/{task_id}/events")
async def proofread_task_events(task_id: str) -> StreamingResponse:
    try:
        task_service.get_task(task_id)
    except task_service.ProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="Proofread task not found") from exc

    return StreamingResponse(
        _task_event_stream(task_service.stream_task_events(task_id)),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache, no-transform",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@app.post("/api/proofread/docx/tasks", response_model=DocxProofreadResult)
async def create_docx_proofread_task(
    request: Request,
    filename: str = Query(..., min_length=1),
    book: str = Query(..., min_length=1),
    session_id: str | None = Query(default=None),
    ai_profile_id: str | None = Query(default=None),
    provider_api: proofread_service.ProviderAPI | None = Query(default=None),
    proofread_mode: proofread_service.ProofreadMode = Query(default="fast"),
    reasoning_enabled: bool = Query(default=False),
    temperature: float = Query(default=0.2, ge=0, le=1.5),
    application_mode: ApplicationMode = Query(default="comment"),
    fallback_summary_truncate_enabled: bool = Query(default=True),
) -> DocxProofreadResult:
    if Path(filename).suffix.lower() == ".doc":
        raise HTTPException(status_code=400, detail=".doc 是旧二进制格式，请先另存为 .docx 后再上传。")
    if Path(filename).suffix.lower() != ".docx":
        raise HTTPException(status_code=400, detail="全书文件审校当前仅支持 .docx。")

    try:
        book_info = BookInfo.model_validate(json.loads(book))
    except Exception as exc:
        raise HTTPException(status_code=422, detail="book 参数必须是有效的书籍信息 JSON。") from exc

    content = await request.body()
    if not content:
        raise HTTPException(status_code=400, detail="上传的 .docx 文件为空。")

    logger.info(
        "docx proofread task create requested filename=%s bytes=%s session_id=%s ai_profile_id=%s provider_api=%s proofread_mode=%s reasoning_enabled=%s temperature=%s application_mode=%s fallback_summary_truncate_enabled=%s",
        filename,
        len(content),
        _mask_session_id(session_id),
        ai_profile_id or "default",
        provider_api or "default",
        proofread_mode,
        reasoning_enabled,
        temperature,
        application_mode,
        fallback_summary_truncate_enabled,
    )

    try:
        proofread_service.resolve_provider_api(provider_api, ai_profile_id)
        return docx_task_service.create_task(
            docx_task_service.DocxProofreadRequestData(
                filename=filename,
                content=content,
                book=book_info,
                session_id=session_id,
                ai_profile_id=ai_profile_id,
                provider_api=provider_api,
                proofread_mode=proofread_mode,
                reasoning_enabled=reasoning_enabled,
                temperature=temperature,
                application_mode=application_mode,
                fallback_summary_truncate_enabled=fallback_summary_truncate_enabled,
            )
        )
    except docx_service.DocxError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/api/proofread/docx/tasks/{task_id}", response_model=DocxProofreadResult)
async def get_docx_proofread_task(task_id: str) -> DocxProofreadResult:
    try:
        return docx_task_service.get_task(task_id)
    except docx_task_service.DocxProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="DOCX proofread task not found") from exc


@app.delete("/api/proofread/docx/tasks/{task_id}", response_model=DocxProofreadResult)
async def cancel_docx_proofread_task(task_id: str) -> DocxProofreadResult:
    try:
        return docx_task_service.cancel_task(task_id)
    except docx_task_service.DocxProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="DOCX proofread task not found") from exc


@app.post("/api/proofread/docx/tasks/{task_id}/retry-current", response_model=DocxProofreadResult)
async def retry_current_docx_proofread_chunk(task_id: str) -> DocxProofreadResult:
    try:
        return docx_task_service.retry_current_chunk(task_id)
    except docx_task_service.DocxProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="DOCX proofread task not found") from exc
    except docx_task_service.DocxProofreadTaskConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc


@app.post("/api/proofread/docx/tasks/{task_id}/retry-failed", response_model=DocxProofreadResult)
async def retry_failed_docx_proofread_chunks(task_id: str) -> DocxProofreadResult:
    try:
        return docx_task_service.retry_failed_chunks(task_id)
    except docx_task_service.DocxProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="DOCX proofread task not found") from exc
    except docx_task_service.DocxProofreadTaskConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc


@app.get("/api/proofread/docx/tasks/{task_id}/events")
async def docx_proofread_task_events(task_id: str) -> StreamingResponse:
    try:
        docx_task_service.get_task(task_id)
    except docx_task_service.DocxProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="DOCX proofread task not found") from exc

    return StreamingResponse(
        _task_event_stream(docx_task_service.stream_task_events(task_id)),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache, no-transform",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@app.get("/api/proofread/docx/tasks/{task_id}/download")
async def download_docx_proofread_result(task_id: str) -> FileResponse:
    try:
        output_path, output_filename = docx_task_service.get_download_path(task_id)
    except docx_task_service.DocxProofreadTaskNotFound as exc:
        raise HTTPException(status_code=404, detail="DOCX proofread task not found") from exc
    except docx_task_service.DocxProofreadTaskConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=output_filename,
    )


async def _proofread_event_stream(request: ProofreadRequest, run_id: str) -> AsyncIterator[str]:
    node_id = agent_trace.start_node(run_id, "stream_proofread")
    agent_trace.start_chunk(
        run_id,
        chunk_index=0,
        chunk_start=0,
        chunk_end=len(request.text),
        chunk_len=len(request.text),
    )
    try:
        stream_kwargs = {
            "session_id": request.session_id,
            "provider_api": request.provider_api,
            "proofread_mode": request.proofread_mode,
        }
        if request.ai_profile_id is not None:
            stream_kwargs["ai_profile_id"] = request.ai_profile_id
        if request.reasoning_enabled:
            stream_kwargs["reasoning_enabled"] = True
        if request.temperature != 0.2:
            stream_kwargs["temperature"] = request.temperature

        async for event in proofread_service.stream_proofread_text(
            request.text,
            request.book,
            **stream_kwargs,
        ):
            event.data["run_id"] = run_id
            if event.event == "result":
                issues = event.data.get("issues", [])
                logger.info(
                    "proofread stream result issue_count=%s located_issue_count=%s",
                    len(issues) if isinstance(issues, list) else 0,
                    _count_located_issue_dicts(issues) if isinstance(issues, list) else 0,
                )
                _debug_log_json("proofread stream result body", event.data)
                issue_count = len(issues) if isinstance(issues, list) else 0
                agent_trace.finish_chunk(run_id, 0, status="succeeded", issue_count=issue_count)
                agent_trace.update_run(
                    run_id,
                    status="succeeded",
                    total_chunks=1,
                    completed_chunks=1,
                    failed_chunks=0,
                    issue_count=issue_count,
                )
            yield _format_sse(event.event, event.data)
        agent_trace.finish_node(node_id, "succeeded")
    except (AIClientError, AIProfileError) as exc:
        logger.warning(
            "proofread stream failed text_len=%s session_id=%s error=%s",
            len(request.text),
            _mask_session_id(request.session_id),
            exc,
            exc_info=True,
        )
        agent_trace.finish_chunk(run_id, 0, status="failed", issue_count=0, error_message=str(exc))
        agent_trace.update_run(
            run_id,
            status="failed",
            total_chunks=1,
            completed_chunks=0,
            failed_chunks=1,
            issue_count=0,
            error_message=str(exc),
        )
        agent_trace.finish_node(node_id, "failed", str(exc))
        error_data = {"message": str(exc), "run_id": run_id}
        _debug_log_json("proofread stream error body", error_data)
        yield _format_sse("error", error_data)


async def _task_event_stream(events: AsyncIterator[Any]) -> AsyncIterator[str]:
    async for event in events:
        yield _format_sse(event.event, event.data)


async def _v2_event_stream(events: list[Any]) -> AsyncIterator[str]:
    for event in events:
        yield _format_sse(event.event, event.data)


def _refresh_v2_memory_from_decisions(project_id: str) -> None:
    candidates = project_store.list_candidates(project_id)
    saved_count = 0
    for item in agent_memory.derive_memory_items(candidates):
        if item.source != "editor_decision":
            continue
        project_store.save_memory_item(
            project_id,
            kind=item.kind,
            key=item.key,
            value=item.value,
            source=item.source,
            confidence=item.confidence,
        )
        saved_count += 1
    if saved_count == 0:
        return
    latest = project_store.latest_run(project_id)
    if latest:
        project_store.add_run_event(
            project_id,
            latest.run_id,
            "memory_updated",
            {"source": "editor_decision", "updated_count": saved_count},
        )


def _format_sse(event: str, data: dict[str, Any]) -> str:
    padding = ":" + (" " * 2048)
    return f"{padding}\nevent: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"


def _mask_session_id(session_id: str | None) -> str:
    if not session_id:
        return "none"

    return f"...{session_id[-8:]}"


def _count_located_issues(issues: list[Any]) -> int:
    return sum(
        1
        for issue in issues
        if getattr(issue, "start", None) is not None and getattr(issue, "end", None) is not None
    )


def _count_located_issue_dicts(issues: list[Any]) -> int:
    return sum(
        1
        for issue in issues
        if isinstance(issue, dict) and issue.get("start") is not None and issue.get("end") is not None
    )


def _debug_log_json(message: str, payload: Any) -> None:
    if logger.isEnabledFor(logging.DEBUG):
        logger.debug("%s: %s", message, json.dumps(payload, ensure_ascii=False, default=str))
