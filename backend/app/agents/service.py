from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from app.agents import trace
from app.agents.graph import build_proofread_graph
from app.agents.state import AgentFlow, AgentOptions, ProofreadChunkCallable
from app.schemas import (
    BookInfo,
    ChunkedProofreadIssue,
    ChunkedProofreadRequest,
    ChunkedProofreadResult,
    ChunkedTaskStatus,
    ProofreadChunk,
    ProofreadIssue,
    ProofreadResponse,
)
from app.services import proofread as proofread_service
from app.services.ai_client import DEFAULT_TEMPERATURE


@dataclass(frozen=True)
class AgentTextResult:
    run_id: str
    status: ChunkedTaskStatus
    total_chunks: int
    completed_chunks: int
    failed_chunks: int
    issues: list[ChunkedProofreadIssue]
    error_message: str | None


class AgentRunner:
    def create_run(
        self,
        flow: AgentFlow,
        *,
        task_id: str | None = None,
        total_chunks: int = 0,
        metadata: dict[str, Any] | None = None,
    ) -> str:
        return trace.create_run(flow, task_id=task_id, total_chunks=total_chunks, metadata=metadata)

    async def proofread_selection(
        self,
        text: str,
        book: BookInfo,
        *,
        options: AgentOptions,
        proofread_callable: ProofreadChunkCallable = proofread_service.proofread_text,
    ) -> ProofreadResponse:
        result = await self.run_text(
            text,
            book,
            flow="selection",
            options=options,
            proofread_callable=proofread_callable,
            continue_on_chunk_error=False,
        )
        issues = [
            ProofreadIssue.model_validate(
                {key: value for key, value in issue.model_dump().items() if key not in {"chunk_index", "global_start", "global_end"}}
            )
            for issue in result.issues
        ]
        return ProofreadResponse(issues=issues, run_id=result.run_id)

    async def proofread_chunked_request(
        self,
        request: ChunkedProofreadRequest,
        *,
        task_id: str | None = None,
        proofread_callable: ProofreadChunkCallable = proofread_service.proofread_text,
    ) -> ChunkedProofreadResult:
        result = await self.run_text(
            request.text,
            request.book,
            flow="chunked",
            task_id=task_id,
            options=options_from_chunked_request(request),
            proofread_callable=proofread_callable,
            continue_on_chunk_error=True,
        )
        return ChunkedProofreadResult(
            task_id=task_id,
            run_id=result.run_id,
            scope=request.scope,
            status=result.status,
            total_chunks=result.total_chunks,
            completed_chunks=result.completed_chunks,
            failed_chunks=result.failed_chunks,
            issues=result.issues,
            error_message=result.error_message,
        )

    async def run_text(
        self,
        text: str,
        book: BookInfo,
        *,
        flow: AgentFlow,
        options: AgentOptions,
        proofread_callable: ProofreadChunkCallable,
        continue_on_chunk_error: bool,
        task_id: str | None = None,
        run_id: str | None = None,
        chunks: list[ProofreadChunk] | None = None,
    ) -> AgentTextResult:
        resolved_run_id = run_id or trace.create_run(
            flow,
            task_id=task_id,
            total_chunks=len(chunks or []),
            metadata=metadata_from_options(options),
        )
        graph = build_proofread_graph()
        try:
            state = await graph.ainvoke(
                {
                    "run_id": resolved_run_id,
                    "flow": flow,
                    "text": text,
                    "book": book,
                    "options": options,
                    "proofread_callable": proofread_callable,
                    "continue_on_chunk_error": continue_on_chunk_error,
                    **({"chunks": chunks} if chunks is not None else {}),
                }
            )
        except Exception as exc:
            trace.update_run(resolved_run_id, status="failed", error_message=str(exc))
            raise

        return AgentTextResult(
            run_id=resolved_run_id,
            status=state.get("status", "succeeded"),
            total_chunks=len(state.get("chunks", [])),
            completed_chunks=state.get("completed_chunks", 0),
            failed_chunks=state.get("failed_chunks", 0),
            issues=state.get("issues", []),
            error_message=state.get("error_message"),
        )

    async def proofread_task_chunk(
        self,
        *,
        run_id: str,
        chunk: ProofreadChunk,
        book: BookInfo,
        options: AgentOptions,
        retry_count: int = 0,
        proofread_callable: ProofreadChunkCallable,
    ) -> list[ProofreadIssue]:
        node_id = trace.start_node(run_id, "proofread_chunk")
        trace.start_chunk(
            run_id,
            chunk_index=chunk.index,
            chunk_start=chunk.start,
            chunk_end=chunk.end,
            chunk_len=len(chunk.text),
            retry_count=retry_count,
        )
        try:
            kwargs = {
                "session_id": options.session_id,
                "provider_api": options.provider_api,
                "proofread_mode": options.proofread_mode,
            }
            if options.ai_profile_id is not None:
                kwargs["ai_profile_id"] = options.ai_profile_id
            if options.reasoning_enabled:
                kwargs["reasoning_enabled"] = True
            if options.temperature != DEFAULT_TEMPERATURE:
                kwargs["temperature"] = options.temperature
            issues = await proofread_callable(chunk.text, book, **kwargs)
        except Exception as exc:
            trace.finish_chunk(
                run_id,
                chunk.index,
                status="failed",
                issue_count=0,
                retry_count=retry_count,
                error_message=str(exc),
            )
            trace.finish_node(node_id, "failed", str(exc))
            raise

        trace.finish_chunk(
            run_id,
            chunk.index,
            status="succeeded",
            issue_count=len(issues),
            retry_count=retry_count,
        )
        trace.finish_node(node_id, "succeeded")
        return issues

    def record_docx_writeback(
        self,
        run_id: str,
        writeback,
        *args,
        issue_count: int,
        **kwargs,
    ):
        with trace.record_node(run_id, "write_docx_output"):
            result = writeback(*args, **kwargs)
        trace.update_run(run_id, issue_count=issue_count)
        return result


def options_from_chunked_request(request: ChunkedProofreadRequest) -> AgentOptions:
    return AgentOptions(
        session_id=request.session_id,
        ai_profile_id=request.ai_profile_id,
        provider_api=request.provider_api,
        proofread_mode=request.proofread_mode,
        reasoning_enabled=request.reasoning_enabled,
        temperature=request.temperature,
        scope=request.scope,
        chunk_size=request.chunk_size,
    )


def metadata_from_options(options: AgentOptions) -> dict[str, Any]:
    return {
        "session_id_suffix": options.session_id[-8:] if options.session_id else None,
        "ai_profile_id": options.ai_profile_id or "default",
        "provider_api": options.provider_api or "default",
        "proofread_mode": options.proofread_mode,
        "reasoning_enabled": options.reasoning_enabled,
        "temperature": options.temperature,
        "scope": options.scope,
        "chunk_size": options.chunk_size,
        "application_mode": options.application_mode,
        "fallback_summary_truncate_enabled": options.fallback_summary_truncate_enabled,
    }


agent_runner = AgentRunner()

