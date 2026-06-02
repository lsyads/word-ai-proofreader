from __future__ import annotations

import logging

from app.agents import trace
from app.agents.state import AgentState
from app.schemas import ChunkedProofreadIssue
from app.services import chunk_service
from app.services.ai_client import DEFAULT_TEMPERATURE

logger = logging.getLogger(__name__)


async def prepare_input(state: AgentState) -> dict:
    with trace.record_node(state["run_id"], "prepare_input"):
        options = state["options"]
        metadata = {
            "flow": state["flow"],
            "scope": options.scope,
            "ai_profile_id": options.ai_profile_id or "default",
            "provider_api": options.provider_api or "default",
            "proofread_mode": options.proofread_mode,
            "reasoning_enabled": options.reasoning_enabled,
            "temperature": options.temperature,
        }
        trace.update_run(state["run_id"], status="running")
        logger.info(
            "agent run prepared run_id=%s flow=%s text_len=%s scope=%s",
            state["run_id"],
            state["flow"],
            len(state["text"]),
            options.scope,
        )
        return {"issues": [], "completed_chunks": 0, "failed_chunks": 0, "error_message": None, "metadata": metadata}


async def split_chunks(state: AgentState) -> dict:
    with trace.record_node(state["run_id"], "split_chunks"):
        if "chunks" in state and state["chunks"]:
            chunks = state["chunks"]
        else:
            options = state["options"]
            chunks = chunk_service.split_text_into_chunks(state["text"], options.scope, options.chunk_size)
        trace.update_run(state["run_id"], total_chunks=len(chunks))
        return {"chunks": chunks}


async def proofread_chunks(state: AgentState) -> dict:
    with trace.record_node(state["run_id"], "proofread_chunks"):
        options = state["options"]
        proofread_callable = state["proofread_callable"]
        aggregated: list[ChunkedProofreadIssue] = []
        completed_chunks = 0
        failed_chunks = 0

        for chunk in state.get("chunks", []):
            trace.start_chunk(
                state["run_id"],
                chunk_index=chunk.index,
                chunk_start=chunk.start,
                chunk_end=chunk.end,
                chunk_len=len(chunk.text),
            )
            try:
                proofread_kwargs = {
                    "session_id": options.session_id,
                    "provider_api": options.provider_api,
                    "proofread_mode": options.proofread_mode,
                }
                if options.ai_profile_id is not None:
                    proofread_kwargs["ai_profile_id"] = options.ai_profile_id
                if options.reasoning_enabled:
                    proofread_kwargs["reasoning_enabled"] = True
                if options.temperature != DEFAULT_TEMPERATURE:
                    proofread_kwargs["temperature"] = options.temperature

                chunk_issues = await proofread_callable(chunk.text, state["book"], **proofread_kwargs)
            except Exception as exc:
                failed_chunks += 1
                trace.finish_chunk(
                    state["run_id"],
                    chunk.index,
                    status="failed",
                    issue_count=0,
                    error_message=str(exc),
                )
                trace.update_run(
                    state["run_id"],
                    completed_chunks=completed_chunks,
                    failed_chunks=failed_chunks,
                    issue_count=len(aggregated),
                )
                logger.exception("agent chunk failed run_id=%s chunk_index=%s", state["run_id"], chunk.index)
                if not state["continue_on_chunk_error"]:
                    raise
                continue

            completed_chunks += 1
            global_issues = chunk_service.globalize_issues(chunk, chunk_issues)
            aggregated.extend(global_issues)
            trace.finish_chunk(
                state["run_id"],
                chunk.index,
                status="succeeded",
                issue_count=len(chunk_issues),
            )
            trace.update_run(
                state["run_id"],
                completed_chunks=completed_chunks,
                failed_chunks=failed_chunks,
                issue_count=len(aggregated),
            )

        status = _status_from_counts(completed_chunks, failed_chunks)
        error_message = "All chunks failed to proofread." if status == "failed" else None
        trace.update_run(
            state["run_id"],
            status=status,
            completed_chunks=completed_chunks,
            failed_chunks=failed_chunks,
            issue_count=len(aggregated),
            error_message=error_message,
        )
        return {
            "issues": aggregated,
            "completed_chunks": completed_chunks,
            "failed_chunks": failed_chunks,
            "status": status,
            "error_message": error_message,
        }


async def finalize_run(state: AgentState) -> dict:
    with trace.record_node(state["run_id"], "finalize"):
        status = state.get("status") or _status_from_counts(
            state.get("completed_chunks", 0),
            state.get("failed_chunks", 0),
        )
        trace.update_run(
            state["run_id"],
            status=status,
            completed_chunks=state.get("completed_chunks", 0),
            failed_chunks=state.get("failed_chunks", 0),
            issue_count=len(state.get("issues", [])),
            error_message=state.get("error_message"),
        )
        return {"status": status}


def _status_from_counts(completed_chunks: int, failed_chunks: int) -> str:
    if completed_chunks == 0 and failed_chunks > 0:
        return "failed"
    if completed_chunks > 0 and failed_chunks > 0:
        return "partial_succeeded"
    return "succeeded"

