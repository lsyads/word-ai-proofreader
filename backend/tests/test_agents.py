import asyncio

import pytest

from app.agents import trace as agent_trace
from app.agents.graph import build_proofread_graph
from app.agents.service import AgentOptions, agent_runner
from app.agents.tools import get_proofread_tools
from app.schemas import BookInfo, ChunkedProofreadRequest, ProofreadIssue
from app.services.ai_client import AIClientError


BOOK = BookInfo(title="测试书名", introduction="测试介绍")


def setup_function():
    agent_trace.clear_traces_for_tests()


def test_graph_compiles():
    graph = build_proofread_graph()

    assert graph is not None


def test_tools_use_snake_case_schema_and_docstrings():
    tools = get_proofread_tools()

    assert {tool.name for tool in tools} >= {
        "split_text",
        "proofread_with_model",
        "globalize_issue_offsets",
        "locate_issue_offsets",
        "parse_docx_document",
        "write_docx_output",
    }
    assert all(tool.name == tool.name.lower() and "-" not in tool.name for tool in tools)
    assert all(tool.args_schema is not None for tool in tools)
    assert all(tool.description for tool in tools)
    split_tool = next(tool for tool in tools if tool.name == "split_text")
    chunks = split_tool.invoke({"text": "甲" * 10, "scope": "selection", "chunk_size": 500})
    assert len(chunks) == 1
    assert chunks[0].start == 0
    assert chunks[0].end == 10


def test_agent_runner_records_trace_without_full_text_or_secret():
    async def fake_proofread(text, book, **kwargs):
        return [
            ProofreadIssue(
                id="issue-1",
                category="typo",
                severity="low",
                original="敏感",
                suggestion="建议",
                start=0,
                end=2,
            )
        ]

    request = ChunkedProofreadRequest(
        text="敏感正文" + ("甲" * 2995) + "。" + ("乙" * 2999) + "。",
        book=BOOK,
        scope="document",
        chunk_size=3000,
    )

    result = asyncio.run(agent_runner.proofread_chunked_request(request, proofread_callable=fake_proofread))
    trace = agent_trace.get_trace(result.run_id)
    serialized_trace = trace.model_dump_json()

    assert result.run_id
    assert result.status == "succeeded"
    assert trace.total_chunks == 2
    assert len(trace.nodes) >= 4
    assert len(trace.chunks) == 2
    assert all(chunk.status == "succeeded" for chunk in trace.chunks)
    assert "敏感正文" not in serialized_trace
    assert "Authorization" not in serialized_trace
    assert "Bearer" not in serialized_trace


def test_agent_runner_records_failed_chunk_and_continues():
    async def fake_proofread(text, book, **kwargs):
        if text.startswith("乙"):
            raise AIClientError("chunk failed with Authorization: Bearer secret-token")
        return [
            ProofreadIssue(
                id="issue-1",
                category="typo",
                severity="low",
                original=text[:1],
                suggestion="建议",
                start=0,
                end=1,
            )
        ]

    request = ChunkedProofreadRequest(
        text=("甲" * 2999) + "。" + ("乙" * 2999) + "。",
        book=BOOK,
        scope="document",
        chunk_size=3000,
    )

    result = asyncio.run(agent_runner.proofread_chunked_request(request, proofread_callable=fake_proofread))
    trace = agent_trace.get_trace(result.run_id)

    assert result.status == "partial_succeeded"
    assert result.completed_chunks == 1
    assert result.failed_chunks == 1
    assert trace.chunks[1].status == "failed"
    assert trace.chunks[1].error_message == "chunk failed with Authorization: Bearer [REDACTED]"


def test_agent_task_chunk_records_retry_count():
    async def fake_proofread(text, book, **kwargs):
        return [
            ProofreadIssue(
                id="issue-1",
                category="typo",
                severity="low",
                original=text[:1],
                suggestion="建议",
                start=0,
                end=1,
            )
        ]

    run_id = agent_runner.create_run("chunked_task", total_chunks=1)
    chunk = ChunkedProofreadRequest(
        text="甲" * 3000,
        book=BOOK,
        scope="document",
        chunk_size=3000,
    )
    proofread_chunk = __import__("app.services.chunking", fromlist=["split_text_into_chunks"]).split_text_into_chunks(
        chunk.text,
        chunk.scope,
        chunk.chunk_size,
    )[0]

    asyncio.run(
        agent_runner.proofread_task_chunk(
            run_id=run_id,
            chunk=proofread_chunk,
            book=BOOK,
            options=AgentOptions(scope="document", chunk_size=3000),
            retry_count=2,
            proofread_callable=fake_proofread,
        )
    )
    trace = agent_trace.get_trace(run_id)

    assert trace.chunks[0].retry_count == 2
    assert trace.chunks[0].issue_count == 1


def test_missing_trace_raises_not_found():
    with pytest.raises(agent_trace.AgentTraceNotFound):
        agent_trace.get_trace("missing-run")
