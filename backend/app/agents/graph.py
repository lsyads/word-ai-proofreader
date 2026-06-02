from __future__ import annotations

from functools import lru_cache

from langgraph.graph import END, START, StateGraph

from app.agents import nodes
from app.agents.state import AgentState


@lru_cache(maxsize=1)
def build_proofread_graph():
    graph = StateGraph(AgentState)
    graph.add_node("prepare_input", nodes.prepare_input)
    graph.add_node("split_chunks", nodes.split_chunks)
    graph.add_node("proofread_chunks", nodes.proofread_chunks)
    graph.add_node("finalize", nodes.finalize_run)

    graph.add_edge(START, "prepare_input")
    graph.add_edge("prepare_input", "split_chunks")
    graph.add_edge("split_chunks", "proofread_chunks")
    graph.add_edge("proofread_chunks", "finalize")
    graph.add_edge("finalize", END)
    return graph.compile()

