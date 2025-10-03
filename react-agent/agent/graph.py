from langgraph.checkpoint.memory import InMemorySaver
from langgraph.constants import END
from langgraph.graph import MessagesState, StateGraph

from agent.nodes import run_agent_reasoning, tool_node

AGENT_REASON = "agent_reason"
ACT = "act"
LAST = -1


def _should_continue(state: MessagesState) -> str:
    if not state["messages"][LAST].tool_calls:
        return END
    return ACT


def build_app():
    g = StateGraph(MessagesState)
    g.add_node(AGENT_REASON, run_agent_reasoning)
    g.set_entry_point(AGENT_REASON)
    g.add_node(ACT, tool_node)
    g.add_conditional_edges(AGENT_REASON, _should_continue, {END: END, ACT: ACT})
    g.add_edge(ACT, AGENT_REASON)
    memory = InMemorySaver()
    graph = g.compile(checkpointer=memory)
    return graph
