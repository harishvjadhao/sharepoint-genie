from langgraph.graph import MessagesState
from langgraph.prebuilt import ToolNode

from context_store import get_context_value
from tools.react import llm, tools

SYSTEM_MESSAGE = """
You are a helpful assistant that can use tools to answer questions.
Greet the user by their name if provided. The user's name is: {user_name}
"""


def run_agent_reasoning(state: MessagesState) -> MessagesState:
    """
    Run the agent reasoning node.
    """
    user_name = get_context_value("USER_NAME") or "User"
    system_message = SYSTEM_MESSAGE.format(user_name=user_name)
    response = llm.invoke(
        [{"role": "system", "content": system_message}, *state["messages"]]
    )
    return {"messages": [response]}


tool_node = ToolNode(tools)
