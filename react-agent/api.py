# api.py
import os
import uuid
from typing import Any, Dict, List

from fastapi import FastAPI, HTTPException, Security
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from langchain_core.messages import AIMessage, HumanMessage, SystemMessage
from pydantic import BaseModel

from agent.graph import build_app
from context_store import clear_context, set_context
from services.acquire_token import AcquireToken
from services.site_info import SiteInfo

app = FastAPI(title="SharePoint Genie Chat API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # tighten in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

graph_app = build_app()

# Simple in-memory store (replace with Redis or DB for production)
SESSIONS: Dict[str, List[Any]] = {}

# Reusable security scheme for extracting Bearer tokens
bearer_scheme = HTTPBearer(auto_error=False)


class InitSessionRequest(BaseModel):
    siteUrl: str


class InitSessionResponse(BaseModel):
    sessionId: str


class ChatRequest(BaseModel):
    sessionId: str
    message: str


class ChatResponse(BaseModel):
    sessionId: str
    reply: str
    toolCalls: List[Dict[str, Any]] | None = None


@app.post("/session/init", response_model=InitSessionResponse)
def init_session(
    req: InitSessionRequest,
    credentials: HTTPAuthorizationCredentials = Security(bearer_scheme),
):
    """Initialize a chat session.

    Authorization:
        Bearer <token> header. The bearer token will be used as the user_assertion (delegated user token for OBO flow).
    """
    sessionId = uuid.uuid4().hex
    SESSIONS[sessionId] = [
        SystemMessage(
            content="You are a helpful assistant for SharePoint and OneDrive tasks."
        )
    ]
    # Store context for this session (if you keep global context var, reset each request)
    SESSIONS[sessionId].append(
        AIMessage(content="Session initialized. How can I help you?")
    )
    # Derive user_assertion: body value wins; fallback to Authorization header if present
    header_token = None
    if credentials and credentials.scheme.lower() == "bearer":
        header_token = credentials.credentials

    token = AcquireToken(
        site_url=req.siteUrl,
        tenant_id=os.environ["TENANT_ID"],
        client_id=os.environ["CLIENT_ID"],
        client_secret=os.environ["CLIENT_SECRET"],
        resource_url=os.environ["RESOURCE"],
        user_assertion=header_token,
    )

    site = SiteInfo(
        site_url=req.siteUrl,
        access_token=token.access_token,
    )

    # Persist per-session context (avoid storing None fields)
    ctx_payload = {
        "SITE_URL": req.siteUrl,
        "SITE_ID": site.site_id,
        "USER_ASSERTION": header_token,
        "ACCESS_TOKEN": token.access_token,
        "OBO_ACCESS_TOKEN": token.obo_access_token,
    }
    # if header_token:
    #     ctx_payload["USER_ASSERTION"] = header_token
    SESSIONS[sessionId + "_ctx"] = ctx_payload
    return InitSessionResponse(sessionId=sessionId)


@app.post("/chat", response_model=ChatResponse)
def chat(
    req: ChatRequest,
    credentials: HTTPAuthorizationCredentials = Security(bearer_scheme),
):
    if req.sessionId not in SESSIONS:
        raise HTTPException(status_code=404, detail="Invalid sessionId")
    
    # Derive user_assertion: body value wins; fallback to Authorization header if present
    header_token = None
    if credentials and credentials.scheme.lower() == "bearer":
        header_token = credentials.credentials

    history = SESSIONS[req.sessionId]
    session_ctx = SESSIONS.get(req.sessionId + "_ctx", {})

    # Check if header_token matches session's USER_ASSERTION
    session_user_assertion = session_ctx.get("USER_ASSERTION")
    if session_user_assertion and header_token != session_user_assertion:
        raise HTTPException(status_code=401, detail="Session terminated: token mismatch")

    # Apply per-session context before invoke
    clear_context()
    set_context(**{k: v for k, v in session_ctx.items() if v})

    history.append(HumanMessage(content=req.message))
    state = {"messages": history[-1]}  # Limit context to last message
    config = {"configurable": {"thread_id": req.sessionId}}
    result = graph_app.invoke(state, config=config)
    new_messages = result["messages"]

    # LangGraph returns entire list; get only newly added tail piece(s)
    # For simplicity, take the last AI/Tool output message:
    reply_msg = new_messages[-1]
    history.append(reply_msg)

    toolCalls = getattr(reply_msg, "toolCalls", None)
    reply_text = getattr(reply_msg, "content", "")

    return ChatResponse(sessionId=req.sessionId, reply=reply_text, toolCalls=toolCalls)


if __name__ == "__main__":
    # FastAPI apps are ASGI; use uvicorn to run the server instead of Flask's app.run
    import uvicorn

    # Allow overriding port via PORT env var (common in cloud platforms)
    port = int(os.getenv("PORT", "8000"))
    # reload=True is handy for local dev; keep False here for explicitness
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=port,
        reload=False,
        log_level="info",
    )
