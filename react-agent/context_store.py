from contextvars import ContextVar
from typing import Any, Dict, Optional

# Context variables for per-request dynamic configuration
_ctx: ContextVar[Dict[str, Any]] = ContextVar("sharepoint_context", default={})

CONFIG_KEYS = [
    "SITE_URL",
    "SITE_ID",
    "USER_NAME",
    "USER_ASSERTION",
    "ACCESS_TOKEN",
    "OBO_ACCESS_TOKEN",
]


def set_context(**kwargs) -> Dict[str, Any]:
    """Set multiple context values for the current request.
    Only known config keys are stored. Returns the updated context dict.
    """
    current = dict(_ctx.get())
    for k, v in kwargs.items():
        if k in CONFIG_KEYS and v is not None:
            current[k] = v
    _ctx.set(current)
    return current


def get_context_value(key: str) -> Optional[str]:
    return _ctx.get().get(key)


def get_all_context() -> Dict[str, Any]:
    return dict(_ctx.get())


def clear_context():
    _ctx.set({})
