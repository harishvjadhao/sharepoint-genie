import os

from dotenv import load_dotenv
from langchain_core.messages import HumanMessage

from agent.graph import build_app
from context_store import set_context
from services.acquire_token import AcquireToken
from services.site_info import SiteInfo

load_dotenv()

LAST = -1


def bootstrap():
    # Optionally set context from env or external caller

    token = AcquireToken(
        site_url=os.environ["SITE_URL"],
        tenant_id=os.environ["TENANT_ID"],
        client_id=os.environ["CLIENT_ID"],
        client_secret=os.environ["CLIENT_SECRET"],
        resource_url=os.environ["RESOURCE"],
        user_assertion=os.environ["USER_ASSERTION"],
    )

    site = SiteInfo(
        site_url=os.environ["SITE_URL"],
        access_token=token.access_token,
    )

    set_context(
        SITE_URL=os.environ["SITE_URL"],
        SITE_ID=site.site_id,
        USER_ASSERTION=os.environ["USER_ASSERTION"],
        ACCESS_TOKEN=token.access_token,
        OBO_ACCESS_TOKEN=token.obo_access_token,
    )
    return build_app()


if __name__ == "__main__":
    app = bootstrap()
    res = app.invoke({"messages": [HumanMessage(content="get the site analytics")]})
    print(res["messages"][LAST].content)
