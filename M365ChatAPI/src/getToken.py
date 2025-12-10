from msal import PublicClientApplication
import logging
import asyncio
import webbrowser
from dotenv import load_dotenv
import os
from .local_token_cache import LocalTokenCache

logger = logging.getLogger(__name__)
load_dotenv()
TOKEN_CACHE = LocalTokenCache("./.local_token_cache.json")

# Configuration from your Azure AD App Registration
CLIENT_ID = "774142ce-9070-446b-83ac-e2053c716879"
TENANT_ID = "8b7a11d9-6513-4d54-a468-f6630df73c8b"

# Required permissions (scopes)
SCOPES = ["https://graph.microsoft.com/.default"] 

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_API_URL = "https://graph.microsoft.com/beta/copilot" # Using the beta endpoint for the Chat API


async def open_browser(url: str):
    logger.debug(f"Opening browser at {url}")
    await asyncio.get_event_loop().run_in_executor(None, lambda: webbrowser.open(url))

def acquire_token():
    pca = PublicClientApplication(
        client_id=CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=TOKEN_CACHE,
    )

    token_request = {
        "scopes": SCOPES,
    }
    accounts = pca.get_accounts()
    retry_interactive = False
    token = None
    try:
        if accounts:
            response = pca.acquire_token_silent(
                token_request["scopes"], account=accounts[0]
            )
            token = response.get("access_token")
        else:
            retry_interactive = True
    except Exception as e:
        retry_interactive = True
        logger.error(
            f"Error acquiring token silently: {e}. Going to attempt interactive login."
        )

    if retry_interactive:
        logger.debug("Attempting interactive login...")
        response = pca.acquire_token_interactive(**token_request)
        token = response.get("access_token")

    return token
