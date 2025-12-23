import msal
from msal import PublicClientApplication
import logging
import asyncio
import webbrowser
from dotenv import load_dotenv
import os
from .local_token_cache import LocalTokenCache


logger = logging.getLogger(__name__)
load_dotenv()

# --- Fix: Initialize the cache object FIRST ---
cache = msal.SerializableTokenCache()

#cache file path
CACHE_FILE = "./msal_cache.json"

TOKEN_CACHE = LocalTokenCache("./.local_token_cache.json")

# Configuration from your Azure AD App Registration
CLIENT_ID = "774142ce-9070-446b-83ac-e2053c716879"
TENANT_ID = "8b7a11d9-6513-4d54-a468-f6630df73c8b"

# Required permissions (scopes)
SCOPES = ["https://graph.microsoft.com/.default"] 

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

async def open_browser(url: str):
    logger.debug(f"Opening browser at {url}")
    await asyncio.get_event_loop().run_in_executor(None, lambda: webbrowser.open(url))

def save_cache_on_exit():
    if cache.has_state_changed:
        print(f"Cache state changed. Saving to {CACHE_FILE}...")
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

def acquire_token():
    # Load the cache from the file if it exists
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r") as f:
                cache.deserialize(f.read())
            print(f"Loaded token cache from {CACHE_FILE}")
        except Exception as e:
            print(f"Error loading cache: {e}")  

    pca = PublicClientApplication(
        client_id=CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
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

    def acquire_non_interactive_token(tenant_id, client_id, client_secret):
        # Acquires an access token using client credentials flow.
        token_url = f"login.microsoftonline.com{tenant_id}/oauth2/v2.0/token"
        payload = {
            'client_id': client_id,
            'client_secret': client_secret,
            'scope': 'https://graph.microsoft.com/.default',
            'grant_type': 'client_credentials'
        }
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}

        response = requests.post(token_url, data=payload, headers=headers)
        response.raise_for_status()
        return response.json().get('access_token')