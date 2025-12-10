import asyncio
import os
import datetime
import sys
import webbrowser
import gradio as gr

from azure.identity import DeviceCodeCredential
from kiota_abstractions.api_error import APIError
from microsoft_agents_m365copilot.agents_m365_copilot_service_client import AgentsM365CopilotServiceClient
from microsoft_agents_m365copilot.generated.copilot.retrieval.retrieval_post_request_body import RetrievalPostRequestBody
from microsoft_agents_m365copilot.generated.models.retrieval_data_source import RetrievalDataSource

scopes = ['Files.Read.All', 'Sites.Read.All']

# Multi-tenant apps can use "common",
# single-tenant apps must use the tenant ID from the Azure portal
TENANT_ID = '8b7a11d9-6513-4d54-a468-f6630df73c8b'

# Values from app registration
CLIENT_ID = '774142ce-9070-446b-83ac-e2053c716879'

# Define a proper callback function that accepts all three parameters
def auth_callback(verification_uri: str, user_code: str, expires_on: datetime):
    print(f"\nTo sign in, use a web browser to open the page {verification_uri}")
    print(f"Enter the code {user_code} to authenticate.")
    print(f"The code will expire at {expires_on}")

async def ainput(string: str) -> str:
    await asyncio.get_event_loop().run_in_executor(
        None, lambda s=string: sys.stdout.write(s + " ")
    )
    return await asyncio.get_event_loop().run_in_executor(None, sys.stdin.readline)

async def ask_question():
    try:
        query = (await ainput("\n>>>: ")).lower().strip()
        if query:
            # Print the URL being used
            # print(f"Using API base URL: {client.request_adapter.base_url}\n")
            print(f"Query: {query}" + ". Search the SharePoint to get the information required. Summarize the information.")
            # Create the retrieval request body
            retrieval_body = RetrievalPostRequestBody()
            retrieval_body.data_source = RetrievalDataSource.SharePoint
            retrieval_body.query_string = query
        
            # Try more parameters that might be required
            # retrieval_body.maximum_number_of_results = 10
            
            # Make the API call
            print("Making retrieval API request...")
            retrieval = await client.copilot.retrieval.post(retrieval_body)
            
            # Process the results
            if retrieval and hasattr(retrieval, "retrieval_hits"):
                print(f"Received {len(retrieval.retrieval_hits)} hits")
                for r in retrieval.retrieval_hits:
                    print(f"Web URL: {r.web_url}\n")
                    for extract in r.extracts:
                        print(f"Text:\n{extract.text}\n")
                print(f"Retrieval response structure: {dir(retrieval)}")   
            else:
                print(f"Retrieval response structure: {dir(retrieval)}")
            await ask_question()        
    except APIError as e:
        print(f"Error: {e.error.code}: {e.error.message}")
        if hasattr(e, 'error') and hasattr(e.error, 'inner_error'):
            print(f"Inner error details: {e.error.inner_error}")
        raise e

async def main():
    print("\nSay Hi to connect to M365 Copilot.... ")

    # Create device code credential with correct callback
    credentials = DeviceCodeCredential(
        client_id=CLIENT_ID,
        tenant_id=TENANT_ID,
        prompt_callback=auth_callback
    )

    global client
    client = AgentsM365CopilotServiceClient(credentials=credentials, scopes=scopes)

    await ask_question()

asyncio.run(main())
