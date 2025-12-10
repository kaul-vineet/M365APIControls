import requests
import asyncio
import json
import time
from src.getToken import acquire_token
import sys

GRAPH_API_URL = "https://graph.microsoft.com/beta/copilot" # Using the beta endpoint for the Chat API

async def ainput(string: str) -> str:
    await asyncio.get_event_loop().run_in_executor(
        None, lambda s=string: sys.stdout.write(s + " ")
    )
    return await asyncio.get_event_loop().run_in_executor(None, sys.stdin.readline)

def get_access_token():
    if acquire_token():
        return acquire_token()
    else:
        raise Exception(f"Could not acquire token")

def create_conversation(token):
    """Creates a new Copilot conversation and returns its ID."""
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    # An empty body is required to create a new conversation
    response = requests.post(f"{GRAPH_API_URL}/conversations", headers=headers, data=json.dumps({}))
    print(f"Response: {response.json()}")
    response.raise_for_status()
    conversation_data = response.json()
    print(f"Created conversation with ID: {conversation_data['id']}")
    return conversation_data["id"]

async def send_message(token, conversation_id):
    prompt_text = (await ainput("\n>>>: ")).lower().strip()

    if prompt_text == "exit":
        sys.exit(0)
    else:
        """Sends a message to an existing conversation and gets the response."""
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        payload = {
            "message": {
            "text": prompt_text
        },
        "locationHint": {
            "timeZone": "America/New_York"
            }
        }

        # The endpoint for continuing a synchronous chat
        url = f"{GRAPH_API_URL}/conversations/{conversation_id}/chat"
        response = requests.post(url, headers=headers, data=json.dumps(payload))
        response.raise_for_status()
        
        # Process the response to extract the Copilot's answer
        response_data = response.json()
        try:
            print(f"\nCopilot: {response_data['messages'][1]['text']}")
        except (KeyError, IndexError):
            print("Error: Could not extract specific message text from response.")
            print(f"Full response data: {response_data}")
        await send_message(token, conversation_id) 

async def main():
    print("\nSay Hi to connect to M365 Copilot.... ")
    try:
        token = get_access_token()
        conversation_id = create_conversation(token)
        await send_message(token, conversation_id)
    except Exception as e:
        print(f"An error occurred: {e}")

asyncio.run(main())        
