import requests
import os   
import asyncio
import json
import time
from src.get_token import acquire_token
from src.tooling import M365CopilotPlugin, LocalDocumentPlugin
import sys
import logging
from dotenv import load_dotenv
from semantic_kernel import Kernel
from semantic_kernel.connectors.ai.open_ai import AzureChatCompletion
from semantic_kernel.connectors.ai.open_ai import AzureChatPromptExecutionSettings
from semantic_kernel.connectors.ai.function_choice_behavior import FunctionChoiceBehavior

# --- Configure Logging to verify which plugin is triggered ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logging.getLogger("semantic_kernel").setLevel(logging.DEBUG) # Use DEBUG for maximum detail on function calls


# Using the beta endpoint for the Chat API
load_dotenv()

# --- 1. Configuration (Use environment variables) ---
# Ensure AZURE_AI_ENDPOINT, AZURE_AI_KEY, AZURE_AI_MODEL environment variables are set.
GRAPH_API_URL = os.getenv("GRAPH_API_URL") 
AZURE_AI_ENDPOINT = os.getenv("AZURE_AI_ENDPOINT")
AZURE_AI_KEY = os.getenv("AZURE_AI_KEY")
AZURE_AI_MODEL = os.getenv("AZURE_AI_MODEL")

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

async def main():
    """
    Main orchestration script using Semantic Kernel configured with an Azure AI Foundry model
    and a command-line interface.
    """
    # --- 2. Acquire Graph Access Token ---
    print(f"Acquiring Graph Access Token...")
    os.environ["M365_TOKEN"] = get_access_token()

    # --- 3. Create Conversation ---
    print(f"Creating Conversation...")
    os.environ["M365_CONVO_ID"]  = create_conversation(os.environ["M365_TOKEN"])

    # --- 4. Initialize Semantic Kernel ---
    print(f"Initializing Semantic Kernel...")    
    kernel = Kernel()

    # --- 5. Initialize the Semantic Kernel and Azure AI Foundry Service ---
    print(f"Initializing Azure AI Foundry Service...")
    kernel.add_service(
        AzureChatCompletion(
            deployment_name=AZURE_AI_MODEL,
            endpoint=AZURE_AI_ENDPOINT,
            api_key=AZURE_AI_KEY
        )
    )
    execution_settings = AzureChatPromptExecutionSettings(
        function_choice_behavior=FunctionChoiceBehavior.Auto(auto_invoke=True)
    )

    # --- 6. Import the Plugin ---
    print(f"Importing Plugin...")
    m365copilot_chat_plugin = M365CopilotPlugin(token=os.environ["M365_TOKEN"], conversation_id=os.environ["M365_CONVO_ID"])
    kernel.add_plugin(plugin_name="M365CopilotChat", plugin=m365copilot_chat_plugin)
    local_document_plugin = LocalDocumentPlugin()
    kernel.add_plugin(plugin_name="LocalDocumentPlugin", plugin=local_document_plugin)

    # --- 7. Check Required Environment Variables ---
    print("Checking Required Environment Variables...")
    if not all([os.environ["AZURE_AI_ENDPOINT"], os.environ["AZURE_AI_KEY"], os.environ["AZURE_AI_MODEL"], os.environ["M365_TOKEN"], os.environ["M365_CONVO_ID"]]):
        print("Error: Required environment variables must be set.")
        return

    print(f"Kernel initialized and ready for input. Type 'exit' to quit.")

    # --- 4. The Command-Line Loop ---
    while True:
        user_prompt = (await ainput("\n>>> User: ")).strip()
        
        if user_prompt.lower() == "exit":
            print("Exiting application...")
            print("\nCleaning up...")
            #cleanup_result = kernel.invoke_function_call("M365CopilotChat", "end_conversation")
            print("\nGoodbye!")
            break
        
        if not user_prompt:
            continue

        print("Thinking...")

        # Invoke the kernel: the LLM automatically decides to call the appropriate plugin function
        try:
            result = await kernel.invoke_prompt(prompt=user_prompt, settings=execution_settings)
            print(f"\n<<< AI: {result}")
        except Exception as e:
            print(f"\nAn error occurred during AI invocation: {e}")
            break

    # Optional: Call the cleanup function after the loop terminates
    print("\nCleaning up...")
    #cleanup_result = kernel.invoke_plugin_function("M365CopilotChat", "end_conversation")
    print("\nGoodbye!")


if __name__ == "__main__":
    # Run the main asynchronous function using asyncio
    asyncio.run(main())