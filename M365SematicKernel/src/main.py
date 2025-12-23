import atexit
from rich import print
from rich.console import Console
from rich.panel import Panel
from rich.text import Text
from rich.prompt import Prompt
import requests
import os   
import asyncio
import json
import time
from src.auth import *   
from src.tooling import M365CopilotPlugin, LocalDocumentGeneratorPlugin, GraphSharePointUploaderPlugin
import sys
import logging
from dotenv import load_dotenv
from semantic_kernel import Kernel
from semantic_kernel.connectors.ai.open_ai import AzureChatCompletion
from semantic_kernel.connectors.ai.open_ai import AzureChatPromptExecutionSettings
from semantic_kernel.connectors.ai.function_choice_behavior import FunctionChoiceBehavior
from semantic_kernel.contents import ChatHistory

# --- Configure Logging to verify which plugin is triggered ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logging.getLogger("semantic_kernel").setLevel(logging.DEBUG) # Use DEBUG for maximum detail on function calls


# Using the beta endpoint for the Chat API
load_dotenv()
console = Console()

# --- 1. Configuration (Use environment variables) ---
# Ensure AZURE_AI_ENDPOINT, AZURE_AI_KEY, AZURE_AI_MODEL, SHAREPOINT environment variables are set.
global GRAPH_API_URL, AZURE_AI_ENDPOINT, AZURE_AI_KEY, AZURE_AI_MODEL, SITE_URL, FOLDER

GRAPH_API_URL = os.getenv("GRAPH_API_URL") 
AZURE_AI_ENDPOINT = os.getenv("AZURE_AI_ENDPOINT")
AZURE_AI_KEY = os.getenv("AZURE_AI_KEY")
AZURE_AI_MODEL = os.getenv("AZURE_AI_MODEL")
SITE_URL = os.environ["SITE_URL"]
FOLDER = os.environ["FOLDER"]

# Register the save_cache_on_exit function to be called when the program exits
atexit.register(save_cache_on_exit)

def display_message(user, message, color="cyan"):
    """Displays a chat message inside a styled panel."""
    panel_content = Text(message)
    panel = Panel(
        panel_content, 
        title=f"[bold]{user}[/bold]", 
        border_style=color, 
        expand=False
    )
    console.print(panel)

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
    #print(f"Response: {response.json()}")
    response.raise_for_status()
    conversation_data = response.json()
    print(f"Created conversation with ID: {conversation_data['id']}")
    return conversation_data["id"]

async def main():
    """
    Main orchestration script using Semantic Kernel configured with an Azure AI Foundry model
    and a command-line interface.
    """
    console.print(f"[bold magenta] Starting Communications with M365 Copilot...[/]")

    # --- 2. Acquire Graph Access Token ---
    console.print("[blue] Acquiring Graph Access Token...[/]")
    os.environ["M365_TOKEN"] = get_access_token()

    # --- 3. Create Conversation ---
    console.print("[green] Creating Conversation...[/]")
    os.environ["M365_CONVO_ID"] = create_conversation(os.environ["M365_TOKEN"])

   # --- 7. Check Required Environment Variables ---
    console.print("[blue] Checking Required Environment Variables...[/]")
    if not all([os.environ["AZURE_AI_ENDPOINT"], os.environ["AZURE_AI_KEY"], os.environ["AZURE_AI_MODEL"], os.environ["M365_TOKEN"], os.environ["M365_CONVO_ID"], os.environ["SITE_URL"]]):
        console.print("Error: Required environment variables must be set.")
        return

    # --- 4. Initialize Semantic Kernel ---
    console.print("[blue] Initializing Semantic Kernel...[/]")    
    kernel = Kernel()

    # --- 5. Initialize the Semantic Kernel and Azure AI Foundry Service ---
    console.print("[cyan] Initializing Azure AI Foundry Service...[/]")

    azure_chat_service = AzureChatCompletion(
        deployment_name=AZURE_AI_MODEL,
        endpoint=AZURE_AI_ENDPOINT,
        api_key=AZURE_AI_KEY
    )
    kernel.add_service(azure_chat_service)

    execution_settings = AzureChatPromptExecutionSettings(
        function_choice_behavior=FunctionChoiceBehavior.Auto(auto_invoke=True)
    ) 

    console.print(f"[bright_green] Adding Chat History...[/]")
    history = ChatHistory(system_message="You are a assistant agent and your role is to help with documentations and information from Office 365.")

    # --- 6. Import the Plugin ---
    console.print(f"[green] Importing Plugin...[/]")
    m365copilot_chat_plugin = M365CopilotPlugin()
    kernel.add_plugin(plugin_name="M365CopilotChat", plugin=m365copilot_chat_plugin)
    
    local_document_generator_plugin = LocalDocumentGeneratorPlugin()
    kernel.add_plugin(plugin_name="LocalDocumentGeneratorPlugin", plugin=local_document_generator_plugin)
    
    graph_sharepoin_uploader_plugin = GraphSharePointUploaderPlugin(generator_plugin=local_document_generator_plugin)
    kernel.add_plugin(plugin_name="GraphSharePointUploaderPlugin", plugin=graph_sharepoin_uploader_plugin)

    console.print(f"[bold magenta] Kernel initialized and ready for input. Type 'exit' to quit.[/]")

    # --- 4. The Command-Line Loop ---
    while True:
        user_prompt = (await ainput("\n>>> User: ")).strip()
        if user_prompt.lower() == "exit":
            print("Exiting application...")
            print("\nCleaning up...")
            #cleanup_result = kernel.invoke_function_call("M365CopilotChat", "end_conversation")    
            print("\nGoodbye!")
            sys.exit()
        
        if not user_prompt:
            continue

        display_message("\n<<< M365 Copilot", f"Thinking.....", color="cyan")

        # Invoke the kernel: the LLM automatically decides to call the appropriate plugin function
        try:
            # result = await kernel.invoke_prompt(prompt=user_prompt, settings=execution_settings)
            # 5. Get Content (The Kernel handles Azure tool-calling loops automatically)
            history.add_user_message(user_prompt)
            result = await azure_chat_service.get_chat_message_content(
                chat_history=history,
                settings=execution_settings,
                kernel=kernel
            )
            display_message("\n<<< M365 Copilot", f"{result}", color="blue")
        except Exception as e:
            print(f"[bold red] An error occurred during AI invocation: {e}[/]")
            break

    # Optional: Call the cleanup function after the loop terminates
    print("\nCleaning up...")
    #cleanup_result = kernel.invoke_plugin_function("M365CopilotChat", "end_conversation")
    print("\nGoodbye!")
    sys.exit()


if __name__ == "__main__":
    # Run the main asynchronous function using asyncio
    asyncio.run(main())