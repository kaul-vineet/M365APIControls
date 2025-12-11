import io
from typing import Annotated
from docx import Document
from semantic_kernel.functions.kernel_function_decorator import kernel_function

class M365CopilotPlugin:
    """
    A plugin to interact with the Microsoft Graph Beta Copilot Chat API.
    Assumes a valid delegated access token and conversation ID are provided upon initialization.
    """
    GRAPH_API_BASE_URL = "https://graph.microsoft.com/beta/copilot"

    def __init__(self, token: str, conversation_id: str):
        # We store necessary context when the plugin is initialized in Python
        self.token = token
        self.conversation_id = conversation_id
        # Define headers used for all requests within this plugin
        self.headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }

    @kernel_function(
        description="Sends a single message prompt to an ongoing Copilot conversation and retrieves the response.",
        name="sendMessageToCopilot"
    )
    def send_message_sync(
        self,
        prompt_text: Annotated[str, "The specific text prompt to send to the M365 Copilot service."]
    ) -> Annotated[str, "The text response received back from the Copilot service."]:
        """
        Sends a single message to an existing conversation via the Graph API sync chat endpoint.
        """
        url = f"{self.GRAPH_API_BASE_URL}/conversations/{self.conversation_id}/chat"
        
        payload = {
            "message": {
                "text": prompt_text
            },
            "locationHint": {
                "timeZone": "UTC" # Changed from 'America/New_York' to generic UTC
            }
        }
        
        try:
            # We use requests.post (synchronous) here as kernel functions are often expected to be sync 
            # unless running in a fully async main loop (which your original code suggested with 'await', but requests library is sync)
            response = requests.post(url, headers=self.headers, data=json.dumps(payload))
            response.raise_for_status() # Raise exception for bad status codes
            
            response_data = response.json()
            
            # The structure of the response might be complex. This attempts to extract the relevant text.
            try:
                copilot_response_text = response_data['messages'][1]['text']
                return copilot_response_text
            except (KeyError, IndexError):
                return f"Error: Could not extract specific message text from response. Full data: {json.dumps(response_data)}"

        except requests.exceptions.RequestException as e:
            # Handle connection or HTTP errors gracefully
            return f"Error connecting to M365 Graph API: {e}"

    @kernel_function(description="Ends the current Copilot conversation when user types exit.")
    def end_conversation(self) -> str:
        """Deletes the conversation resource."""
        url = f"{self.GRAPH_API_BASE_URL}/conversations/{self.conversation_id}"
        try:
            response = requests.delete(url, headers=self.headers)
            response.raise_for_status()
            return f"Conversation {self.conversation_id} successfully ended/deleted."
        except requests.exceptions.RequestException as e:
            return f"Error ending conversation: {e}"

class LocalDocumentPlugin:
    """
    Plugin solely for generating a Word document file content locally in memory.
    """
    def __init__(self):
        # No graph client needed for local generation
        pass

    @kernel_function(description="Generates a Word document file content from given content and returns a success status.")
    def generate_word_document_bytes(
        self,
        filename: Annotated[str, "The name of the document file, e.g., 'MeetingNotes.docx'"],
        content: Annotated[str, "The text content to put into the document"]
    ) -> Annotated[str, "A status message indicating the file content was generated"]:
        """
        Generates a Word file in an in-memory buffer.
        Note: In a real application, you might use the returned buffer object for further processing.
        """
        
        # 1. Create the Word document in memory using python-docx
        document = Document()
        document.add_paragraph(content)
        
        # 2. Save document to an in-memory bytes buffer
        buffer = io.BytesIO()
        document.save(buffer)
        buffer.seek(0) # Reset buffer position to start for reading

        # We store the buffer object locally if we wanted to process it further in Python,
        # but for demonstration via Semantic Kernel return value, we just confirm status.
        # document_bytes_variable = buffer.read() 

        return f"Successfully generated Word document content for '{filename}' in memory (Bytes available for local use)."