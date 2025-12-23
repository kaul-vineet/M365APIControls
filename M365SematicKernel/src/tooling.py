import io, urllib
import json
import requests
from typing import Annotated, Optional 
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
        description="Sends a prompt to the M365 Copilot and waits for the response. Use this to get content, answers, or summaries that will be used in subsequent steps (like creating a document).",
        name="sendMessageToCopilot"
    )
    def send_message_sync(
        self,
        prompt_text: Annotated[str, "The specific text prompt to send to the M365 Copilot service."]
    ) -> Annotated[str, "The text response from Copilot. You MUST capture this output to use as the 'content' argument for document generation tools."]:
        """
        Sends a single message to an existing conversation via the Graph API sync chat endpoint.
        """
        url = f"{self.GRAPH_API_BASE_URL}/conversations/{self.conversation_id}/chat"
        
        payload = {
            "message": {
                "text": prompt_text
            },
            "additionalContext": [{
                "text": "Respond in high class British English used by gentlemen of the 18th century. " # Changed from 'America/New_York' to generic UTC
            }],
            "locationHint": {
                "timeZone": "America/New_York"
            }
        }

        try:
            # We use requests.post (synchronous) here as kernel functions are often expected to be sync 
            # unless running in a fully async main loop (which your original code suggested with 'await', but requests library is sync)
            response = requests.post(url, headers=self.headers, data=json.dumps(payload))
            #response.raise_for_status() # Raise exception for bad status codes
            response_data = response.json()
            print(response_data)
            # The structure of the response might be complex. This attempts to extract the relevant text.
            try:
                copilot_response_text = response_data['messages'][1]['text']
                return copilot_response_text
            except (KeyError, IndexError):
                return f"Error: Could not extract specific message text from response. Full data: {json.dumps(response_data)}"

        except requests.exceptions.RequestException as e:
            # Handle connection or HTTP errors gracefully
            return f"Error connecting to M365 Graph API: {e}"

    @kernel_function(description="Terminates the current session and deletes conversation context. Call this ONLY when the user explicitly says 'exit', 'quit', or 'goodbye'.")
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

    @kernel_function(description="Creates a Word document file (.docx) in memory containing the provided text.")
    def generate_word_document_bytes(
        self,
        filename: Annotated[str, "The name of the file to create (e.g., 'ProjectSummary.docx')."],
        content: Annotated[str, "The specific text to write into the document. **IMPORTANT:** This value MUST be the exact output string received from the `sendMessageToCopilot` tool. Do NOT ask the user for this content; pipe the previous result directly here."]
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
        self.document_buffer = buffer 
        self.filename = filename
        # We store the buffer object locally if we wanted to process it further in Python,
        # but for demonstration via Semantic Kernel return value, we just confirm status.
        # document_bytes_variable = buffer.read() 

        return f"Successfully generated Word document content for '{filename}' in memory (Bytes available for local use)."

class GraphSharePointUploaderPlugin:
    """
    Plugin for uploading in-memory bytes to SharePoint using the Microsoft Graph API and an Access Token.
    Configured using Site URL and Library Name.
    """
    def __init__(self, access_token: str, site_url: str, target_folder_path: str, generator_plugin: LocalDocumentPlugin):
        self.access_token = access_token

        # Store a direct reference to the other plugin's instance
        self.generator_plugin_ref = generator_plugin 

        # Construct the correct base URL
        self.base_url = f"{site_url}{target_folder_path}/"

    @kernel_function(description="Uploads the previously generated in-memory Word document to SharePoint. This tool MUST be called immediately after `generate_word_document_bytes`.")
    def upload_generated_file(
        self,
        target_folder_path: Annotated[str, "The destination folder path in the SharePoint library (e.g., 'Reports' or empty string per user request)."]
    ) -> Annotated[str, "The WebUrl of the newly uploaded file, or an error message."]:
        
        # --- THIS IS WHERE WE ACCESS THE WORD DOCUMENT VARIABLE ---
        if self.generator_plugin_ref.document_buffer is None:
            return "Error: No file content found in the associated generator plugin's buffer."
            
        file_content_bytes = self.generator_plugin_ref.document_buffer.read()
        filename = self.generator_plugin_ref.filename
        # --------------------------------------------------------

        if not file_content_bytes:
             return "Error: Buffer was empty."

        # Construct the final URL with folder path and file name
        endpoint_url = f"{self.base_url}{filename}:/content"
        print(f"SharePoint Base URL: {endpoint_url}")

        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' # MIME type for .docx
        }

        try:
            response = requests.put(url=endpoint_url, data=file_content_bytes, headers=headers)
            response.raise_for_status() 
            uploaded_item_info = response.json()
            return f"Success: Uploaded to {uploaded_item_info['webUrl']}"

        except requests.exceptions.HTTPError as e:
            return f"Error uploading via Graph API (Status {response.status_code}): {response.text}"