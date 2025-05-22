from anthropic import AnthropicVertex
import base64
import os
from pathlib import Path

LOCATION = "us-east5"
PROJECT_ID = "kpmg-project-2"
MODEL = "claude-3-7-sonnet@20250219"

def get_mime_type(file_extension):
    """Get the MIME type based on file extension."""
    mime_types = {
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.webp': 'image/webp'
    }
    return mime_types.get(file_extension.lower(), None)

def extract_text_with_claude(file_path, prompt_text="Please extract all text from this document."):
    """
    Send a file to Claude and get the extracted text.

    Args:
        file_path: Path to the file to process
        prompt_text: Text to send to Claude along with the file

    Returns:
        The text response from Claude
    """
    try:
        # Initialize the client
        client = AnthropicVertex(region=LOCATION, project_id=PROJECT_ID)

        # Get file extension and MIME type
        _, file_extension = os.path.splitext(file_path)
        mime_type = get_mime_type(file_extension)

        if not mime_type:
            return f"Unsupported file type for Claude: {file_extension}"

        # Read the file as binary
        with open(file_path, 'rb') as file:
            file_content = file.read()

        # Encode the file content as base64
        base64_content = base64.b64encode(file_content).decode('utf-8')

        # Create the message with the file content
        message = client.messages.create(
            max_tokens=4096,  # Increased token limit for longer responses
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": prompt_text
                        },
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": mime_type,
                                "data": base64_content
                            }
                        }
                    ]
                }
            ],
            model=MODEL
        )

        # Extract and return the text response
        return message.content[0].text

    except Exception as e:
        return f"Error processing file with Claude: {str(e)}"


def process_client_files_folder(folder_path):
    """Process all files in the client files folder using Claude."""
    folder = Path(folder_path)

    if not folder.exists():
        print(f"Folder '{folder_path}' does not exist.")
        return

    files = list(folder.iterdir())

    if not files:
        print(f"No files found in '{folder_path}'.")
        return

    print(f"Found {len(files)} files in '{folder_path}'.")

    for file_path in files:
        if not file_path.is_file():
            continue

        print(f"\n{'=' * 50}")
        print(f"Processing: {file_path.name}")
        print(f"{'=' * 50}")

        # Get file extension
        _, file_extension = os.path.splitext(file_path)

        # Check if file type is supported by Claude
        if get_mime_type(file_extension):
            # Extract text using Claude
            text = extract_text_with_claude(file_path)
            # Print the extracted text
            print(text)
        else:
            print(f"Skipping unsupported file type: {file_extension}")

        print(f"{'=' * 50}\n")


if __name__ == "__main__":
    # Process the client files folder
    client_files_folder = "קבצי קוח"
    process_client_files_folder(client_files_folder)
