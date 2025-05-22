import os
import sys
import subprocess
from pathlib import Path
from google.cloud import documentai_v1 as documentai
from google.api_core.client_options import ClientOptions

# project_id = "kpmg-project-2"
# location   = "us"   # or wherever your processor lives
#
# # point the client at the right regional endpoint
# opts   = ClientOptions(api_endpoint=f"{location}-documentai.googleapis.com")
# client = documentai.DocumentProcessorServiceClient(client_options=opts)
#
# parent     = client.common_location_path(project_id, location)
# processors = client.list_processors(parent=parent)
#
# for p in processors:
#     # p.name looks like:
#     # projects/kpmg-project-2/locations/us/processors/ₓₓₓₓₓₓₓₓₓₓ
#     print("→", p.name)0

PROJECT_ID   = "kpmg-project-2"
LOCATION     = "us"
PROCESSOR_ID = "29819412e88bdf7e"


def get_mime_type(file_extension):
    """Get the MIME type based on file extension."""
    mime_types = {
        '.pdf': 'application/pdf',
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.tiff': 'image/tiff',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.doc': 'application/msword'
    }
    return mime_types.get(file_extension.lower(), None)


def extract_text_with_documentai(file_path):
    """
    Extract text from a file using Google Cloud Document AI.

    Args:
        file_path: Path to the file to process

    Returns:
        The extracted text
    """
    try:
        # Get file extension and MIME type
        _, file_extension = os.path.splitext(file_path)
        mime_type = get_mime_type(file_extension)

        if not mime_type:
            return f"Unsupported file type for Document AI: {file_extension}"


        # Read the file as binary
        with open(file_path, "rb") as f:
            file_content = f.read()

        # Create Document AI client
        client = documentai.DocumentProcessorServiceClient()
        name = client.processor_path(PROJECT_ID, LOCATION, PROCESSOR_ID)

        # Prepare the request
        document = {"content": file_content, "mime_type": mime_type}
        request = {"name": name, "raw_document": document}

        # Send the request
        result = client.process_document(request=request)

        # Extract text from the result
        document_object = result.document
        return document_object.text

    except Exception as e:
        return f"Error processing file with Document AI: {str(e)}"


def process_client_files_folder(folder_path):
    """Process all files in the client files folder using Document AI."""
    folder = Path(folder_path)

    if not folder.exists():
        print(f"Folder '{folder_path}' does not exist.")
        return

    files = list(folder.iterdir())
    files = [f for f in files if f.is_file() and not f.name.startswith('~$')]  # Skip temporary files

    if not files:
        print(f"No files found in '{folder_path}'.")
        return

    print(f"Found {len(files)} files in '{folder_path}'.")

    # Try to import tqdm for progress bar
    try:
        from tqdm import tqdm
        files_iter = tqdm(files)
    except ImportError:
        files_iter = files

    for file_path in files_iter:
        print(f"\n{'=' * 50}")
        print(f"Processing: {file_path.name}")
        print(f"{'=' * 50}")

        # Get file extension
        _, file_extension = os.path.splitext(file_path)

        # Check if file type is supported
        if get_mime_type(file_extension):
            # Extract text using Document AI
            text = extract_text_with_documentai(file_path)
            # Print the extracted text
            print(text)
        else:
            print(f"Skipping unsupported file type: {file_extension}")

        print(f"{'=' * 50}\n")


if __name__ == "__main__":
    # Process the client files folder
    client_files_folder = "קבצי קוח"
    process_client_files_folder(client_files_folder)
