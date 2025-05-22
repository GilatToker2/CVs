# from anthropic import AnthropicVertex
#
# LOCATION = "us-east5"
#
# client = AnthropicVertex(region=LOCATION, project_id="kpmg-project-2")
#
# message = client.messages.create(
#     max_tokens=1024,
#     messages=[
#         {
#             "role": "user",
#             "content": "Send me a recipe for banana bread.",
#         }
#     ],
#     model="claude-3-7-sonnet@20250219"
# )
# print(" message", message)
# print(message.model_dump_json(indent=2))

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
        '.pdf': 'application/pdf',
        '.doc': 'application/msword',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.tiff': 'image/tiff',
        '.bmp': 'image/bmp'
    }
    return mime_types.get(file_extension.lower(), 'application/octet-stream')


# def extract_text_with_claude(file_path, prompt_text="Please extract all text from this document."):
#     """
#     Send a file to Claude and get the extracted text.
#
#     Args:
#         file_path: Path to the file to process
#         prompt_text: Text to send to Claude along with the file
#
#     Returns:
#         The text response from Claude
#     """
#     try:
#         # Initialize the client
#         client = AnthropicVertex(region=LOCATION, project_id=PROJECT_ID)
#
#         # Get file extension and MIME type
#         _, file_extension = os.path.splitext(file_path)
#         mime_type = get_mime_type(file_extension)
#
#         # Read the file as binary
#         with open(file_path, 'rb') as file:
#             file_content = file.read()
#
#         # Encode the file content as base64
#         base64_content = base64.b64encode(file_content).decode('utf-8')
#
#         # Create the message with the file content
#         message = client.messages.create(
#             max_tokens=4096,  # Increased token limit for longer responses
#             messages=[
#                 {
#                     "role": "user",
#                     "content": [
#                         {
#                             "type": "text",
#                             "text": prompt_text
#                         },
#                         {
#                             "type": "image",
#                             "source": {
#                                 "type": "base64",
#                                 "media_type": mime_type,
#                                 "data": base64_content
#                             }
#                         }
#                     ]
#                 }
#             ],
#             model=MODEL
#         )
#
#         # Extract and return the text response
#         return message.content[0].text
#
#     except Exception as e:
#         return f"Error processing file with Claude: {str(e)}"


def combine_ocr_results(ocr_results, file_type=None, language=None):
    """
    Combine multiple OCR results into a single, optimized text using Claude.

    Args:
        ocr_results: Dictionary of OCR results from different methods
        file_type: The type of file that was processed (optional)
        language: The detected language of the text (optional)

    Returns:
        The combined and optimized text
    """
    try:
        # Initialize the client
        client = AnthropicVertex(region=LOCATION, project_id=PROJECT_ID)

        # Create a smart prompt that instructs Claude how to combine the results
        prompt = """You are an expert OCR text optimizer. You've been given multiple text extractions from the same document, 
each produced by a different OCR method. Your task is to create a single, optimized version that combines the best elements 
from each extraction.

Follow these guidelines:
1. Prioritize completeness - include all meaningful content from all versions
2. Fix obvious OCR errors (e.g., misrecognized characters, broken words)
3. Maintain the original document structure (paragraphs, lists, tables)
4. Preserve proper names, numbers, and technical terms accurately
5. If versions conflict, prefer the version that appears most grammatically and contextually correct
6. Remove any extraction artifacts or error messages that aren't part of the original document
7. For Hebrew text, ensure proper right-to-left formatting and character recognition
8. Pay special attention to numbers, dates, and other critical information
9. Maintain the original formatting as much as possible (bullet points, numbering, etc.)
10. If a section appears in one version but not others, include it if it seems legitimate

"""

        if file_type:
            prompt += f"\nThis text was extracted from a {file_type.upper()} file. "

        if language:
            prompt += f"\nThe primary language of the document is {language.upper()}. "
            if language.lower() == "hebrew":
                prompt += "Pay special attention to right-to-left text direction and Hebrew characters. "

        prompt += "\n\nHere are the different OCR extractions:\n\n"

        # Add each OCR result to the prompt
        for method, text in ocr_results.items():
            if isinstance(text, str) and not text.startswith("Failed") and not text.startswith("Error"):
                prompt += f"--- {method} ---\n{text}\n\n"

        prompt += "\nPlease provide a single, optimized version of the text that combines the best elements from all extractions. Return ONLY the optimized text without any explanations or notes:"

        # Create the message with the prompt
        message = client.messages.create(
            max_tokens=4096,
            messages=[
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            model=MODEL
        )

        # Extract and return the text response
        return message.content[0].text

    except Exception as e:
        return f"Error combining OCR results with Claude: {str(e)}"


def process_file_with_combined_ocr(file_path, ocr_results=None):
    """
    Process a file with combined OCR methods and return the optimized result.

    Args:
        file_path: Path to the file to process
        ocr_results: Dictionary of OCR results if already available (optional)

    Returns:
        The optimized text from combined OCR methods
    """
    try:
        # Get file type and language if available
        _, file_extension = os.path.splitext(file_path)
        file_type = file_extension.lower().replace('.', '')
        language = None

        # If OCR results are not provided, use the ones from OCR_Combined.py
        if not ocr_results:
            # Import here to avoid circular imports
            from OCR_Combined import process_png_file, process_pdf_file, process_docx_file

            if file_extension.lower() == '.png':
                ocr_results = process_png_file(file_path, use_documentai=True)
            elif file_extension.lower() == '.pdf':
                ocr_results = process_pdf_file(file_path)
            elif file_extension.lower() == '.docx':
                ocr_results = process_docx_file(file_path)
                # Try to get language from the results
                if "Note" in ocr_results and "hebrew" in ocr_results["Note"].lower():
                    language = "hebrew"
                else:
                    language = "english"
            else:
                return f"Unsupported file type: {file_extension}"

        # Try to detect language from the results if not already set
        if not language:
            # Check if any result mentions Hebrew
            for key, value in ocr_results.items():
                if isinstance(value, str) and "hebrew" in value.lower():
                    language = "hebrew"
                    break

            # Default to English if no language detected
            if not language:
                language = "english"

        # Combine the OCR results
        combined_text = combine_ocr_results(ocr_results, file_type, language)

        return combined_text

    except Exception as e:
        return f"Error processing file with combined OCR: {str(e)}"


def process_client_files_folder():
    """Process all files in the client files folder using combined OCR methods."""
    folder_path = "קבצי קוח"
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
        print(f"\n{'=' * 80}")
        print(f"Processing: {file_path.name}")
        print(f"{'=' * 80}")

        # Get file extension
        _, file_extension = os.path.splitext(file_path)
        file_extension = file_extension.lower()

        # Skip unsupported file types
        if file_extension not in ['.png', '.pdf', '.docx']:
            print(f"Skipping unsupported file type: {file_extension}")
            continue

        # Process the file with combined OCR
        result = process_file_with_combined_ocr(file_path)

        # Print the result
        print("\n=== COMBINED OCR RESULT ===\n")
        print(result)

        # Save the result to a file
        output_dir = Path("OCR_Results")
        output_dir.mkdir(exist_ok=True)

        output_file = output_dir / f"{file_path.stem}_combined.txt"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(result)

        print(f"\nSaved combined result to: {output_file}")
        print(f"\n{'=' * 80}\n")


# Main execution
if __name__ == "__main__":
    process_client_files_folder()
