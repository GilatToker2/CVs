import os
import sys
import subprocess
import tempfile
from pathlib import Path


# def install_required_packages():
#     """Install required packages if they're not already installed."""
#     required_packages = [
#         'PyPDF2',
#         'python-docx',
#         'Pillow',
#         'pytesseract',
#         'pywin32'  # For doc to docx conversion
#     ]
#
#     for package in required_packages:
#         try:
#             __import__(package.replace('-', '_').split('[')[0])
#             print(f"{package} is already installed.")
#         except ImportError:
#             print(f"Installing {package}...")
#             subprocess.check_call([sys.executable, "-m", "pip", "install", package])
#             print(f"{package} has been installed.")


def extract_text_from_pdf(file_path):
    """Extract text from a PDF file."""
    import PyPDF2

    text = ""
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text += page.extract_text() + "\n"
    except Exception as e:
        text = f"Error extracting text from PDF: {str(e)}"

    print("File contents pdf:", text)

    return text


def extract_text_from_docx(file_path):
    """Extract text from a DOCX file."""
    import docx

    text = ""
    try:
        doc = docx.Document(file_path)
        for para in doc.paragraphs:
            text += para.text + "\n"
    except Exception as e:
        text = f"Error extracting text from DOCX: {str(e)}"

    return text


def convert_doc_to_docx(doc_path):
    """Convert a .doc file to .docx format using pywin32."""
    try:
        import win32com.client
        import os

        # Create a temporary file path for the docx file
        docx_path = os.path.splitext(doc_path)[0] + "_converted.docx"

        # Initialize Word application
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        # Open the doc file
        doc = word.Documents.Open(os.path.abspath(doc_path))

        # Save as docx
        doc.SaveAs(os.path.abspath(docx_path), 16)  # 16 represents the wdFormatDocumentDefault (docx)
        doc.Close()
        word.Quit()

        print(f"Successfully converted {doc_path} to {docx_path}")
        return docx_path
    except Exception as e:
        print(f"Error converting DOC to DOCX: {str(e)}")
        return None


def extract_text_from_doc(file_path):
    """Extract text from a DOC file by first converting it to DOCX."""
    try:
        # Convert the DOC file to DOCX
        docx_path = convert_doc_to_docx(file_path)

        if docx_path and os.path.exists(docx_path):
            # Extract text from the converted DOCX file
            text = extract_text_from_docx(docx_path)

            # Clean up the temporary DOCX file
            try:
                os.remove(docx_path)
                print(f"Removed temporary file: {docx_path}")
            except:
                pass

            return text
        else:
            return f"Failed to convert DOC file: {file_path}\nConversion to DOCX was unsuccessful."
    except Exception as e:
        return f"Error processing DOC file: {str(e)}"


def extract_text_from_image(file_path):
    """Extract text from an image file using OCR."""
    from PIL import Image
    import pytesseract

    text = ""
    try:
        image = Image.open(file_path)
        text = pytesseract.image_to_string(image)
    except Exception as e:
        text = f"Error extracting text from image: {str(e)}"

    return text


def process_resume_folder(folder_path):
    """Process all files in the resume folder."""
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

        extension = file_path.suffix.lower()

        if extension == '.pdf':
            text = extract_text_from_pdf(file_path)
        elif extension == '.docx':
            text = extract_text_from_docx(file_path)
        elif extension == '.doc':
            text = extract_text_from_doc(file_path)
        elif extension in ['.png', '.jpg', '.jpeg', '.tiff', '.bmp']:
            text = extract_text_from_image(file_path)
        else:
            text = f"Unsupported file type: {extension}"

        print(text)
        print(f"{'=' * 50}\n")


if __name__ == "__main__":
    # # Install required packages
    # install_required_packages()

    # Process the resume folder
    resume_folder = "קבצי קוח"
    process_resume_folder(resume_folder)
