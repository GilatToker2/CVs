"""
Process Files Module

This is the main controller module that orchestrates the entire workflow:
1. Processes files in the directory
2. Extracts text using OCR
3. Extracts aspects using Claude
4. Saves results to JSON files
5. Updates Excel with the extracted data
"""

import os
import json
import pandas as pd
import re
from pathlib import Path
import datetime

# Import functions from other modules
from excel_to_concepts import excel_to_concept_mapping, create_empty_workbook
from Claude_as_OCR import process_file_with_combined_ocr
from Claude_extract_values import extract_aspects_with_claude


def extract_file_id(filename):
    """
    Extract the ID number from a filename.

    Args:
        filename: The filename to extract the ID from

    Returns:
        The extracted ID, or None if no ID was found
    """
    match = re.search(r'(\d+)', filename)
    if match:
        return match.group(1)
    return None


def process_single_file(file_path, concept_mapping, output_dir="נתונים_שחולצו"):
    """
    Process a single file: perform OCR, extract aspects, and save results.

    Args:
        file_path: Path to the file to process
        concept_mapping: Dictionary of concept mappings from excel_to_concepts
        output_dir: Directory to save the JSON results

    Returns:
        Dictionary with extracted aspects and file ID
    """
    try:
        print(f"Processing file: {file_path}")

        # Get file extension and ID
        file_name = os.path.basename(file_path)
        file_id = extract_file_id(file_name)
        _, file_extension = os.path.splitext(file_path)
        file_extension = file_extension.lower()

        # Skip unsupported file types
        if file_extension not in ['.png', '.pdf', '.docx']:
            print(f"Skipping unsupported file type: {file_extension}")
            return None

        if not file_id:
            print(f"Could not extract ID from file: {file_name}")
            return None

        # Process the file with OCR
        print(f"Extracting text with OCR...")
        ocr_text = process_file_with_combined_ocr(file_path)

        # Get all concept values from all sheets
        all_concept_values = []
        for sheet, concepts in concept_mapping.items():
            for concept in concepts:
                if "value" in concept and concept["value"] not in all_concept_values:
                    all_concept_values.append(concept["value"])

        # Extract aspects from the OCR text
        print(f"Extracting aspects with Claude...")
        extracted_data = extract_aspects_with_claude(ocr_text, all_concept_values, document_type=file_extension[1:])

        # Add metadata
        extracted_data["מטא-נתונים"] = {
            "שם_קובץ": file_name,
            "סוג_קובץ": file_extension[1:],
            "תאריך_עיבוד": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        # Create output directory if it doesn't exist
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)

        # Save the extracted data to a JSON file
        json_path = output_path / f"{Path(file_path).stem}_aspects.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(extracted_data, f, ensure_ascii=False, indent=2)

        print(f"Extracted data saved to {json_path}")

        return {"id": file_id, "data": extracted_data, "file_type": file_extension[1:]}

    except Exception as e:
        print(f"Error processing file {file_path}: {str(e)}")
        return None


def update_excel_with_json_data(json_data, concept_mapping, output_excel_path):
    """
    Update Excel file with JSON data.

    Args:
        json_data: Dictionary with extracted data and file ID
        concept_mapping: Dictionary of concept mappings from excel_to_concepts
        output_excel_path: Path to the output Excel file

    Returns:
        True if successful, False otherwise
    """
    try:
        file_id = json_data["id"]
        extracted_data = json_data["data"]
        file_type = json_data["file_type"]

        # Check if the file exists
        file_exists = os.path.isfile(output_excel_path)

        if file_exists:
            # If file exists, read existing data
            existing_data = {}
            with pd.ExcelFile(output_excel_path) as xls:
                for sheet_name in xls.sheet_names:
                    existing_data[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)

            # Update each sheet
            for sheet_name, concepts in concept_mapping.items():
                if sheet_name not in existing_data:
                    # If sheet doesn't exist in the Excel file, create it
                    existing_data[sheet_name] = pd.DataFrame(columns=["ID", "File Type"])

                # Create a row for this file
                row = {"ID": file_id, "File Type": file_type}

                # Add JSON data to the row
                for concept in concepts:
                    concept_value = concept["value"]
                    if concept_value in extracted_data:
                        row[concept_value] = extracted_data[concept_value]
                    else:
                        row[concept_value] = "Not found"

                # Check if this ID already exists in the sheet
                if 'ID' in existing_data[sheet_name].columns:
                    existing_ids = existing_data[sheet_name]['ID'].astype(str).tolist()

                    if file_id in existing_ids:
                        # Update the existing row
                        idx = existing_data[sheet_name].index[
                            existing_data[sheet_name]['ID'].astype(str) == file_id].tolist()[0]
                        for key, value in row.items():
                            if key in existing_data[sheet_name].columns:
                                existing_data[sheet_name].at[idx, key] = value
                    else:
                        # Add a new row
                        existing_data[sheet_name] = pd.concat([existing_data[sheet_name], pd.DataFrame([row])],
                                                              ignore_index=True)
                else:
                    # Add a new row
                    existing_data[sheet_name] = pd.concat([existing_data[sheet_name], pd.DataFrame([row])],
                                                          ignore_index=True)

            # Write the updated data back to the Excel file
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                for sheet_name, df in existing_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            # If file doesn't exist, create a new one with empty sheets
            create_empty_workbook(concept_mapping.keys(), output_excel_path)

            # Create a dictionary to store data for each sheet
            sheet_data = {}

            # Add data to each sheet
            for sheet_name, concepts in concept_mapping.items():
                # Create a row for this file
                row = {"ID": file_id, "File Type": file_type}

                # Add JSON data to the row
                for concept in concepts:
                    concept_value = concept["value"]
                    if concept_value in extracted_data:
                        row[concept_value] = extracted_data[concept_value]
                    else:
                        row[concept_value] = "Not found"

                # Add the row to the sheet data
                sheet_data[sheet_name] = pd.DataFrame([row])

            # Write the data to the Excel file
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                for sheet_name, df in sheet_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Excel file updated: {output_excel_path}")
        return True

    except Exception as e:
        print(f"Error updating Excel file: {str(e)}")
        return False


def process_client_files(client_files_dir="קבצי קוח", excel_template_path="נתוני דורשי עבודה בינה והשמה - רשימות.xlsx",
                         output_dir="נתונים_שחולצו", output_excel_path="extracted_data.xlsx"):
    """
    Process all client files in a directory.

    Args:
        client_files_dir: Directory containing client files
        excel_template_path: Path to the Excel template file
        output_dir: Directory to save the JSON results
        output_excel_path: Path to the output Excel file

    Returns:
        Number of successfully processed files
    """
    try:
        # Get concept mapping from Excel template
        concept_mapping = excel_to_concept_mapping(excel_template_path)

        # Create output directory if it doesn't exist
        Path(output_dir).mkdir(exist_ok=True)

        # Find all files in the directory
        files = list(Path(client_files_dir).glob("*.*"))
        files = [f for f in files if f.is_file() and not f.name.startswith('~$')]  # Skip temporary files

        if not files:
            print(f"No files found in {client_files_dir}")
            return 0

        print(f"Found {len(files)} files in {client_files_dir}")

        # Try to use tqdm for progress bar
        try:
            from tqdm import tqdm
            files_iter = tqdm(files)
        except ImportError:
            files_iter = files

        # Process each file
        successful_files = 0
        for file_path in files_iter:
            # Process the file
            json_data = process_single_file(file_path, concept_mapping, output_dir)

            if json_data:
                # Update Excel with the extracted data
                if update_excel_with_json_data(json_data, concept_mapping, output_excel_path):
                    successful_files += 1

        print(f"Successfully processed {successful_files} out of {len(files)} files")
        return successful_files

    except Exception as e:
        print(f"Error processing client files: {str(e)}")
        return 0


def main():
    """Main function to run the script."""
    import argparse

    parser = argparse.ArgumentParser(description='Process client files and extract information')
    parser.add_argument('--client_files_dir', default="קבצי קוח",
                        help='Directory containing client files')
    parser.add_argument('--excel_template', default="נתוני דורשי עבודה בינה והשמה - רשימות.xlsx",
                        help='Path to Excel template file')
    parser.add_argument('--output_dir', default="נתונים_שחולצו",
                        help='Directory to save the JSON results')
    parser.add_argument('--output_excel', default="extracted_data.xlsx",
                        help='Path to the output Excel file')
    parser.add_argument('--single_file', default=None,
                        help='Process a single file instead of the entire directory')

    args = parser.parse_args()

    if args.single_file:
        # Process a single file
        concept_mapping = excel_to_concept_mapping(args.excel_template)
        json_data = process_single_file(args.single_file, concept_mapping, args.output_dir)
        if json_data:
            update_excel_with_json_data(json_data, concept_mapping, args.output_excel)
    else:
        # Process all files in the directory
        process_client_files(args.client_files_dir, args.excel_template, args.output_dir, args.output_excel)


if __name__ == "__main__":
    main()
