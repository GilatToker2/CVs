"""
Excel Concepts Module

This module provides functions for working with Excel files containing concept mappings.
It can read Excel files and extract concept mappings, as well as create empty Excel workbooks.
"""

import pandas as pd
import json
from openpyxl import Workbook


def excel_to_concept_mapping(xlsx_path: str) -> dict:
    """
    Reads every sheet in the Excel file and returns a dict of concept mappings.

    Args:
        xlsx_path: Path to the Excel file

    Returns:
        Dictionary mapping sheet names to lists of concept dictionaries:
        { sheet_name: [ { "id": original_id, "value": text }, ... ] }
    """
    try:
        xls = pd.ExcelFile(xlsx_path, engine="openpyxl")
        mapping = {}

        for sheet in xls.sheet_names:
            # Read with no header so row0 is the sheet header
            df = xls.parse(sheet, header=None)

            # Drop any fully-empty rows and reset index
            df = df.dropna(how='all').reset_index(drop=True)

            # If there's fewer than 2 rows (no data), skip this sheet
            if df.shape[0] < 2:
                print(f"Skipping empty sheet: {sheet}")
                continue

            # Skip the header row, keep only the real data rows
            data = df.iloc[1:].reset_index(drop=True)

            # Extract IDs and values
            if data.shape[1] >= 2:
                raw_ids = data.iloc[:, 0]  # e.g. 0,1,2...
                raw_vals = data.iloc[:, 1]  # e.g. "עברית","רוסית",...
            else:
                # Fallback: use row-index as id and single column as value
                raw_ids = data.index
                raw_vals = data.iloc[:, 0]

            # Try to cast IDs to int; if that fails, keep as-is
            try:
                ids = raw_ids.astype(int).tolist()
            except:
                ids = raw_ids.tolist()

            # Clean up values: drop blanks, cast to str, strip whitespace
            vals = raw_vals.dropna().astype(str).str.strip().tolist()

            # Zip into list of {"id":..., "value":...}
            pairs = [
                {"id": ids[i], "value": vals[i]}
                for i in range(min(len(ids), len(vals)))
            ]
            mapping[sheet] = pairs

        # Remove any sheets that ended up with empty lists
        mapping = {sheet: values for sheet, values in mapping.items() if values}

        return mapping
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        return {}

def create_empty_workbook(sheet_names, output_path: str = "empty_structure.xlsx"):
    """
    Creates a new Excel workbook with empty sheets.

    Args:
        sheet_names: Iterable of sheet names to create
        output_path: Path to save the Excel file

    Returns:
        True if successful, False otherwise
    """
    try:
        wb = Workbook()

        # Remove default sheet
        default = wb.active
        wb.remove(default)

        # Create each sheet
        for name in sheet_names:
            wb.create_sheet(title=name)

        wb.save(output_path)
        print(f"Empty workbook created: {output_path}")
        return True
    except Exception as e:
        print(f"Error creating empty workbook: {str(e)}")
        return False


# Example usage
if __name__ == "__main__":
    xlsx_path = "נתוני דורשי עבודה בינה והשמה - רשימות.xlsx"
    mapping = excel_to_concept_mapping(xlsx_path)
    print(json.dumps(mapping, ensure_ascii=False, indent=2))

    # Create an empty workbook with the same sheets
    create_empty_workbook(mapping.keys(), "empty_structure.xlsx")
