"""
   Convert json files to excel files
"""

import json
import pandas as pd


def json_to_excel(json_file, excel_file):
    # Read JSON file into a pandas DataFrame
    df = pd.read_json(json_file)

    # Write DataFrame to Excel file
    df.to_excel(excel_file, index=False)

# Example usage
xml_file = "C:\Codes\etap\Example.json"  # Your json file
excel_file = "C:\Codes\etap\output.xlsx"  # Output Excel file
json_to_excel(xml_file, excel_file)