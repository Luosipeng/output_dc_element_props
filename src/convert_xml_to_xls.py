"""
    Convert the xml file to xlxs format using pandas
"""

import pandas as pd

def xml_to_excel(xml_file, excel_file):
    data = pd.read_xml(xml_file, parser='etree')
    data.to_excel(excel_file)

# Example usage
xml_file = 'C:\Codes\etap\Feeder.xml'  # Your XML file
excel_file = 'C:\Codes\etap\output.xlsx'  # Output Excel file
xml_to_excel(xml_file, excel_file)


