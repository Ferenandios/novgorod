import json
import pandas as pd
from openpyxl import Workbook


def JSON_to_Excel_table(json_data):
    """
    Converts a list of dictionaries in JSON format into an Excel workbook.
    
    :param json_data: List of dictionaries containing component information
    :return: None (saves output directly to Excel file)
    """
    try:
        df = pd.DataFrame(json_data)
        
        # Create a new Excel workbook using OpenPyXL
        wb = Workbook()
        ws = wb.active
        
        # Write column headers from DataFrame columns
        ws.append(df.columns.tolist())
        
        # Append each row's data to the worksheet
        for index, row in df.iterrows():
            ws.append(row.values.tolist())
            
        # Save the workbook to disk
        wb.save('output.xlsx')
        print("Data successfully converted to Excel.")
    except Exception as e:
        print(f"Error converting JSON to Excel: {e}")