# ===========================================
# This function create data.json by
# extracting data from excel file
# You need to give this function a path to
# the file in string format
# ===========================================

import pandas as pd
import json

def convert_excel_to_JSON(path):
    # Read the Excel file
    excel_data_df = pd.read_excel(path)

    # Convert Excel to JSON records (up-to-down orientation)
    thisisjson = excel_data_df.to_json(orient='records')

    # Load the JSON string back into a dictionary/list structure
    thisisjson_dict = json.loads(thisisjson)

    # Write the JSON object to a file while preserving Unicode characters
    with open('data.json', 'w', encoding='utf-8') as json_file:
        json.dump(thisisjson_dict, json_file, indent=4, ensure_ascii=False)

    # Optionally print the final result
    print(json.dumps(thisisjson_dict, indent=4, ensure_ascii=False))