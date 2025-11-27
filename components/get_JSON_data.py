import json

def get_JSON_data():
    with open('data.json') as f:
        data = json.load(f)
        return data