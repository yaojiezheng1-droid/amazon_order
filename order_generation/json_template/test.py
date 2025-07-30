import json
from pathlib import Path

# Path to the folder containing JSON files
JSON_FOLDER = Path("c:/Users/Cheng/Desktop/amazon_order/order_generation/json_template")

# Code snippet to identify files to replace
target_code = {
    "A24": {
        "key": "5：注意事项",
        "value": ""
    },
    "B24": {
        "key": "以上价格含税运含包装",
        "value": ""
    },
    "A25": {
        "key": "5：注意事项",
        "value": ""
    },
    "B25": {
        "key": "以上价格含税运含包装",
        "value": ""
    },
    "A26": {
        "key": "5：注意事项",
        "value": ""
    },
    "B26": {
        "key": "以上价格含税运含包装",
        "value": ""
    }
}

# Load the content of the reference JSON file
reference_file = JSON_FOLDER / "EC403-2.json"
with open(reference_file, "r", encoding="utf-8") as ref_file:
    reference_content = json.load(ref_file)

def replace_json_content():
    for json_file in JSON_FOLDER.glob("*.json"):
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # Check if the target code exists in the "cells" section
        if "cells" in data and all(item in data["cells"].items() for item in target_code.items()):
            # Replace the content with the reference JSON content
            with open(json_file, "w", encoding="utf-8") as f:
                json.dump(reference_content, f, ensure_ascii=False, indent=2)
            print(f"Replaced content in {json_file}")

if __name__ == "__main__":
    replace_json_content()