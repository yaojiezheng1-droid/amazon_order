import json
from pathlib import Path

# Path to the folder containing JSON templates
TEMPLATE_DIR = Path("order_generation/json_template")

# Load the reference JSON file
with open(TEMPLATE_DIR / "2EC-1-1.json", "r", encoding="utf-8") as ref_file:
    reference_data = json.load(ref_file)

def update_json_files():
    for json_file in TEMPLATE_DIR.glob("*.json"):
        if json_file.name == "2EC-1-1.json":
            continue  # Skip the reference file
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # Update the structure to match the reference file
        data["cells"] = reference_data["cells"]
        data["products"] = reference_data["products"]
        data["footer"] = reference_data["footer"]
        
        # Save the updated JSON file
        with open(json_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"Updated {json_file}")

if __name__ == "__main__":
    update_json_files()