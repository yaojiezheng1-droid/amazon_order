import os
import json

# Directory containing the JSON files
directory = r"c:\Users\Cheng\Desktop\amazon_order\order_generation\json_template"

# Path to the reference JSON file
reference_file = os.path.join(directory, "B10-TJ2-16.json")

# Load the content of the reference JSON file
with open(reference_file, "r", encoding="utf-8") as ref_file:
    reference_content = ref_file.read()

# Iterate through all files in the directory
for filename in os.listdir(directory):
    if filename.endswith(".json") and filename != "B10-TJ2-16.json":
        file_path = os.path.join(directory, filename)
        # Overwrite the file with the reference content
        with open(file_path, "w", encoding="utf-8") as json_file:
            json_file.write(reference_content)

print("All JSON files have been updated to match the reference JSON.")