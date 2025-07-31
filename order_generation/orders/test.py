import json
import os

# Paths
mapping_path = r'c:\Users\Cheng\Desktop\amazon_order\order_generation\docs\complete_mapping_with_po.json'
orders_folder = r'c:\Users\Cheng\Desktop\amazon_order\order_generation\orders'

# 1. Collect all purchase_order_file values from the JSON
def collect_purchase_order_files(data):
    files = set()
    def walk(obj):
        if isinstance(obj, dict):
            for k, v in obj.items():
                if k == "purchase_order_file":
                    files.add(str(v))
                else:
                    walk(v)
        elif isinstance(obj, list):
            for item in obj:
                walk(item)
    walk(data)
    return files

with open(mapping_path, encoding='utf-8') as f:
    data = json.load(f)
purchase_order_files = collect_purchase_order_files(data)

# 2. List all files in orders/all_orders
for filename in os.listdir(orders_folder):
    file_path = os.path.join(orders_folder, filename)
    # 3. Delete files not in purchase_order_files
    if filename not in purchase_order_files:
        os.remove(file_path)
        print(f"Deleted: {filename}")