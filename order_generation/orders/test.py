import os
import pandas as pd

# Define paths
orders_folder = r'C:\Users\Cheng\Desktop\amazon_order\order_generation\orders'
采购单_file = r'C:\Users\Cheng\Desktop\amazon_order\order_generation\docs\采购单_modified.xlsx'

# Read the "order" column from 采购单modified.xlsx
df = pd.read_excel(采购单_file)
valid_orders = df['order'].astype(str).str.strip().str.lower().tolist()  # Normalize valid orders
print("Valid orders loaded:", valid_orders)

# Iterate through files in the orders folder
for file_name in os.listdir(orders_folder):
    file_path = os.path.join(orders_folder, file_name)
    
    # Check if the file is an Excel file
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        # Extract the order name (assuming the file name matches the order name)
        order_name = os.path.splitext(file_name)[0].strip().lower()  # Normalize file name
        
        # Remove the file if it's not in the valid orders list
        if order_name not in valid_orders:
            os.remove(file_path)
            print(f"Removed: {file_path}")

print("Cleanup complete.")