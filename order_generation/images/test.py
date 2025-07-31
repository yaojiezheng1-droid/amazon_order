import os
import pandas as pd
import requests
from tqdm import tqdm  # Optional: for progress bar

# Define file paths
excel_file = r'c:\Users\Cheng\Desktop\amazon_order\order_generation\images\导出产品-按SKU-809420939716304896.xlsx'  # Replace with your Excel file name
output_folder = r'c:\Users\Cheng\Desktop\amazon_order\order_generation\images\products'  # Replace with your desired output folder

# Create the output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Load the Excel file
df = pd.read_excel(excel_file)

# Check if the required columns exist
if "图片链接" not in df.columns or "*SKU" not in df.columns:
    raise ValueError("The columns '图片链接' and/or '*SKU' are not found in the Excel file.")

# Download each image
for index, row in tqdm(df.dropna(subset=["图片链接", "*SKU"]).iterrows(), desc="Downloading images"):
    url = row["图片链接"]
    sku = row["*SKU"]
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()  # Raise an error for bad status codes

        # Generate a file name based on the SKU
        file_name = os.path.join(output_folder, f'{sku}.jpg')

        # Save the image
        with open(file_name, 'wb') as file:
            for chunk in response.iter_content(1024):
                file.write(chunk)

    except requests.RequestException as e:
        print(f"Failed to download {url} for SKU {sku}: {e}")

print("Download complete!")