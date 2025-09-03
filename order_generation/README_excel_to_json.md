# Excel to JSON Template Converter

This script converts Excel files in the format of `empty_base_template.xlsx` to JSON template files suitable for the `json_template` folder.

## Features

- **Automatic Cell Mapping**: Extracts all standard template cells (supplier info, dates, notes, etc.)
- **Product Detection**: Automatically finds and extracts product table data
- **Image Path Resolution**: Attempts to find product images in the images directory
- **Duplicate Handling**: Combines quantities for duplicate SKUs
- **Product Name Lookup**: Uses accessory mapping to populate product names
- **Batch Processing**: Can process multiple Excel files at once

## Usage

### Command Line
```bash
# Single file
python excel_to_json_template.py order_file.xlsx

# Multiple files
python excel_to_json_template.py file1.xlsx file2.xlsx

# Glob pattern
python excel_to_json_template.py docs/*.xlsx
```

### Batch File (Windows)
```bash
# Single file
convert_excel_to_json.bat order_file.xlsx

# Multiple files  
convert_excel_to_json.bat docs/*.xlsx
```

## Excel File Requirements

The Excel file should follow the `empty_base_template.xlsx` format:

### Required Structure:
- **Row 3**: Supplier info (B3) and Order number (G3)
- **Row 4**: Phone (B4) and Date (G4)  
- **Row 5**: Contact (B5) and Order arranger (G5)
- **Row 7+**: Product table with headers:
  - Column A: 产品编号 (Product Code/SKU)
  - Column B: 产品图片 (Product Image)
  - Column C: 描述 (Description)
  - Column D: 数量/个 (Quantity)
  - Column E: 单价 (Unit Price)
  - Column G: 包装方式 (Packaging)

### Optional Fields:
- **Row 12**: Warehouse address
- **Row 13**: Payment terms
- **Row 14**: Delivery time, color cards, logos
- **Rows 19-30**: Notes and special instructions
- **Row 69**: Buyer (B69) and Supplier (E69) info

## Output

For each unique SKU in the Excel file, the script generates a JSON template file named `{SKU}.json` in the `json_template` directory.

### Generated JSON Structure:
```json
{
  "cells": {
    "B3": {"key": "供货商：", "value": "Supplier Name"},
    "G4": {"key": "日期", "value": "2025-09-02"},
    // ... all other cell mappings
  },
  "products": [
    {
      "产品编号": "SKU123",
      "产品名称": "Product Name",
      "产品图片": "order_generation/images/products/SKU123.jpg",
      "描述": "Product description",
      "数量/个": 1000,
      "单价": 5.5,
      "包装方式": "Box packaging"
    }
  ],
  "footer": {
    "buyer": "Buyer Company",
    "supplier": "Supplier Company"
  }
}
```

## Integration

The generated JSON templates can be used with:

1. **direct_sku_to_json.py**: Generate factory orders from SKUs
2. **product_search_gui.py**: Visual product search and order generation
3. **json_PO_excel.py**: Convert back to Excel purchase orders

## Error Handling

- **Missing Products**: Warning if no products found
- **Invalid Data**: Gracefully handles missing or invalid cell values
- **Duplicate SKUs**: Automatically combines quantities
- **Missing Images**: Uses provided path or leaves empty

## Examples

### Convert a single order file:
```bash
python excel_to_json_template.py my_order.xlsx
```

### Convert all Excel files in docs folder:
```bash
python excel_to_json_template.py docs/*.xlsx
```

### Using the batch file:
```bash
convert_excel_to_json.bat "Order - Product ABC.xlsx"
```
