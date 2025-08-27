#!/usr/bin/env python3
"""Replace specific supplier names with XX印刷厂 in JSON templates.

This script finds all JSON templates where the supplier is one of:
- 菲迪印刷
- 凯源印刷  
- 深圳市众拓印刷

And replaces them with "XX印刷厂" in both the "供货商：" cell and footer supplier field.
"""

import json
from pathlib import Path

def replace_specific_suppliers():
    """Replace specific supplier names with XX印刷厂."""
    template_dir = Path(__file__).parent / "json_template"
    
    if not template_dir.exists():
        print(f"Template directory not found: {template_dir}")
        return
    
    # Suppliers to replace
    suppliers_to_replace = ["菲迪印刷", "凯源印刷", "深圳市众拓印刷", "宁波菲迪印刷有限公司", "宁波凯源印刷有限公司"]
    replacement_supplier = "XX印刷厂"
    
    json_files = list(template_dir.glob("*.json"))
    print(f"Found {len(json_files)} JSON files to process...")
    print(f"Looking for suppliers: {suppliers_to_replace}")
    print(f"Will replace with: {replacement_supplier}")
    print()
    
    updated_count = 0
    
    for json_file in json_files:
        try:
            # Read the JSON file
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            modified = False
            
            # Check and update "供货商：" in cells section
            if "cells" in data:
                for cell_key, cell_data in data["cells"].items():
                    if isinstance(cell_data, dict) and "key" in cell_data:
                        if cell_data["key"] == "供货商：":
                            current_value = cell_data.get("value", "")
                            if current_value in suppliers_to_replace:
                                cell_data["value"] = replacement_supplier
                                print(f"  {json_file.name}: Updated '供货商：' from '{current_value}' to '{replacement_supplier}'")
                                modified = True
            
            # Check and update "supplier" in footer section
            if "footer" in data and "supplier" in data["footer"]:
                current_supplier = data["footer"]["supplier"]
                if current_supplier in suppliers_to_replace:
                    data["footer"]["supplier"] = replacement_supplier
                    print(f"  {json_file.name}: Updated footer supplier from '{current_supplier}' to '{replacement_supplier}'")
                    modified = True
            
            # Write back the modified JSON if changes were made
            if modified:
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                updated_count += 1
            
        except Exception as e:
            print(f"Error processing {json_file.name}: {e}")
    
    print(f"\nUpdated {updated_count} JSON files successfully.")

def main():
    """Main function to replace specific suppliers."""
    print("Starting supplier replacement process...")
    replace_specific_suppliers()
    print("\nSupplier replacement process completed.")

if __name__ == "__main__":
    main()
