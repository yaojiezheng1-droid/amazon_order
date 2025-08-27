#!/usr/bin/env python3
"""Clear product quantity values in JSON templates.

This script clears the "数量/个" field values in all JSON template files,
setting them to 0 or empty.
"""

import json
from pathlib import Path

def clear_product_quantities():
    """Clear all product quantity values in JSON templates."""
    template_dir = Path(__file__).parent / "json_template"
    
    if not template_dir.exists():
        print(f"Template directory not found: {template_dir}")
        return
    
    json_files = list(template_dir.glob("*.json"))
    print(f"Found {len(json_files)} JSON files to process...")
    print("Clearing product '数量/个' values...")
    print()
    
    updated_count = 0
    
    for json_file in json_files:
        try:
            # Read the JSON file
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            modified = False
            
            # Check products section
            if "products" in data:
                for product in data["products"]:
                    if "数量/个" in product:
                        old_quantity = product["数量/个"]
                        product["数量/个"] = 0  # Set to 0
                        print(f"  {json_file.name}: Cleared '数量/个' from {old_quantity} to 0")
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
    """Main function to clear product quantities."""
    print("Starting product quantity clearing process...")
    clear_product_quantities()
    print("\nProduct quantity clearing process completed.")

if __name__ == "__main__":
    main()
