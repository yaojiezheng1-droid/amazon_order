import json
import sys
from typing import Dict, Any


def merge(accessory_path: str, parent_path: str) -> Dict[str, Any]:
    with open(accessory_path, 'r', encoding='utf-8') as f:
        accessory_data = json.load(f).get("products", {})
    with open(parent_path, 'r', encoding='utf-8') as f:
        parent_data = json.load(f).get("parents", {})

    merged: Dict[str, Any] = {"parents": {}}
    for parent_sku, info in parent_data.items():
        parent_entry = {"name": info.get("name", ""), "children": []}
        for child_sku in info.get("children", []):
            child_info = accessory_data.get(child_sku, {})
            parent_entry["children"].append({
                "sku": child_sku,
                "name": child_info.get("name", ""),
                "accessories": child_info.get("accessories", [])
            })
        merged["parents"][parent_sku] = parent_entry
    return merged


def main() -> int:
    if len(sys.argv) != 4:
        print("usage: merge_mappings.py <accessory.json> <parent_child.json> <output.json>")
        return 1
    merged = merge(sys.argv[1], sys.argv[2])
    with open(sys.argv[3], 'w', encoding='utf-8') as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)
    return 0


if __name__ == '__main__':
    sys.exit(main())
