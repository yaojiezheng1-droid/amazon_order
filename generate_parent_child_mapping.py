import json
import sys
from typing import Dict, Any
from excel_to_json import read_rows


def build_mapping(xlsx: str) -> Dict[str, Any]:
    rows = read_rows(xlsx)
    if not rows:
        return {"parents": {}}
    header = rows[0]
    try:
        sku_idx = header.index('*SKU')
    except ValueError:
        raise ValueError('SKU column "*SKU" not found')
    name_idx = header.index('品名') if '品名' in header else None
    mapping: Dict[str, Dict[str, Any]] = {}
    for row in rows[1:]:
        if sku_idx >= len(row):
            continue
        sku = row[sku_idx].strip()
        if not sku:
            continue
        name = row[name_idx].strip() if name_idx is not None and name_idx < len(row) else ''
        if '-' in sku:
            parent = sku.rsplit('-', 1)[0]
        else:
            parent = sku
        entry = mapping.setdefault(parent, {"name": name, "children": []})
        if not entry["name"]:
            entry["name"] = name
        entry["children"].append(sku)
    mapping = {p: v for p, v in mapping.items() if len(v["children"]) > 1}
    return {"parents": mapping}


def main():
    if len(sys.argv) != 3:
        print('usage: generate_parent_child_mapping.py <input.xlsx> <output.json>')
        return 1
    mapping = build_mapping(sys.argv[1])
    with open(sys.argv[2], 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
    return 0


if __name__ == '__main__':
    sys.exit(main())
