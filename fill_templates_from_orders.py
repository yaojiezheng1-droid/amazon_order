import json
from pathlib import Path
from typing import Dict, Set

import os

# advanced_excel_to_json expects to run with the working directory set to
# ``order_generation`` so it can load mapping files from ``docs``.
def parse_order(path: str) -> dict:
    cwd = os.getcwd()
    os.chdir('order_generation')
    import sys
    sys.path.insert(0, '')
    from advanced_excel_to_json import parse_order as adv_parse
    rel = Path(path)
    base = Path('order_generation').resolve()
    if not str(rel).startswith(str(base)):
        rel = rel.resolve().relative_to(base)
    else:
        rel = rel.relative_to('order_generation')
    data = adv_parse(str(rel))
    os.chdir(cwd)
    if isinstance(data, list):
        return data[0] if data else {"cells": {}, "products": [], "footer": {}}
    return data

MAPPING_PATH = Path('order_generation/docs/complete_mapping_with_po.json')
ORDERS_DIR = Path('order_generation/orders')
TEMPLATE_DIR = Path('order_generation/json_template')


def build_po_map() -> Dict[str, Set[str]]:
    """Return mapping of purchase_order_file -> set of SKUs."""
    with MAPPING_PATH.open('r', encoding='utf-8') as f:
        data = json.load(f)

    mapping: Dict[str, Set[str]] = {}

    def walk(obj):
        if isinstance(obj, dict):
            po_file = obj.get('purchase_order_file')
            sku = obj.get('sku')
            if po_file and sku:
                mapping.setdefault(po_file, set()).add(sku)
            for v in obj.values():
                walk(v)
        elif isinstance(obj, list):
            for item in obj:
                walk(item)

    walk(data)
    return mapping


def merge_template(template: dict, data: dict) -> dict:
    cells = template.get('cells', {})
    key_map = {meta.get('key', '').strip(): addr for addr, meta in cells.items() if isinstance(meta, dict)}
    for _addr, meta in data.get('cells', {}).items():
        key = str(meta.get('key', '')).strip()
        if key and key in key_map:
            cells[key_map[key]]['value'] = meta.get('value')
    template['cells'] = cells
    template['products'] = data.get('products', [])
    template['footer'] = data.get('footer', {})
    return template


def process_order(excel_path: Path, skus: Set[str]):
    parsed = parse_order(str(excel_path))
    for sku in skus:
        template_path = TEMPLATE_DIR / f'{sku}.json'
        if not template_path.exists():
            print(f'Skipping {excel_path.name}: template {template_path.name} not found')
            continue
        with template_path.open('r', encoding='utf-8') as f:
            template = json.load(f)
        filled = merge_template(template, parsed)
        with template_path.open('w', encoding='utf-8') as f:
            json.dump(filled, f, ensure_ascii=False, indent=2)
        print(f'Filled template {template_path.name} from {excel_path.name}')


def main():
    po_map = build_po_map()
    for excel_file in ORDERS_DIR.glob('*.xlsx'):
        skus = po_map.get(excel_file.name)
        if not skus:
            continue
        process_order(excel_file, skus)


if __name__ == '__main__':
    main()
