import json
import re
from datetime import datetime
from pathlib import Path
import sys

# Add order_generation path for excel_to_json
sys.path.append('order_generation')
from excel_to_json import read_workbook

PURCHASE_PATH = Path('order_generation/docs/采购单.xlsx')
MAPPING_PATH = Path('order_generation/docs/complete_mapping.json')
OUTPUT_PATH = Path('order_generation/docs/complete_mapping_with_po.json')


def parse_time(t: str) -> datetime:
    try:
        return datetime.strptime(t.strip(), '%Y-%m-%d %H:%M:%S')
    except Exception:
        return datetime.min


def load_purchase_orders(path: Path):
    cells = read_workbook(str(path))
    rows = {}
    for addr, (val, _color, _formula) in cells.items():
        m = re.match(r'([A-Z]+)(\d+)', addr)
        if not m:
            continue
        col = m.group(1)
        row = int(m.group(2))
        rows.setdefault(row, {})[col] = val

    current_po = None
    latest = {}
    for r in sorted(rows.keys()):
        if r == 1:
            continue
        row = rows[r]
        if row.get('A'):
            current_po = row['A']
        sku = row.get('AO')
        if not sku:
            continue
        time_str = row.get('L', '')
        t = parse_time(time_str)
        info = latest.get(sku)
        if info is None or t > info['time']:
            latest[sku] = {'po': current_po, 'time': t}
    return {sku: info['po'] for sku, info in latest.items()}


def extend_mapping(mapping_path: Path, po_map: dict, output_path: Path):
    with mapping_path.open('r', encoding='utf-8') as f:
        data = json.load(f)

    for parent in data.get('parents', {}).values():
        for child in parent.get('children', []):
            child_po = po_map.get(child.get('sku'))
            if child_po:
                child['purchase_order'] = child_po
            for acc in child.get('accessories', []):
                acc_po = po_map.get(acc.get('sku'))
                if acc_po:
                    acc['purchase_order'] = acc_po

    with output_path.open('w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def main():
    po_map = load_purchase_orders(PURCHASE_PATH)
    extend_mapping(MAPPING_PATH, po_map, OUTPUT_PATH)
    print(f'Wrote {OUTPUT_PATH}')


if __name__ == '__main__':
    main()
