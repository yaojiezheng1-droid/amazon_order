import sys
from pathlib import Path

# Add order_generation path to import excel_to_json
sys.path.append('order_generation')

from add_purchase_orders import load_purchase_orders, extend_mapping

DEFAULT_PO = Path('order_generation/docs/采购单_modified.xlsx')
DEFAULT_MAPPING = Path('order_generation/docs/complete_mapping_with_po.json')


def main(purchase_path: Path = DEFAULT_PO, mapping_path: Path = DEFAULT_MAPPING, output_path: Path = None):
    if output_path is None:
        output_path = mapping_path
    po_map = load_purchase_orders(purchase_path)
    extend_mapping(mapping_path, po_map, output_path)
    print(f'Updated {output_path}')


if __name__ == '__main__':
    purchase = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_PO
    mapping = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_MAPPING
    output = Path(sys.argv[3]) if len(sys.argv) > 3 else mapping
    main(purchase, mapping, output)
