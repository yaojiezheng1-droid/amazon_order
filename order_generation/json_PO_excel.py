import json
from pathlib import Path
from openpyxl import load_workbook
from generate_order_template import fill_cells, insert_products, BUYER_ROW

DEFAULT_TEMPLATE = Path('docs/empty_base_template.xlsx')


def json_to_excel(json_path: Path, template_path: Path = DEFAULT_TEMPLATE, output_path: Path = None) -> None:
    if output_path is None:
        output_path = template_path
    with Path(json_path).open('r', encoding='utf-8') as f:
        data = json.load(f)
    wb = load_workbook(template_path)
    ws = wb.active
    fill_cells(ws, data.get('cells', {}))
    offset = insert_products(ws, data.get('products', []))
    footer = data.get('footer', {})
    if footer:
        if 'buyer' in footer:
            ws[f'B{BUYER_ROW + offset}'] = footer['buyer']
        if 'supplier' in footer:
            ws[f'E{BUYER_ROW + offset}'] = footer['supplier']
    wb.save(output_path)
    print(f'Wrote {output_path}')


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Write JSON PO data into an Excel template')
    parser.add_argument('json', help='Path to order data JSON')
    parser.add_argument('--template', default=str(DEFAULT_TEMPLATE), help='Template workbook path')
    parser.add_argument('--output', default=None, help='Output Excel file path (defaults to template path)')
    args = parser.parse_args()

    tpl = Path(args.template)
    out = Path(args.output) if args.output else tpl
    json_to_excel(Path(args.json), tpl, out)
