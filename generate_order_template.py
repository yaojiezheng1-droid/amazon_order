import json
from typing import List, Dict, Any
import xlsxwriter


def write_description(ws, row: int, col: int, runs: List[Dict[str, Any]], workbook):
    """Write rich text description supporting bold and background color."""
    segments = []
    for run in runs:
        text = run.get('text', '')
        fmt_args = {}
        if run.get('bold'):
            fmt_args['bold'] = True
        if run.get('bgcolor'):
            fmt_args['bg_color'] = run['bgcolor']
        if fmt_args:
            fmt = workbook.add_format(fmt_args)
            segments.append(fmt)
        segments.append(text)
    ws.write_rich_string(row, col, *segments)


def create_order_workbook(data: Dict[str, Any], output: str) -> None:
    wb = xlsxwriter.Workbook(output)
    ws = wb.add_worksheet()

    row = 0
    normal = wb.add_format({'font_size': 11})
    bold = wb.add_format({'bold': True})

    # company header
    for line in data.get('company_header', []):
        ws.write(row, 0, line, normal)
        row += 1

    # document title
    title = data.get('title', '订单')
    ws.write(row, 0, title, bold)
    row += 1

    # supplier/order info lines
    info_lines = data.get('info_lines', [])
    for info in info_lines:
        ws.write_row(row, 0, info, normal)
        row += 1

    # table header
    columns = data['table']['columns']
    ws.write_row(row, 0, columns, bold)
    row += 1

    # items
    for item in data['table']['items']:
        col_idx = 0
        for key in columns:
            value = item.get(key, '')
            if key == '描述' and isinstance(value, list):
                write_description(ws, row, col_idx, value, wb)
            else:
                ws.write(row, col_idx, value, normal)
            col_idx += 1
        row += 1

    # totals if present
    if 'total' in data['table']:
        ws.write_row(row, 0, data['table']['total'], normal)
        row += 1

    # notes
    for note in data.get('notes', []):
        ws.write(row, 0, note, normal)
        row += 1

    # footer (signatures)
    if 'footer' in data:
        ws.write_row(row, 0, [data['footer'].get('buyer', ''), '', data['footer'].get('supplier', '')], normal)
    wb.close()


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Generate order excel from JSON data')
    parser.add_argument('json', help='Path to order data json')
    parser.add_argument('output', help='Output Excel file path')
    args = parser.parse_args()
    with open(args.json, 'r', encoding='utf-8') as f:
        order_data = json.load(f)
    create_order_workbook(order_data, args.output)
