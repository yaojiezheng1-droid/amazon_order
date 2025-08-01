import sys
import json
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image
except ModuleNotFoundError as exc:  # pragma: no cover - dependency missing in tests
    raise SystemExit("openpyxl and pillow are required to run this script") from exc

PRODUCT_START_ROW = 7
COLUMN_MAP = {
    '产品编号': 'A',
    '产品图片': 'B',
    '描述': 'C',
    '数量/个': 'D',
    '单价': 'E',
    '包装方式': 'G',
}


def fill_workbook(template: Path, data: dict):
    """Return workbook filled with ``data`` using ``template``."""
    wb = load_workbook(template)
    ws = wb.active

    for addr, info in data.get('cells', {}).items():
        ws[addr] = info.get('value', '')

    row = PRODUCT_START_ROW
    for product in data.get('products', []):
        for key, col in COLUMN_MAP.items():
            if key in product:
                if key == '产品图片':
                    img_path = Path(product[key])
                    if not img_path.is_absolute() and not img_path.exists():
                        alt_path = template.parent.parent / img_path
                        if alt_path.exists():
                            img_path = alt_path
                    try:
                        img = Image(img_path)
                    except Exception:
                        ws[f"{col}{row}"] = product[key]
                    else:
                        ws.add_image(img, f"{col}{row}")
                else:
                    ws[f"{col}{row}"] = product[key]
        qty = product.get('数量/个')
        price = product.get('单价')
        if qty not in (None, '') and price not in (None, ''):
            ws[f"F{row}"] = f"=D{row}*E{row}"
        row += 1

    footer = data.get('footer', {})
    if 'buyer' in footer:
        ws['B69'] = footer['buyer']
    if 'supplier' in footer:
        ws['E69'] = footer['supplier']
    return wb


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print("usage: json_PO_excel.py <input.json> <output.xlsx>")
        return 1

    json_path = Path(argv[1])
    out_path = Path(argv[2])
    template = Path(__file__).resolve().parent / 'docs' / 'empty_base_template.xlsx'

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    wb = fill_workbook(template, data)
    wb.save(out_path)
    return 0


if __name__ == '__main__':
    raise SystemExit(main(sys.argv))
