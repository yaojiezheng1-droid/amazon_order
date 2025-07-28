import sys
import json
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, Tuple, Any, List

NAMESPACE = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
YELLOW = 'FFFFFF00'


def guess_key(address: str, cells: Dict[str, Tuple[str, str, str]]) -> str:
    """Return a label for the given cell address.

    The label is taken from the nearest non-yellow cell either to the left
    on the same row or above in the same column.  If no such cell exists the
    returned key is an empty string.
    """
    # separate column letters and row number
    col = ''.join(ch for ch in address if ch.isalpha())
    row = int(''.join(ch for ch in address if ch.isdigit()))

    # look left
    if len(col) == 1 and col > 'A':
        left_col = chr(ord(col) - 1)
        left_addr = f'{left_col}{row}'
        left = cells.get(left_addr)
        if left and left[1] != YELLOW and left[0]:
            return str(left[0])

    # look upward
    r = row - 1
    while r > 0:
        up_addr = f'{col}{r}'
        cell = cells.get(up_addr)
        if cell:
            if cell[1] != YELLOW and cell[0]:
                return str(cell[0])
        r -= 1
    return ''


def read_workbook(path: str) -> Dict[str, Tuple[str, str, str]]:
    """Return mapping of cell address -> (value, color, formula)."""
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('.//' + NAMESPACE + 'si'):
                text = ''.join(t.text or '' for t in si.findall('.//' + NAMESPACE + 't'))
                shared.append(text)

        style_root = ET.fromstring(z.read('xl/styles.xml'))
        fills = style_root.find(NAMESPACE + 'fills')
        fill_colors = []
        for f in fills.findall(NAMESPACE + 'fill'):
            pattern = f.find(NAMESPACE + 'patternFill')
            color = None
            if pattern is not None:
                fg = pattern.find(NAMESPACE + 'fgColor')
                if fg is not None:
                    color = fg.get('rgb')
            fill_colors.append(color)
        cell_xfs = style_root.find(NAMESPACE + 'cellXfs')
        xfs = [int(xf.get('fillId')) for xf in cell_xfs.findall(NAMESPACE + 'xf')]

        sheet_root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        cells = {}
        for c in sheet_root.findall('.//' + NAMESPACE + 'c'):
            addr = c.get('r')
            t = c.get('t')
            v = c.find(NAMESPACE + 'v')
            val = ''
            if v is not None:
                val = shared[int(v.text)] if t == 's' else v.text
            f = c.find(NAMESPACE + 'f')
            formula = f.text if f is not None else None
            color = None
            s = c.get('s')
            if s is not None:
                color = fill_colors[xfs[int(s)]]
            cells[addr] = (val, color, formula)
        return cells


def parse_order(path: str) -> Dict[str, Any]:
    cells = read_workbook(path)

    # Determine product rows based on yellow fill in columns A-G starting at row 7
    products = []
    row = 7
    while True:
        row_cells = [cells.get(f'{chr(65+i)}{row}', ('', None, None)) for i in range(7)]
        if not all(c[1] == YELLOW for c in row_cells):
            break
        products.append({
            '产品编号': row_cells[0][0],
            '产品图片': row_cells[1][0],
            '描述': row_cells[2][0],
            '数量/个': row_cells[3][0],
            '单价': row_cells[4][0],
            '包装方式': row_cells[6][0],
        })
        row += 1

    yellow_cells = {}
    for addr, (val, color, _) in cells.items():
        if color == YELLOW:
            r = int(''.join(ch for ch in addr if ch.isdigit()))
            if r < 7 or r >= row:
                yellow_cells[addr] = {
                    'key': guess_key(addr, cells),
                    'value': val
                }

    footer = {}
    if 'B69' in cells:
        footer['buyer'] = cells['B69'][0]
    if 'E69' in cells:
        footer['supplier'] = cells['E69'][0]

    data = {
        'cells': yellow_cells,
        'products': products,
        'footer': footer,
    }
    return data


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('usage: excel_to_json.py <input.xlsx> <output.json>')
        sys.exit(1)
    data = parse_order(sys.argv[1])
    with open(sys.argv[2], 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
