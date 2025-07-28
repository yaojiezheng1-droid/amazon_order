import sys
import zipfile
import json
import xml.etree.ElementTree as ET
from typing import List, Dict

NAMESPACE = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def read_rows(path: str) -> List[List[str]]:
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('.//' + NAMESPACE + 'si'):
                text = ''.join(t.text or '' for t in si.findall('.//' + NAMESPACE + 't'))
                shared.append(text)
        sheet = z.read('xl/worksheets/sheet1.xml')
        root = ET.fromstring(sheet)
        rows = []
        for row in root.findall('.//' + NAMESPACE + 'row'):
            cells = []
            for c in row.findall(NAMESPACE + 'c'):
                t = c.get('t')
                v = c.find(NAMESPACE + 'v')
                val = ''
                if v is not None:
                    val = shared[int(v.text)] if t == 's' else v.text
                elif t == 'inlineStr':
                    is_elem = c.find(NAMESPACE + 'is')
                    if is_elem is not None:
                        val = ''.join(tel.text or '' for tel in is_elem.findall('.//' + NAMESPACE + 't'))
                cells.append(val if val is not None else '')
            rows.append(cells)
        return rows

def parse_order(rows: List[List[str]]) -> Dict:
    # trim empty trailing rows
    while rows and not any(rows[-1]):
        rows.pop()

    title_idx = None
    for i, r in enumerate(rows):
        text = ''.join(r)
        if '订单' in text or '打样单' in text:
            title_idx = i
            break
    if title_idx is None:
        raise ValueError('title row not found')

    company_header = []
    for r in rows[:title_idx]:
        if r and r[0]:
            for line in r[0].split('\n'):
                line = line.strip()
                if line:
                    company_header.append(line)

    title = ''.join(rows[title_idx]).strip()

    info_lines = []
    header_row = None
    idx = title_idx + 1
    while idx < len(rows):
        r = rows[idx]
        if any(x in r for x in ('客号', '型号', '图片', '商品名称', '产品描述')):
            header_row = r
            break
        if any(cell for cell in r):
            info_lines.append(r)
        idx += 1
    if header_row is None:
        raise ValueError('table header not found')

    columns = [c for c in header_row if c]
    idx += 1
    items = []
    total = None
    while idx < len(rows):
        r = rows[idx]
        if not any(r):
            idx += 1
            continue
        join = ''.join(r)
        if join.startswith('总计') or join.startswith('TOTAL') or join.startswith('TOTAL:'):
            total = r
            idx += 1
            continue
        if any(word in r[0] for word in ('备注', '以上', '包装', '交货', '付款', '注意')):
            break
        item = {columns[i]: (r[i] if i < len(r) else '') for i in range(len(columns))}
        items.append(item)
        idx += 1

    notes = []
    footer = {}
    while idx < len(rows):
        r = rows[idx]
        if any('宁波品秀美容科技有限公司' in c for c in r):
            footer['buyer'] = '宁波品秀美容科技有限公司'
        if any('瑾秀制刷' in c or '菲迪印刷' in c or '和鑫制刷厂' in c for c in r):
            footer['supplier'] = next((c for c in r if c), '')
        note = next((c for c in r if c), '')
        if note:
            notes.append(note)
        idx += 1

    data = {
        'company_header': company_header,
        'title': title,
        'info_lines': info_lines,
        'table': {'columns': columns, 'items': items},
    }
    if total:
        data['table']['total'] = total
    if notes:
        data['notes'] = notes
    if footer:
        data['footer'] = footer
    return data

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('usage: excel_to_json.py <input.xlsx> <output.json>')
        sys.exit(1)
    rows = read_rows(sys.argv[1])
    data = parse_order(rows)
    with open(sys.argv[2], 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
