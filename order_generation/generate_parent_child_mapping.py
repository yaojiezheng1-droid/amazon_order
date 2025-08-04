import json
import sys
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, Any

NAMESPACE = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def read_rows(xlsx: str):
    """Return list of rows from the first worksheet of ``xlsx``.

    Only standard library modules are used to avoid external dependencies."""
    with zipfile.ZipFile(xlsx) as z:
        shared: list[str] = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('.//' + NAMESPACE + 'si'):
                text = ''.join(t.text or '' for t in si.findall('.//' + NAMESPACE + 't'))
                shared.append(text)
        sheet_root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        rows_maps = []
        max_col = 0
        for row in sheet_root.findall('.//' + NAMESPACE + 'row'):
            row_map: Dict[int, str] = {}
            for c in row.findall(NAMESPACE + 'c'):
                addr = c.get('r')
                col_letters = ''.join(ch for ch in addr if ch.isalpha())
                col_idx = 0
                for ch in col_letters:
                    col_idx = col_idx * 26 + ord(ch) - 64
                col_idx -= 1
                t = c.get('t')
                if t == 's':
                    v = c.find(NAMESPACE + 'v')
                    val = shared[int(v.text)] if v is not None else ''
                elif t == 'inlineStr':
                    is_elem = c.find(NAMESPACE + 'is')
                    val = ''.join(
                        tn.text or '' for tn in is_elem.findall('.//' + NAMESPACE + 't')
                    ) if is_elem is not None else ''
                else:
                    v = c.find(NAMESPACE + 'v')
                    val = v.text if v is not None else ''
                row_map[col_idx] = val
                if col_idx > max_col:
                    max_col = col_idx
            rows_maps.append(row_map)
        rows = []
        for row_map in rows_maps:
            rows.append([row_map.get(i, '') for i in range(max_col + 1)])
        return rows


def build_mapping(xlsx: str) -> Dict[str, Any]:
    rows = read_rows(xlsx)
    if not rows:
        return {"parents": {}}
    header = rows[0]
    sku_idx = header.index('SKU')
    parent_asin_idx = header.index('父ASIN')
    name_idx = header.index('品名') if '品名' in header else None
    groups: Dict[str, Dict[str, Any]] = {}
    for row in rows[1:]:
        if len(row) <= max(sku_idx, parent_asin_idx):
            continue
        sku = row[sku_idx].strip()
        parent_asin = row[parent_asin_idx].strip()
        name = row[name_idx].strip() if name_idx is not None and len(row) > name_idx else ''
        if not sku or not parent_asin:
            continue
        group = groups.setdefault(parent_asin, {"name": name, "skus": []})
        if not group["name"] and name:
            group["name"] = name
        if sku not in group["skus"]:
            group["skus"].append(sku)
    parents: Dict[str, Dict[str, Any]] = {}
    for info in groups.values():
        skus = info["skus"]
        if len(skus) > 1:
            parent_sku = skus[0]
            parents[parent_sku] = {"name": info["name"], "children": skus}
    return {"parents": parents}


def main() -> int:
    if len(sys.argv) != 3:
        print('usage: generate_parent_child_mapping.py <input.xlsx> <output.json>')
        return 1
    mapping = build_mapping(sys.argv[1])
    with open(sys.argv[2], 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
    return 0


if __name__ == '__main__':
    sys.exit(main())
