# coding: utf-8
"""Convert order spreadsheets to JSON using heuristics.

This script scans xlsx files in the repository and outputs JSON files in
`json_exports` similar to `docs/order_template_example.json`.
It does not rely on cell fill colors. Instead it detects the product table
by looking for header keywords like "数量" and "单价".
If products from multiple parent SKUs are detected, the order is split into
multiple JSON files keeping the same general info.
"""

import json
import re
from pathlib import Path
from typing import Any, Dict, List

from excel_to_json import read_workbook, guess_key

# load accessory and parent-child mappings
with open("docs/accessory_mapping.json", "r", encoding="utf-8") as f:
    ACCESSORY_MAP = json.load(f)["products"]
with open("docs/parent_child_mapping.json", "r", encoding="utf-8") as f:
    PARENT_DATA = json.load(f)["parents"]
CHILD_TO_PARENT = {
    child: parent for parent, info in PARENT_DATA.items() for child in info["children"]
}

HEADER_MAP = {
    "产品编号": "产品编号",
    "型号": "产品编号",
    "客号": "产品编号",
    "产品图片": "产品图片",
    "图片": "产品图片",
    "商品名称": "描述",
    "名称": "描述",
    "规格": "描述",
    "描述": "描述",
    "数量": "数量/个",
    "数量/个": "数量/个",
    "单价": "单价",
    "单价/个": "单价",
    "金额": None,
    "包装方式": "包装方式",
    "备注": "备注",
}

COLUMNS = [chr(65 + i) for i in range(12)]  # A-L


def slugify(name: str) -> str:
    base = Path(name).stem
    return re.sub(r"[^A-Za-z0-9_-]+", "_", base)


def convert_number(val: str) -> Any:
    try:
        if val.strip() == "":
            return ""
        if "." in val:
            return float(val)
        return int(val)
    except Exception:
        return val


def detect_header(cells: Dict[str, Any]) -> int:
    max_row = max(int(re.findall(r"\d+", addr)[0]) for addr in cells)
    for r in range(1, max_row + 1):
        row_vals = [str(cells.get(f"{col}{r}", ("", None, None))[0]).strip() for col in COLUMNS]
        qty_present = any(v in ("数量", "数量/个") for v in row_vals)
        price_present = any(v in ("单价", "单价/个") for v in row_vals)
        if qty_present and price_present:
            return r
    return -1


def parse_table(cells: Dict[str, Any], header_row: int) -> (List[Dict[str, Any]], int):
    header = {}
    for col in COLUMNS:
        text = str(cells.get(f"{col}{header_row}", ("", None, None))[0]).strip()
        if text in HEADER_MAP:
            header[col] = HEADER_MAP[text]
    products: List[Dict[str, Any]] = []
    row = header_row + 1
    while True:
        row_vals = {col: str(cells.get(f"{col}{row}", ("", None, None))[0]).strip() for col in COLUMNS}
        if all(v == "" for v in row_vals.values()):
            break
        first = row_vals.get("A")
        if first.startswith("TOTAL") or first == "总计":
            break
        if all(v == "" for v in [row_vals.get(c, "") for c in header]):
            break
        item: Dict[str, Any] = {}
        for col, key in header.items():
            if not key:
                continue
            val = row_vals.get(col, "")
            item[key] = convert_number(val)
        sku = item.get("产品编号")
        if sku and sku in ACCESSORY_MAP:
            item["产品名称"] = ACCESSORY_MAP[sku]["name"]
        products.append(item)
        row += 1
    return products, row - 1


def collect_cells(cells: Dict[str, Any], start: int, end: int) -> Dict[str, Any]:
    used_rows = set(range(start, end + 1))
    out: Dict[str, Any] = {}
    for addr, (val, _color, _formula) in cells.items():
        r = int(re.findall(r"\d+", addr)[0])
        if r >= start and r <= end:
            continue
        if str(val).strip() == "":
            continue
        out[addr] = {"key": guess_key(addr, cells), "value": val}
    return out


def group_by_parent(products: List[Dict[str, Any]]):
    groups: Dict[str, List[Dict[str, Any]]] = {}
    for p in products:
        sku = p.get("产品编号")
        parent = CHILD_TO_PARENT.get(sku, sku)
        groups.setdefault(parent, []).append(p)
    return groups


def parse_order(path: str) -> List[Dict[str, Any]]:
    cells = read_workbook(path)
    header_row = detect_header(cells)
    if header_row == -1:
        return [{"cells": collect_cells(cells, 0, 0), "products": [], "footer": {}}]
    products, end_row = parse_table(cells, header_row)
    info_cells = collect_cells(cells, header_row, end_row)
    groups = group_by_parent(products)
    outputs = []
    for parent, items in groups.items():
        data = {
            "cells": info_cells,
            "products": items,
            "footer": {},
        }
        outputs.append(data)
    return outputs


def main():
    out_dir = Path("json_exports")
    out_dir.mkdir(exist_ok=True)
    for path in Path(".").glob("*.xlsx"):
        for idx, data in enumerate(parse_order(str(path))):
            slug = slugify(path.name)
            if idx:
                slug += f"_{idx+1}"
            out_path = out_dir / f"{slug}.json"
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"converted {path} -> {out_path}")


if __name__ == "__main__":
    main()
