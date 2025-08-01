import json
from pathlib import Path
from typing import Dict, Any, List

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

TEMPLATE_DEFAULT = Path("order_generation/docs/empty_base_template.xlsx")
START_ROW = 7  # first product row in the template
PLACEHOLDER_ROWS = 3  # number of placeholder product rows in the template
BUYER_ROW = 69  # base row for buyer name (before any inserted rows)


def verify_cells(ws, cells: Dict[str, Any]) -> None:
    """Ensure the ``key`` for each cell matches the template text.

    This validates that the JSON structure aligns with the workbook so the
    script fails fast when the template changes.
    """
    mismatches = []
    for address, meta in cells.items():
        expected = meta.get("key") if isinstance(meta, dict) else None
        if expected is None:
            continue
        template_val = ws[address].value
        template_text = str(template_val).strip() if template_val is not None else ""
        if template_text != expected:
            mismatches.append(f"{address}: template='{template_text}' json='{expected}'")
    if mismatches:
        msg = "Cell key mismatch between JSON and template:\n" + "\n".join(mismatches)
        raise ValueError(msg)


def fill_cells(ws, cells: Dict[str, Any]) -> None:
    """Write values from the json ``cells`` mapping into the worksheet."""
    for address, meta in cells.items():
        if isinstance(meta, dict):
            value = meta.get("value", "")
        else:
            value = meta
        ws[address] = value


def insert_products(ws, products: List[Dict[str, Any]]) -> int:
    """Insert product rows into the worksheet starting at ``START_ROW``.

    Returns the number of rows inserted which is used to adjust footer
    positions below the product table.
    """
    extra = max(0, len(products) - PLACEHOLDER_ROWS)
    if extra:
        ws.insert_rows(START_ROW + PLACEHOLDER_ROWS, extra)

    row_count = max(len(products), PLACEHOLDER_ROWS)
    for idx in range(row_count):
        r = START_ROW + idx
        item = products[idx] if idx < len(products) else {}
        ws.cell(row=r, column=1, value=item.get("产品编号", ""))
        img_path = item.get("产品图片")
        if img_path:
            try:
                ws.add_image(XLImage(img_path), f"B{r}")
            except Exception:
                ws.cell(row=r, column=2, value=img_path)
        else:
            ws.cell(row=r, column=2, value="")
        name = item.get("产品名称", "").strip()
        desc = item.get("描述", "").strip()
        if name:
            desc = f"{name} {desc}" if desc else name
        ws.cell(row=r, column=3, value=desc)
        ws.cell(row=r, column=4, value=item.get("数量/个", ""))
        ws.cell(row=r, column=5, value=item.get("单价", ""))
        ws.cell(row=r, column=6, value=f"=E{r}*D{r}")
        ws.cell(row=r, column=7, value=item.get("包装方式", ""))

    total_row = START_ROW + row_count
    ws.cell(row=total_row, column=4, value=f"=SUM(D{START_ROW}:D{total_row-1})")
    ws.cell(row=total_row, column=6, value=f"=SUM(F{START_ROW}:F{total_row-1})")
    return extra


def create_order_workbook(data: Dict[str, Any], template: Path, output: Path) -> None:
    wb = load_workbook(template)
    ws = wb.active

    cells = data.get("cells", {})
    verify_cells(ws, cells)
    fill_cells(ws, cells)
    offset = insert_products(ws, data.get("products", []))

    footer = data.get("footer", {})
    if footer:
        buyer_cell = f"B{BUYER_ROW + offset}"
        supplier_cell = f"E{BUYER_ROW + offset}"
        if "buyer" in footer:
            ws[buyer_cell] = footer["buyer"]
        if "supplier" in footer:
            ws[supplier_cell] = footer["supplier"]

    wb.save(output)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Fill order template spreadsheet from JSON")
    parser.add_argument("json", help="Path to order data json")
    parser.add_argument("output", help="Output Excel file path")
    parser.add_argument("--template", default=str(TEMPLATE_DEFAULT), help="Template workbook path")
    args = parser.parse_args()

    with open(args.json, "r", encoding="utf-8") as f:
        data = json.load(f)

    create_order_workbook(data, Path(args.template), Path(args.output))