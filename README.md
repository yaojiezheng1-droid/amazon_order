# Amazon Order Workflow

This repository contains scripts and templates for producing purchase orders in
Excel from JSON templates and SKU requests.

## Directory Structure

- `order_generation/docs/` – base Excel template, accessory mappings, and
  project notes
- `order_generation/PO_excel/` – past orders and newly generated spreadsheets
- `order_generation/accessories/` – packaging and accessory information
- `order_generation/json_exports/` – JSON files generated from SKU requests
- `sales/` – sales history and reports

Each directory is tracked with an empty `.gitkeep` file so it remains in the
repository even when no spreadsheets are present.


## Data Formats

The scripts share a common representation of order data:

- **JSON templates** have three sections:
  - `cells`: mapping of Excel addresses to values
  - `products`: list of items with fields such as `产品编号` (SKU),
    `产品图片`, `描述`, `数量/个`, `单价`, and `包装方式`
  - `footer`: buyer and supplier information
- **Excel files** follow `docs/empty_base_template.xlsx` where metadata lives in
  cells like `B3` (供货商) and `G4` (日期); the product table starts around row 7
  with columns `A`=SKU, `B`=图片, `C`=描述, `D`=数量, `E`=单价, and
  `G`=包装方式

## Writing JSON Templates to Excel

Use `json_PO_excel.py` with a JSON file that follows
`order_generation/docs/order_template.md`. The script fills
`order_generation/docs/empty_base_template.xlsx`, inserts product rows, and
places any image referenced in `产品图片`. The product name (`产品名称`) is
automatically prepended to the description.

An example template is provided in
`order_generation/docs/order_template_example_blank.json`.

## Order Rules

- Keep each order to a single parent product. If child SKUs belong to different
  parents or factories, create a separate spreadsheet for each group.

## Preparing JSON from Purchase Orders

1. Locate the purchase‑order (PO) files in `order_generation/PO_excel/`.
2. Fill the corresponding JSON templates without adding or removing keys.
3. Preserve bold text, highlights, and font colors when copying product
   descriptions from the PO.
4. If information is missing, make a reasonable assumption and continue.

## Handling Direct SKU Requests

1. Look up accessory requirements in
   `order_generation/docs/accessory_mapping.json`.
2. Copy each product's template from `order_generation/json_template/` and set
   the `数量/个` field to the requested quantity.
3. Ensure any `产品图片` paths point to files within
   `order_generation/images/`.
4. Group items by factory. When multiple templates use the same supplier, merge
   them with `order_generation/merge_json_templates.py`.
5. For each factory group, run `json_PO_excel.py` to apply the JSON data to the
   Excel template. Do **not** craft Excel files manually.
6. Verify that the populated spreadsheet matches the expected cell addresses.
7. If multiple outputs share the same factory name, merge their JSON templates
   and regenerate the spreadsheet.

### Automating steps 1–4

Instead of performing the above steps manually, run:

```bash
python order_generation/direct_sku_to_json.py <sku1> <qty1> [<sku2> <qty2> ...]
```

The script reads `docs/accessory_mapping.json`, sets each template's `数量/个`
field, groups items by supplier, merges templates, and writes the results to
`order_generation/json_exports/`. It can also produce Excel and PO‑import files.

## Additional Tools

### `product_search_gui.py`
- Visual interface to search templates by name or SKU
- Builds commands for `direct_sku_to_json.py`
- Optional warehouse selection and PO‑import generation

### `excel_to_json_template.py`
- Converts `empty_base_template.xlsx`‑style files to JSON templates
- GUI or command‑line usage
- Example: `python order_generation/excel_to_json_template.py order.xlsx`

### `json_templates_to_excel.py`
- Batch converts JSON templates back into Excel using `json_PO_excel.py`
- Example: `python order_generation/json_templates_to_excel.py --templates ST1122-1 EEHB-NBB`

### `fill_po_import.py`
- Fills `docs/PO_import_empty.xlsx` using JSON exports
- Run directly or through `direct_sku_to_json.py --po-import`
- Example: `python order_generation/fill_po_import.py AM2025 --warehouse "义乌仓库"`

## Typical Workflow

1. Search and select products with `product_search_gui.py` or call
   `direct_sku_to_json.py` with SKU/quantity pairs to create JSON exports and
   factory‑grouped Excel orders.
2. When a supplier needs the PO‑import format, run `fill_po_import.py` or add
   `--po-import` and `--warehouse` when generating orders.
3. Convert existing spreadsheets to templates with `excel_to_json_template.py`.
4. Rebuild Excel files from templates using `json_templates_to_excel.py`.

These scripts remain separate because each handles a distinct stage of the
workflow while sharing the common JSON and Excel formats described above.

