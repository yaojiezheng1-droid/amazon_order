# Amazon Order Workflow

This repository organizes files used for managing Amazon orders. It is designed for manual updating of spreadsheets with help from AI tools. The main directories are:

- `order_generation/docs/` – Project notes, templates, and mapping JSON files.
- `order_generation/PO_excel/` – Past orders and new purchase orders.
- `order_generation/accessories/` – Packaging and accessory information.
- `order_generation/json_exports/` – Generated JSON exports from Excel files.
- `sales/` – Sales history and reports.

Each directory contains an empty `.gitkeep` file so that it is tracked by Git even when no spreadsheets are present.

## JSON and Excel Structure

The tools in `order_generation` share a common data format:

- **JSON templates** contain three sections:
  - `cells`: mapping of Excel cell addresses to key/value pairs.
  - `products`: list of items with fields such as `产品编号` (SKU), `产品图片`, `描述`,
    `数量/个`, `单价`, and `包装方式`.
  - `footer`: buyer and supplier information.
- **Excel files** follow `docs/empty_base_template.xlsx` where:
  - Metadata lives in cells like `B3` (供货商), `G4` (日期), etc.
  - The product table starts around row 7 with columns `A`=SKU, `B`=图片,
    `C`=描述, `D`=数量, `E`=单价, and `G`=包装方式.

## Generating JSON Templates

Use `generate_JSON_template.py` with a JSON file that follows the

structure documented in `order_generation/docs/order_template.md`. The JSON lists every
yellow cell from the Excel template with its label (`key`) and the value
to write. The script fills those cells in `order_generation/docs/empty_base_template.xlsx`
and inserts product rows. Green cells contain formulas that compute totals
automatically.
An example JSON file is provided in `order_generation/docs/order_template_example.json`.

When generating the spreadsheet the script inserts any image referenced in the
`产品图片` field of each product row. The value should be a path to the image
file on disk.  The product name (`产品名称`) is automatically prepended to the
description when filling the "描述" column.


## Order Rules

- Each new order may contain only one parent product as defined in
`order_generation/docs/parent_child_mapping.json`. When ordering child SKUs belonging to
  different parents, create a separate spreadsheet for each parent product.


## Converting PO_files to JSON_template

Manually Fill valus in JSON

Locate the purchase‑order (PO) file(s) whose SKU matches each JSON filename (look in order_generation/PO_excel/).

Using the PO, manually fill every field in the JSON template. Do not add, remove, or rename keys; keep the structure exactly the same.

When copying the product‑description text, preserve any bold text, highlights, and font‑color formatting found in the PO.

If you cannot find a matching PO, fill the template with the most reasonable information you can infer and continue; never abort the process.

## Handling Direct SKU Requests

Sometimes an order request only specifies a SKU and quantity, for example:

"I want to place order of 800 of 48-82P3-QSFG product."


> You are generating purchase orders from SKU/quantity requests. Work through
> the numbered checklist one item at a time and only move to the next item once
> the current one is fully complete. Your goal is to populate
> `order_generation/docs/empty_base_template.xlsx` with the correct values by json_PO_excel.py from
> the JSON templates and produce a finished spreadsheet.

1. Look up the SKU in `order_generation/docs/complete_mapping.json` to
   determine all required accessories and their ratios.
2. For the main product and each accessory, copy the corresponding JSON
   template from `order_generation/json_template/` and set the `数量/个` field to
   the requested quantity.
3. download images in the json_templates by its path together with `empty_base_template.xlsx`, `json_PO_excel.py` in structure of following
      order_generation/
      ├── json_PO_excel.py           # ← Should stay here
      ├── merge_json_templates.py    # Other processing scripts
      ├── other_scripts.py
      ├── docs/
      │   └── empty_base_template.xlsx
      ├── images/
      │   ├── accessories/
      │   ├── colors/
      │   ├── logos/
      │   ├── products/
      └── json_template/
         ├── template1.json
         └── template2.json
4. Decide if any items originate from the same factory:
   - **Different factories** – keep the templates separate and create one Excel
     file per factory.
   - **Same factory** – use `order_generation/merge_json_templates.py` to merge
     the JSON data from the same factory. The script appends every entry to the `products` list and
     chooses the most appropriate value for each cell in the merged `cells`
     section.

### Automating steps 1, 2, and 4

Instead of performing the previous steps manually, run:

```bash
python order_generation/direct_sku_to_json.py <asin1> <qty1> [<asin2> <qty2> ...]
```

The script reads accessory ratios from `order_generation/docs/complete_mapping.json`,
fills in each template's `数量/个` field, groups items by supplier, and writes a
merged JSON file for each factory to `order_generation/json_exports/`. Continue
with step 5 using these generated files.
5. For each factory group, run `json_PO_excel.py` to write the JSON values into
   `order_generation/docs/empty_base_template.xlsx`.
   !excel must be created by json_PO_excel.py not manually!
   !!do not make your own Excel generating python strickly adhere to the readme, use json_PO_excel.py for excel generation!!
   !!do not make your own Excel generating python strickly adhere to the readme, use json_PO_excel.py for excel generation!!
   !!do not make your own Excel generating python strickly adhere to the readme, use json_PO_excel.py for excel generation!!
6. Confirm that the populated `empty_base_template.xlsx` matches the cell
   addresses expected by the JSON. 
7. If you find any excels having the same factory name, it is a mistake go back to step 3 and merge json_template with same factory

This process keeps products from different factories on separate spreadsheets
while still providing a single sheet when everything is sourced from one
factory.

## Additional Tools

### `product_search_gui.py`
- Visual interface to search templates by name or SKU
- Builds commands for `direct_sku_to_json.py`
- Optional warehouse selection and PO-import generation

### `excel_to_json_template.py`
- Converts `empty_base_template.xlsx`-style files to JSON templates
- GUI or command-line usage
- Example: `python order_generation/excel_to_json_template.py order.xlsx`

### `json_templates_to_excel.py`
- Batch converts JSON templates back into Excel
- Leverages `json_PO_excel.py` for consistent formatting
- Example: `python order_generation/json_templates_to_excel.py --templates ST1122-1 EEHB-NBB`

### `fill_po_import.py`
- Fills `docs/PO_import_empty.xlsx` using JSON exports
- Run directly or through `direct_sku_to_json.py --po-import`
- Example: `python order_generation/fill_po_import.py AM2025 --warehouse "义乌仓库"`

## Typical Workflow

1. Search and select products with `product_search_gui.py` or call
   `direct_sku_to_json.py` with SKU/quantity pairs to create JSON exports and
   factory-grouped Excel orders.
2. When a supplier needs the PO-import format, run `fill_po_import.py` or add
   `--po-import` and `--warehouse` when generating orders.
3. To create new templates from an existing spreadsheet, convert it with
   `excel_to_json_template.py`.
4. To rebuild Excel files from templates, use `json_templates_to_excel.py`.

These scripts remain separate because each handles a distinct stage of the
workflow while sharing the common JSON and Excel structures described above.
