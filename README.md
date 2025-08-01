# Amazon Order Workflow

This repository organizes files used for managing Amazon orders. It is designed for manual updating of spreadsheets with help from AI tools. The main directories are:

- `order_generation/docs/` – Project notes, templates, and mapping JSON files.
- `order_generation/orders/` – Past orders and new purchase orders.
- `order_generation/accessories/` – Packaging and accessory information.
- `order_generation/json_exports/` – Generated JSON exports from Excel files.
- `sales/` – Sales history and reports.

Each directory contains an empty `.gitkeep` file so that it is tracked by Git even when no spreadsheets are present.


## Generating Order Templates

Use `PO_excel.py` with a JSON file that follows the structure documented in
`order_generation/docs/order_template.md`. The JSON lists every yellow cell from
the Excel template with its label (`key`) and the value to write. The script
verifies that each cell reference matches the labels in
`order_generation/docs/empty_base_template.xlsx`, fills those cells, and inserts
product rows. Green cells contain formulas that compute totals automatically.
An example JSON file is provided in `order_generation/docs/order_template_example.json`.

When generating the spreadsheet the script inserts any image referenced in the
`产品图片` field of each product row. The value should be a path to the image file
on disk. The product name (`产品名称`) is automatically prepended to the
description when filling the "描述" column.


## Order Rules

- Each new order may contain only one parent product as defined in
`order_generation/docs/parent_child_mapping.json`. When ordering child SKUs belonging to
  different parents, create a separate spreadsheet for each parent product.


## Converting PO_files to JSON_template

Manually Fill valus in JSON

Locate the purchase‑order (PO) file(s) whose SKU matches each JSON filename (look in order_generation/orders/).

Using the PO, manually fill every field in the JSON template. Do not add, remove, or rename keys; keep the structure exactly the same.

When copying the product‑description text, preserve any bold text, highlights, and font‑color formatting found in the PO.

If you cannot find a matching PO, fill the template with the most reasonable information you can infer and continue; never abort the process.

## Handling Direct SKU Requests

Sometimes an order request only specifies a SKU and quantity, for example:

"I want to place order of 800 of 48-82P3-QSFG product."

Follow these steps to generate the purchase order:

1. Look up the SKU in `order_generation/docs/complete_mapping.json` to
   determine any accessories and their ratios.
2. Copy the JSON templates for the main product and each accessory from
   `order_generation/json_template/` and update the `数量/个` field using the
   requested quantity.
3. Decide whether the items come from one factory or several:
   - **Different factories** – keep the templates separate and create one
     Excel file for each factory.
   - **Same factory** – merge the JSON data first by appending all entries to
     the `products` list and combining the `cells` section by selecting the
     most appropriate value for each cell.
4. Ensure that `order_generation/docs/empty_base_template.xlsx` matches the
   cell addresses used in the JSON templates. If they do not match, update
   `PO_excel.py` accordingly.
5. Run `PO_excel.py` to fill the spreadsheet(s). The script defaults to
   `empty_base_template.xlsx` and writes the values from the JSON file.
6. Review the generated Excel file(s) manually and fix any remaining issues.

This process keeps products from different factories on separate spreadsheets
while still providing a single sheet when everything is sourced from one
factory.
