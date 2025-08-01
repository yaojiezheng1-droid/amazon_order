# Amazon Order Workflow

This repository organizes files used for managing Amazon orders. It is designed for manual updating of spreadsheets with help from AI tools. The main directories are:

- `order_generation/docs/` – Project notes, templates, and mapping JSON files.
- `order_generation/orders/` – Past orders and new purchase orders.
- `order_generation/accessories/` – Packaging and accessory information.
- `order_generation/json_exports/` – Generated JSON exports from Excel files.
- `sales/` – Sales history and reports.

Each directory contains an empty `.gitkeep` file so that it is tracked by Git even when no spreadsheets are present.


## Generating Order Templates

Use `generate_order_template.py` with a JSON file that follows the

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

Locate the purchase‑order (PO) file(s) whose SKU matches each JSON filename (look in order_generation/orders/).

Using the PO, manually fill every field in the JSON template. Do not add, remove, or rename keys; keep the structure exactly the same.

When copying the product‑description text, preserve any bold text, highlights, and font‑color formatting found in the PO.

If you cannot find a matching PO, fill the template with the most reasonable information you can infer and continue; never abort the process.

## Handling Direct SKU Requests

Sometimes an order request only specifies a SKU and quantity, for example:

"I want to place order of 800 of 48-82P3-QSFG product."

To ensure an AI agent completes the task without skipping any steps, give it
this prompt and require it to pause after each numbered step before continuing:

> You are generating purchase orders from SKU/quantity requests. Work through
> the numbered checklist one item at a time and only move to the next item once
> the current one is fully complete. Your goal is to populate
> `order_generation/docs/empty_base_template.xlsx` with the correct values from
> the JSON templates and produce a finished spreadsheet.

1. Look up the SKU in `order_generation/docs/complete_mapping.json` to
   determine all required accessories and their ratios.
2. For the main product and each accessory, copy the corresponding JSON
   template from `order_generation/json_template/` and set the `数量/个` field to
   the requested quantity.
3. Decide whether the items originate from one factory or several:
   - **Different factories** – keep the templates separate and create one Excel
     file per factory.
   - **Same factory** – merge the JSON data by appending every entry to the
     `products` list and merging the `cells` section, choosing the most
     appropriate value for each cell.
4. For each factory group, run `json_PO_excel.py` to write the JSON values into
   `order_generation/docs/empty_base_template.xlsx`.
5. Confirm that the populated `empty_base_template.xlsx` matches the cell
   addresses expected by the JSON. Update `generate_order_template.py` if any
   mismatch appears.
6. Execute `generate_order_template.py` to produce the final spreadsheet(s)
   from the filled JSON data.
7. Review the generated Excel file(s) and correct any remaining issues.

This process keeps products from different factories on separate spreadsheets
while still providing a single sheet when everything is sourced from one
factory.
