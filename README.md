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
to write. The script fills those cells in `order_generation/docs/template_order_excel_1.xlsx`
and inserts product rows. Green cells contain formulas that compute totals
automatically.
An example JSON file is provided in `order_generation/docs/order_template_example.json`.


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
   determine any accessories. The entry for `48-82P3-QSFG` lists
   accessories `US-RB01-01`, `SSD`, `ST1122-1-2` and `ST1122-5` with a 1:1
   ratio.
2. Copy the JSON templates for the main product and each accessory from
   `order_generation/json_template/`. Set the `数量/个` field to the
   requested quantity multiplied by the accessory ratio (800 in this
   example).
3. Open `order_generation/docs/empty_base_template.xlsx` and manually copy
   the values from each JSON file into the matching cells. Do **not** run any
   Python scripts to generate the spreadsheet.

This procedure ensures the PO includes the product and all required
accessories for the specified quantity.
