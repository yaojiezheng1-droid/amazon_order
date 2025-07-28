# Amazon Order Workflow

This repository organizes files used for managing Amazon orders. It is designed for manual updating of spreadsheets with help from AI tools. The main directories are:

- `inventory/` – Inventory information spreadsheets.
- `sales/` – Sales history and reports.
- `orders/` – Past orders and new purchase orders.
- `accessories/` – Packaging and accessory information, including accessory past orders.
- `docs/` – Project notes and other documentation.

Each directory contains an empty `.gitkeep` file so that it is tracked by Git even when no spreadsheets are present.


## Generating Order Templates

Use `generate_order_template.py` with a JSON file that follows the
structure documented in `docs/order_template.md` to create a purchase
order spreadsheet. An example JSON file is provided in
`docs/order_template_example.json`.
