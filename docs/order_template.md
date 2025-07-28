# Order Template Format

This document describes a JSON structure that can be used with
`generate_order_template.py` to create purchase order spreadsheets in
the same style as the samples in the `orders/` directory.

## JSON Structure

```json
{
  "company_header": ["<line1>", "<line2>", ...],
  "title": "<document title>",
  "info_lines": [
    ["<label1>", "<value1>", ...],
    ...
  ],
  "table": {
    "columns": ["<col1>", "<col2>", ...],
    "items": [
      {
        "<col1>": "value",
        "<col2>": "value",
        "描述": [
          {"text": "part1", "bold": true, "bgcolor": "#FFFF00"},
          {"text": "part2"}
        ]
      },
      ...
    ],
    "total": ["<label>", "", "", "<quantity>", "", "<amount>"]
  },
  "notes": ["note line 1", "note line 2", ...],
  "footer": {"buyer": "<buyer>", "supplier": "<supplier>"}
}
```

- **company_header** – lines appearing at the very top of the sheet with
  company address and contact information.
- **title** – document title such as `\u8BA2\u5355`.
- **info_lines** – rows of supplier and order information. Each entry is
  written to a row.
- **table.columns** – list of column headers.
- **table.items** – list of line items. When the `\u63CF\u8FF0` column is a
  list of runs, each run can specify `bold` or `bgcolor` to preserve rich
  text formatting.
- If columns named `产品图片`, `颜色图片`, or `印刷logo图片` are present, the
  corresponding values should be relative paths to image files in the
  `images/products`, `images/colors`, or `images/logos` directories
  respectively. These images will be inserted into the generated
  spreadsheet.
- **table.total** – optional row summarising quantity and amount.
- **notes** – list of additional lines following the product table.
- **footer** – buyer and supplier names used at the bottom of the sheet.

An example JSON file is provided in `order_template_example.json`.
