# Order Template Format

The order template is an Excel workbook that already contains most of the layout
and formulas.  Only the yellow cells in `template_order_excel_1.xlsx` are meant
to be filled from JSON.  Green cells contain formulas and will be calculated by
Excel automatically.

`generate_order_template.py` reads a JSON file and writes the values into the
corresponding yellow cells before inserting the product rows.

## JSON Structure

```json
{
  "cells": {
    "B3": {"key": "供货商：", "value": "<supplier name>"},
    "G3": {"key": "订单号", "value": "<order number>"},
    ...
  },
  "products": [
    {
      "产品编号": "EC1601",
      "产品名称": "Sample brush",
      "产品图片": "path/to/image.png",
      "描述": "product description",
      "数量/个": 100,
      "单价": 1.5,
      "包装方式": "pack info"
    },
    ...
  ],
  "footer": {"buyer": "<buyer>", "supplier": "<supplier>"}
}
```

- **cells** – mapping of Excel cell addresses to objects containing ``key`` and
  ``value``. Every yellow cell from the template is included even if the value
  is empty. `template_order_excel_1.xlsx` contains 47 yellow cells and the
  example lists them all.
- **products** – list of product rows inserted starting at row 7. The amount
  column is calculated automatically with a formula.
- **footer** – optional buyer and supplier names written near the bottom of the
  sheet.
  
`order_generation/docs/order_template_example.json` shows a full example.

## Manually Creating JSON from a Purchase Order

When a purchase order (PO) spreadsheet cannot be parsed automatically, you can
fill out the JSON file by hand using the data in that PO:

1. Open the PO `.xlsx` file and locate every field that maps to a yellow cell in
   `template_order_excel_1.xlsx`.
2. Copy the text from each cell directly into the corresponding `value` in the
   JSON. **Do not change the keys** and make sure every field from the template
   appears in the JSON even if the value is empty.
3. Create a product entry for each SKU listed in the PO.
   - Copy the SKU and all related details such as quantity, unit price and
     packaging.
  - If the PO includes an image for the product, save that image in the
    directory `/image/<SKU>/` (create the folder if it does not exist) and set
    the `产品图片` field to that path.  The path is used to place the image in
    the Excel sheet.
  - Include the product name in the `产品名称` field.  When the spreadsheet is
    generated, this name is prepended to the description so the final "描述"
    column contains both.
  - When copying the description, preserve any bold text, highlights and font
    colors exactly as shown in the PO.
4. Complete the `footer` section with the buyer and supplier information from the
   bottom of the PO spreadsheet.

After the JSON is prepared you can run `generate_order_template.py` to populate
`template_order_excel_1.xlsx` with the PO data.
