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
